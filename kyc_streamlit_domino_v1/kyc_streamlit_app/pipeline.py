"""Orchestration du moteur original. Aucun composant Streamlit, aucun nbconvert."""
from dataclasses import replace
from pathlib import Path
from contextlib import contextmanager
from collections import defaultdict
import logging
import time
import uuid

from config import SETTINGS
from exceptions import EnvironmentNotReady, InvalidArchive
from file_manager import RunLock, sha256_file, validate_archive, readiness, atomic_json
from logging_config import configure_logging
from model_manager import get_core, require_h100, guard_model, finish_run
from progress import ProgressEvent, utc_now
from reporting import PipelineResult, invalidate_report, validate_report, publish_receipt

@contextmanager
def instrument(core, emit):
    """Callbacks autour de fonctions originales; aucun changement de retour ou de règle."""
    originals={}
    counters=defaultdict(int)
    def wrap(name, stage, message, total, count_documents=False):
        original=getattr(core,name); originals[name]=original
        def call(*args,**kwargs):
            result=original(*args,**kwargs)
            if not count_documents or args[0].get("document_found"):
                counters[name]+=1
            emit(ProgressEvent(stage,message,counters[name],total))
            return result
        setattr(core,name,call)
    total_docs=sum(s["document_found"] for s in core.SELECTED_DOCUMENTS)
    wrap("run_raw_ocr_document","RAW OCR","Documents OCR terminés",total_docs,True)
    wrap("process_structured_customer","Extraction structurée","Clients structurés",len(core.CUSTOMERS_SELECTED))
    wrap("process_customer","Matching base / Contrôles","Clients rapprochés",len(core.CUSTOMERS_SELECTED))
    originals["raw_for_page"]=core.raw_for_page
    page_original=core.raw_for_page
    page_total=sum(s.get("page_count") or 0 for s in core.SELECTED_DOCUMENTS)
    def raw_page(*args,**kwargs):
        emit(ProgressEvent("RAW OCR","Lecture des pages",counters["pages"],page_total))
        result=page_original(*args,**kwargs)
        counters["pages"]+=1
        emit(ProgressEvent("RAW OCR","Pages OCR terminées",counters["pages"],page_total))
        return result
    core.raw_for_page=raw_page
    originals["ensure_model_loaded"]=core.ensure_model_loaded
    core.ensure_model_loaded=guard_model(core,lambda loaded=False:emit(ProgressEvent(
        "Préparation des inférences" if loaded else "Chargement du modèle",
        "Modèle disponible sur le H100" if loaded else "Chargement initial de Qwen sur le H100")))
    originals["load_stage_cache"]=core.load_stage_cache
    cached_loader=core.load_stage_cache
    # Stage B rebuilds document projections; the Python matching must be reapplied to them.
    # RAW, structured, benchmark caches remain intact. No repeat inference for matching.
    core.load_stage_cache=lambda stage,key: None if stage=="matching" else cached_loader(stage,key)
    try: yield
    finally:
        for name,original in originals.items(): setattr(core,name,original)

def preserve_previous_projections(core, run_id):
    """Archive previous materialized views; caches/proofs remain untouched.

    A failed customer in this run cannot silently appear through an older JSON.
    """
    for customer_id in core.CUSTOMERS_SELECTED:
        directory=core.customer_json_path(customer_id).parent
        if directory.exists():
            target=core.CFG.OUTPUT_DIR/"projection_history"/run_id/directory.name
            target.parent.mkdir(parents=True,exist_ok=True)
            directory.rename(target)

def execute_locked(settings, run_id, input_sha256, progress_callback=None):
    """Private: caller owns RunLock continuously from input publication to receipt."""
    started=time.perf_counter(); started_ns=time.time_ns()
    logger=configure_logging(settings.output_dir)
    core=None
    def emit(event):
        logger.info("run=%s stage=%s current=%s total=%s",run_id,event.stage,event.current,event.total)
        if progress_callback: progress_callback(event)
    try:
        errors=readiness(settings)
        if errors: raise EnvironmentNotReady("\n".join(errors))
        invalidate_report(settings,run_id)
        emit(ProgressEvent("Validation ZIP","Vérification intégrale de l'archive"))
        validate_archive(settings.input_path,settings,verify_crc=True)
        if sha256_file(settings.input_path)!=input_sha256:
            raise InvalidArchive("L'archive a changé avant le traitement.")
        core=get_core()
        config=settings.notebook_config(input_sha256)
        core.initialize_run(config,run_id)
        require_h100(core)
        core.DIRS["reports"]=settings.report_path.parent
        core.DIRS["reports"].mkdir(parents=True,exist_ok=True)
        emit(ProgressEvent("Extraction ZIP","Extraction sûre des PDF cibles"))
        core.ALL_CUSTOMERS,core.INVENTORY,core.SELECTED_DOCUMENTS=core.prepare_inventory()
        if not core.ALL_CUSTOMERS: raise InvalidArchive("Aucun dossier client découvert.")
        core.CUSTOMERS_SELECTED=core.select_customers(core.ALL_CUSTOMERS)
        core.assert_full_population(core.ALL_CUSTOMERS,core.CUSTOMERS_SELECTED)
        core.SELECTIONS_BY_CUSTOMER=defaultdict(list)
        for selection in core.SELECTED_DOCUMENTS:
            core.SELECTIONS_BY_CUSTOMER[selection["customer_id"]].append(selection)
        core.OUTPUT_POPULATION=set(core.ALL_CUSTOMERS)
        core.SELECTED_CUSTOMER_IDS=core.CUSTOMERS_SELECTED
        emit(ProgressEvent("Inventaire clients","Clients recensés",len(core.ALL_CUSTOMERS),len(core.ALL_CUSTOMERS)))
        n_docs=sum(s["document_found"] for s in core.SELECTED_DOCUMENTS)
        emit(ProgressEvent("Inventaire documents","PDF sélectionnés selon les règles du notebook",n_docs,n_docs))
        with instrument(core,emit):
            core.load_raw_model_info()
            if core.CFG.RUN_PREFLIGHT_BENCHMARK:
                emit(ProgressEvent("RAW OCR","Validation initiale froid/chaud ou reprise du benchmark"))
                core.BENCHMARK_REPORT=core.run_raw_preflight()
            core.RAW_RUN_SUMMARY=core.process_raw_dataset(core.CUSTOMERS_SELECTED)
            core.initialize_structure_stage()
            preserve_previous_projections(core,run_id)
            emit(ProgressEvent("Extraction structurée","Extraction depuis les preuves OCR sauvegardées",0,len(core.CUSTOMERS_SELECTED)))
            core.STRUCTURE_RUN_SUMMARY=core.process_structured_dataset(core.CUSTOMERS_SELECTED)
            # Reference contents enter memory only AFTER extraction, as in cell 98.
            emit(ProgressEvent("Matching base","Lecture de la base de référence après extraction"))
            core.REFERENCE=core.load_reference_csv(core.CFG.REFERENCE_CSV)
            emit(ProgressEvent("Contrôles","Rapprochement documentaire et consolidation clients",0,len(core.CUSTOMERS_SELECTED)))
            core.MATCH_RUN_SUMMARY=core.process_all_customers(core.CUSTOMERS_SELECTED,core.REFERENCE)
            emit(ProgressEvent("Génération Excel","Création du rapport du pipeline validé"))
            core.REPORT_SUMMARY=core.export_reports()
            core.AUDIT=core.final_audit()
        digest=validate_report(settings.report_path,started_ns,len(core.CUSTOMERS_SELECTED))
        counts=core.MATCH_RUN_SUMMARY["counts"]
        manual=sum(core.read_customer_result(cid)["manual_review"]["required"] for cid in core.CUSTOMERS_SELECTED)
        result=PipelineResult(success=True,run_id=run_id,report_path=str(settings.report_path),
            customers_discovered=len(core.ALL_CUSTOMERS),customers_processed=sum(counts.values()),
            customers_failed=counts.get("PROCESSING_ERROR",0), documents_processed=core.REPORT_SUMMARY["rows"]["documents"],
            conform_count=counts.get("CONFORME_A_LA_BASE",0),non_conform_count=counts.get("NON_CONFORME_A_LA_BASE",0),
            recheck_count=counts.get("A_RECONTROLER",0),reference_not_found_count=counts.get("REFERENCE_NOT_FOUND",0),
            manual_review_count=manual, errors=[f"{e.get('stage','pipeline')}: {e.get('failure_type','erreur')}" for e in core.ERRORS],
            elapsed_seconds=time.perf_counter()-started,qwen_calls=core.BUDGET.used,
            pipeline_version=core.CFG.PIPELINE_VERSION,matching_version=core.CFG.MATCHING_VERSION,
            report_sha256=digest)
        result.warnings=bool(result.errors or result.customers_failed or core.STRUCTURE_RUN_SUMMARY.get("customers_failed")
            or core.RAW_RUN_SUMMARY.get("summary",{}).get("pages_failed"))
        # A result is successful when the audited report is generated, even with per-document anomalies.
        logger.info("run=%s report_generated customers=%s documents=%s",run_id,result.customers_processed,result.documents_processed)
        return result
    finally:
        if core is not None:
            try: finish_run(core)
            except Exception: logger.exception("run=%s cleanup_failure",run_id)

def run_kyc_pipeline(zip_path: Path, output_dir: Path=SETTINGS.output_dir,
                     reference_csv: Path=SETTINGS.reference_csv, progress_callback=None) -> PipelineResult:
    """API synchrone/CLI sûre. Le contrôleur UI utilise la même exécution sous verrou."""
    from file_manager import save_upload
    settings=replace(SETTINGS,output_dir=Path(output_dir),reference_csv=Path(reference_csv))
    run_id=uuid.uuid4().hex;metadata_path=settings.metadata_dir/(run_id+".json")
    with RunLock(settings.lock_path,run_id):
        state={"run_id":run_id,"status":"RUNNING","started_at":utc_now(),"input_zip":str(settings.input_path)}
        atomic_json(metadata_path,state)
        try:
            if Path(zip_path).resolve()!=settings.input_path.resolve():
                class FileUpload:
                    name=settings.input_filename
                    def __init__(self,stream): self.stream=stream
                    def read(self,n): return self.stream.read(n)
                    def seek(self,n): return self.stream.seek(n)
                with Path(zip_path).open("rb") as stream: digest=save_upload(FileUpload(stream),settings)
            else: digest=sha256_file(settings.input_path)
            state["input_sha256"]=digest; atomic_json(metadata_path,state)
            result=execute_locked(settings,run_id,digest,progress_callback)
            state.update(status="COMPLETED_WITH_WARNINGS" if result.warnings else "COMPLETED",result=result.to_dict(),finished_at=utc_now())
            atomic_json(metadata_path,state); publish_receipt(settings,result,digest)
            return result
        except BaseException:
            state.update(status="FAILED",finished_at=utc_now());atomic_json(metadata_path,state)
            raise

if __name__=="__main__":
    import argparse
    parser=argparse.ArgumentParser()
    parser.add_argument("--zip",type=Path,required=True)
    args=parser.parse_args()
    result=run_kyc_pipeline(args.zip,progress_callback=lambda e:print(e.stage,e.current,e.total,flush=True))
    print(result.to_dict())
