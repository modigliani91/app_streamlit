"""Interface française; aucune inférence déclenchée par l'import ou un rerun."""
import hashlib
import streamlit as st
from config import SETTINGS, APP_VERSION
from exceptions import KYCAppError
from file_manager import readiness, validate_archive, is_busy
from jobs import JobController
from reporting import download_report

st.set_page_config(page_title="Recertification KYC",page_icon="📋",layout="wide")
st.title("Recertification KYC — Contrôle documentaire assisté par IA")
st.caption("Analyse automatisée des documents KYC et rapprochement avec la base de référence.")

@st.cache_resource(show_spinner=False)
def get_controller():
    # Cache only the controller; model_manager retains the single core/model.
    return JobController(SETTINGS)

errors=readiness(SETTINGS)
if errors:
    for error in errors:st.error(error)
    st.stop()
controller=get_controller()
st.success("Application prête")
st.caption("Les contrôles du GPU et du moteur sont effectués au lancement du traitement.")
st.subheader("1. Charger les documents KYC")
upload=st.file_uploader("kyc_documents_test.zip",type=["zip"],accept_multiple_files=False,
                        help=f"Taille maximale : {SETTINGS.max_upload_mb} Mo. Un dossier par identifiant client.")
valid=False
if upload is not None:
    st.write(f"**Nom du fichier :** {upload.name} — **Taille :** {upload.size / 1024**2:.1f} Mo")
    try:
        if upload.name != SETTINGS.input_filename:
            st.error("Le fichier doit être nommé kyc_documents_test.zip.")
        elif upload.size > SETTINGS.max_upload_mb*1024**2:
            st.error("L'archive dépasse la taille maximale configurée.")
        else:
            digest=hashlib.sha256(upload.getbuffer()).hexdigest()
            if st.session_state.get("validated_upload")!=digest:
                info=validate_archive(upload,SETTINGS)
                st.session_state.validated_upload=digest
                st.session_state.upload_customers=info.customers
            if st.session_state.get("selected_upload")!=digest:
                st.session_state.selected_upload=digest
                # An old report is never presented as the result of a newly selected ZIP.
                old=controller.read(st.session_state.get("run_id") or st.query_params.get("run",""))
                if old and old.get("input_sha256")!=digest:
                    st.session_state.pop("run_id",None)
                    if "run" in st.query_params:del st.query_params["run"]
            st.success(f"Archive valide — {st.session_state.upload_customers} dossier(s) client(s).")
            st.caption("Le contrôle CRC intégral est effectué avant l'extraction.")
            valid=True
    except KYCAppError as exc:st.error(str(exc))

st.subheader("2. Lancer l'analyse")
busy=is_busy(SETTINGS.lock_path)
if busy:st.info("Un traitement est déjà en cours. Le serveur accepte un seul traitement à la fois.")
if st.button("Lancer le traitement",type="primary",disabled=not valid or busy):
    try:
        run_id=controller.start(upload)
        st.session_state.run_id=run_id
        st.query_params["run"]=run_id
    except KYCAppError as exc:st.error(str(exc))
    except Exception:
        controller.logger.exception("upload_or_job_start_failure")
        st.error("Impossible de démarrer le traitement. Vérifiez les permissions et les journaux.")

visible_run_id=st.session_state.get("run_id") or st.query_params.get("run")
visible_state=controller.read(visible_run_id) if visible_run_id else None
poll_active=bool(visible_state and visible_state.get("status")=="RUNNING")

@st.fragment(run_every=2 if poll_active else None)
def show_job():
    run_id=st.session_state.get("run_id") or st.query_params.get("run")
    st.subheader("3. Progression")
    if not run_id:
        st.info("Chargez une archive puis lancez le traitement.")
        return
    state=controller.read(run_id)
    if not state:
        st.warning("Cette exécution n'est pas disponible.")
        return
    current=state.get("progress",{})
    running=state["status"]=="RUNNING"
    completed=state["status"] in {"COMPLETED","COMPLETED_WITH_WARNINGS"}
    if poll_active and not running:
        st.rerun()  # Stop polling and refresh the start button once the job ends.
    with st.status(current.get("stage","Traitement"),expanded=True,
                   state="running" if running else "complete" if completed else "error"):
        st.write(current.get("message",""))
        if current.get("total"):
            st.progress(min(1.,(current.get("current") or 0)/current["total"]),
                text=f"{current.get('message','Progression')} : {current.get('current',0)} / {current['total']}")
    st.caption(f"Exécution : {run_id} · début : {state['started_at']}")
    if running:
        st.caption("Le traitement continue sur le serveur si vous actualisez cette page. Conservez cette URL pour le retrouver.")
        return
    if not completed:
        st.error(state.get("error","Traitement interrompu."))
        return
    result=state["result"]
    if state["status"]=="COMPLETED_WITH_WARNINGS":st.warning("Traitement terminé avec anomalies techniques — consultez le rapport.")
    else:st.success("Traitement terminé.")
    metrics=[("Clients traités",result["customers_processed"]),
             ("Conformes à la base",result["conform_count"]),
             ("Non conformes à la base",result["non_conform_count"]),
             ("À recontrôler",result["recheck_count"]),
             ("Absents de la base",result["reference_not_found_count"]),
             ("Documents traités",result["documents_processed"])]
    for col,(label,value) in zip(st.columns(3),metrics[:3]):col.metric(label,value)
    for col,(label,value) in zip(st.columns(3),metrics[3:]):col.metric(label,value)
    st.caption(f"Relecture humaine signalée : {result['manual_review_count']} dossier(s). Erreurs client : {result['customers_failed']}.")
    st.caption("La conformité à la base et les contrôles documentaires restent distincts, comme dans le notebook. Les relectures peuvent recouper les autres compteurs.")
    with st.expander("Détails techniques"):
        st.json({"app_version":APP_VERSION,"pipeline_version":result["pipeline_version"],"matching_version":result["matching_version"],
                 "modèle":str(SETTINGS.model_path),"GPU":"NVIDIA H100 (contrôlé au lancement)","run_id":run_id,
                 "durée_secondes":round(result["elapsed_seconds"],1),"appels_Qwen":result["qwen_calls"],
                 "répertoire_sorties":str(SETTINGS.output_dir),"erreurs":result["errors"]})
    st.subheader("4. Télécharger le rapport")
    try:
        # Polling has stopped; bytes live in the download widget, never in session state.
        payload=download_report(SETTINGS,run_id,state["input_sha256"])
        st.success("Rapport généré avec succès.")
        st.download_button("Télécharger kyc_recertification_control.xlsx",data=payload,
            file_name=SETTINGS.report_filename,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="report-"+run_id,on_click="ignore")
    except KYCAppError as exc:st.info(str(exc))

show_job()
