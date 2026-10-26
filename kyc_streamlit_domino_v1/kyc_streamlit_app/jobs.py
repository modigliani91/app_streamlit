"""Un seul thread de traitement dans le processus Streamlit; état durable JSON."""
import json
import re
import threading
import uuid
from config import APP_VERSION
from exceptions import KYCAppError, ConcurrentRun
from file_manager import RunLock, atomic_json, save_upload, is_busy
from logging_config import configure_logging
from progress import utc_now
from reporting import publish_receipt

class JobController:
    def __init__(self, settings):
        self.settings=settings
        self.logger=configure_logging(settings.output_dir)
        self.logger.info("app_start version=%s",APP_VERSION)
        self._thread=None

    def start(self, upload):
        settings=self.settings
        run_id=uuid.uuid4().hex
        lock=RunLock(settings.lock_path,run_id).acquire()
        try:
            digest=save_upload(upload,settings)
            state={"run_id":run_id,"pid":__import__('os').getpid(),"app_version":APP_VERSION,
                "started_at":utc_now(),"finished_at":None,"status":"RUNNING",
                "input_zip":str(settings.input_path),"input_sha256":digest,
                "report_path":str(settings.report_path),"progress":{"stage":"Validation ZIP","message":"Archive enregistrée; traitement démarré"}}
            atomic_json(settings.metadata_dir/(run_id+".json"),state)
            self.logger.info("run=%s upload_saved",run_id)
            self._thread=threading.Thread(target=self._work,args=(state,lock),name="kyc-single-job",daemon=True)
            self._thread.start()
            return run_id
        except BaseException:
            lock.release();raise

    def _work(self,state,lock):
        settings=self.settings;path=settings.metadata_dir/(state["run_id"]+".json")
        def progress(event):
            state["progress"]=event.to_dict();atomic_json(path,state)
        try:
            from pipeline import execute_locked
            result=execute_locked(settings,state["run_id"],state["input_sha256"],progress)
            state.update(status="COMPLETED_WITH_WARNINGS" if result.warnings else "COMPLETED",
                         result=result.to_dict(),finished_at=utc_now())
            atomic_json(path,state)
            publish_receipt(settings,result,state["input_sha256"])
            self.logger.info("run=%s completed status=%s",state["run_id"],state["status"])
        except BaseException as exc:
            self.logger.exception("run=%s pipeline_failure",state["run_id"])
            state.update(status="FAILED",finished_at=utc_now(),
                error=str(exc) if isinstance(exc,KYCAppError) else "Le traitement a été interrompu. Consultez les journaux techniques puis relancez la même archive pour reprendre.",
                failure_type=type(exc).__name__)
            try:atomic_json(path,state)
            except OSError:self.logger.exception("run=%s metadata_write_failed",state["run_id"])
        finally:lock.release()

    def read(self,run_id):
        if not re.fullmatch(r"[0-9a-f]{32}",str(run_id)):return None
        path=self.settings.metadata_dir/(run_id+".json")
        try:state=json.loads(path.read_text(encoding="utf-8"))
        except (OSError,ValueError):return None
        if state.get("status")=="RUNNING" and not is_busy(self.settings.lock_path):
            # Re-read under lock: never mark an active job stale based only on age/PID.
            try:
                with RunLock(self.settings.lock_path,"recover-"+run_id):
                    state=json.loads(path.read_text(encoding="utf-8"))
                    if state.get("status")=="RUNNING":
                        state.update(status="INTERRUPTED",finished_at=utc_now(),
                            error="Le processus a été arrêté. Rechargez la même archive et relancez le traitement pour reprendre les checkpoints.")
                        atomic_json(path,state)
            except ConcurrentRun:pass
        return state
