"""Tests sans GPU; aucune mesure de précision OCR n'est revendiquée."""
import ast
import io
import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

APP_ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(APP_ROOT))
os.environ["KYC_CORE_CPU_TEST"]="1"
from config import AppConfig
from exceptions import ConcurrentRun, InvalidArchive, ReportUnavailable
from file_manager import RunLock, is_busy, validate_archive, save_upload, sha256_file, atomic_json
from reporting import download_report, publish_receipt
import validated_core as core

class Upload(io.BytesIO):
    name="kyc_documents_test.zip"

def archive(files):
    buf=Upload()
    with zipfile.ZipFile(buf,"w",compression=zipfile.ZIP_STORED) as z:
        for name,data in files:z.writestr(name,data)
    buf.seek(0)
    return buf

class ApplicationTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
        self.settings=AppConfig(input_dir=self.root/"input",extract_dir=self.root/"extracted",
            output_dir=self.root/"output",model_path=self.root/"model",reference_csv=self.root/"reference.csv")
        self.settings.model_path.mkdir()
        (self.settings.model_path/"config.json").write_text('{}')
        self.settings.reference_csv.write_text('Id tiers;Date de naissance;Nom abrege tiers;Date expiration document;Numero de compte\n001;1980-01-01;TEST NAME;2030-01-01;001234\n')

    def tearDown(self):self.tmp.cleanup()

    def test_all_original_definitions_preserved(self):
        notebook=json.loads((APP_ROOT/"reference/validated_notebook.ipynb").read_text())
        actual={}
        for module in ("validated_core.py","notebook_config.py"):
            for node in ast.parse((APP_ROOT/module).read_text()).body:
                if isinstance(node,(ast.FunctionDef,ast.ClassDef)):actual[node.name]=ast.dump(node,include_attributes=False)
        count=0
        for cell in notebook["cells"]:
            if cell["cell_type"]!="code":continue
            for node in ast.parse(''.join(cell["source"])).body:
                if isinstance(node,(ast.FunctionDef,ast.ClassDef)):
                    self.assertEqual(ast.dump(node,include_attributes=False),actual[node.name],node.name);count+=1
        self.assertEqual(count,230)

    def test_original_80_similarity_assertions(self):
        self.assertEqual(core.run_name_similarity_tests()["checks"],80)

    def test_no_model_loaded_on_core_import(self):
        self.assertFalse(core.MODEL_READY)
        self.assertFalse(hasattr(core,"model"))

    def test_model_guard_reuses_and_rejects_changed_repository(self):
        from types import SimpleNamespace
        import model_manager
        from exceptions import EnvironmentNotReady
        fake=SimpleNamespace(MODEL_READY=False,RAW_MODEL_INFO={"model_signature":"one"})
        loads=[];events=[]
        def original():
            if not fake.MODEL_READY:
                loads.append("weights");fake.MODEL_READY=True
        fake.ensure_model_loaded=original
        with patch.object(model_manager,"_POISONED",False),patch.object(model_manager,"_LOADED_REPOSITORY_SIGNATURE",None):
            load=model_manager.guard_model(fake,lambda done=False:events.append(done))
            load();load()
            self.assertEqual(loads,["weights"])
            self.assertEqual(events,[False,True])
            fake.RAW_MODEL_INFO["model_signature"]="two"
            with self.assertRaises(EnvironmentNotReady):load()

    def test_safe_archive(self):
        result=validate_archive(archive([("001/FATCA.PDF",b"data"),("002/note.txt",b"note")]),self.settings,verify_crc=True)
        self.assertEqual(result.customers,2)

    def test_wrapper_folder(self):
        result=validate_archive(archive([("kyc_documents_test/001/FATCA.PDF",b"data")]),self.settings)
        self.assertEqual(result.customers,1)

    def test_unsafe_paths_rejected(self):
        for name in ("../x","/tmp/x","001/../../x","C:\\tmp\\x","001/../x","001\\..\\x"):
            with self.subTest(name=name),self.assertRaises(InvalidArchive):
                validate_archive(archive([(name,b"x")]),self.settings)

    def test_symlink_rejected(self):
        stream=Upload()
        with zipfile.ZipFile(stream,"w") as z:
            info=zipfile.ZipInfo("001/FATCA.PDF");info.create_system=3;info.external_attr=(0o120777<<16)
            z.writestr(info,"/etc/passwd")
        with self.assertRaises(InvalidArchive):validate_archive(stream,self.settings)

    def test_duplicate_rejected(self):
        with self.assertWarns(UserWarning):data=archive([("001/x",b"a"),("001/x",b"b")])
        with self.assertRaises(InvalidArchive):validate_archive(data,self.settings)

    def test_path_collision_rejected(self):
        with self.assertRaises(InvalidArchive):
            validate_archive(archive([("001/x",b"a"),("001/x/y",b"b")]),self.settings)

    def test_crc_rejected(self):
        data=archive([("001/FATCA.PDF",b"unique-payload")]).getvalue()
        damaged=Upload(data.replace(b"unique-payload",b"Unique-payload"))
        with self.assertRaises(InvalidArchive):validate_archive(damaged,self.settings,verify_crc=True)

    def test_empty_and_no_customer_rejected(self):
        for stream in (Upload(b""),archive([]),archive([("FATCA.PDF",b"data")])):
            with self.assertRaises(InvalidArchive):validate_archive(stream,self.settings)

    def test_lock_blocks_threads_and_releases(self):
        with RunLock(self.settings.lock_path,"one"):
            self.assertTrue(is_busy(self.settings.lock_path))
            with self.assertRaises(ConcurrentRun):RunLock(self.settings.lock_path,"two").acquire()
        self.assertFalse(is_busy(self.settings.lock_path))
        # Stale JSON does not block: the kernel lock is authoritative.
        with RunLock(self.settings.lock_path,"three"):pass

    def test_lock_blocks_other_process(self):
        code="from file_manager import RunLock; from exceptions import ConcurrentRun; import sys\ntry: RunLock(sys.argv[1],'other').acquire()\nexcept ConcurrentRun: sys.exit(7)\n"
        with RunLock(self.settings.lock_path,"one"):
            result=subprocess.run([sys.executable,"-c",code,str(self.settings.lock_path)],env={**os.environ,"PYTHONPATH":str(APP_ROOT)},capture_output=True)
        self.assertEqual(result.returncode,7,result.stderr)

    def test_lock_recovers_after_process_exit(self):
        code="from file_manager import RunLock; import os,sys; lock=RunLock(sys.argv[1],'crashed').acquire(); os._exit(3)"
        result=subprocess.run([sys.executable,"-c",code,str(self.settings.lock_path)],env={**os.environ,"PYTHONPATH":str(APP_ROOT)},capture_output=True)
        self.assertEqual(result.returncode,3)
        with RunLock(self.settings.lock_path,"recovered"):pass

    def test_upload_atomic_and_name(self):
        good=archive([("001/note.txt",b"notes")])
        digest=save_upload(good,self.settings)
        self.assertEqual(digest,sha256_file(self.settings.input_path))
        bad=Upload(b"invalid")
        with self.assertRaises(InvalidArchive):save_upload(bad,self.settings)
        self.assertEqual(digest,sha256_file(self.settings.input_path))
        good.name="other.zip"
        with self.assertRaises(InvalidArchive):save_upload(good,self.settings)

    def test_end_to_end_error_pdf_resume_and_fresh_report(self):
        # All real notebook stages run; only hardware/repository inspection is stubbed.
        # One unreadable PDF plus missing documents: zero Qwen calls expected.
        import pipeline
        source=archive([("001/FATCA.PDF",b"This is not a PDF"),("002/note.txt",b"notes")])
        digest=save_upload(source,self.settings)
        profile={"model_path":str(self.settings.model_path),"model_type":"TEST_ONLY","architectures":[],
                 "quantization_config":{"quant_method":"fp8"},"model_signature":"synthetic-signature","signature_method":"test"}
        for i in (1,2):
            run_id=str(i)*32;events=[]
            with RunLock(self.settings.lock_path,run_id),patch.object(pipeline,"get_core",return_value=core),patch.object(pipeline,"require_h100"),patch.object(core,"inspect_local_repository",return_value=profile):
                result=pipeline.execute_locked(self.settings,run_id,digest,events.append)
                self.assertTrue(result.success)
                self.assertEqual(result.customers_processed,2)
                self.assertEqual(result.reference_not_found_count,1)
                self.assertEqual(result.documents_processed,1)
                self.assertEqual(result.qwen_calls,0)
                self.assertTrue(result.warnings)
                atomic_json(self.settings.metadata_dir/(run_id+".json"),{"status":"COMPLETED_WITH_WARNINGS"})
                publish_receipt(self.settings,result,digest)
            payload=download_report(self.settings,run_id,digest)
            self.assertGreater(len(payload),100)
            self.assertTrue(any(e.stage=="Génération Excel" for e in events))
        with self.assertRaises(ReportUnavailable):download_report(self.settings,"1"*32,digest)
        self.settings.report_path.write_bytes(b"modified")
        with self.assertRaises(ReportUnavailable):download_report(self.settings,"2"*32,digest)

    def test_interrupted_metadata_recovery(self):
        from jobs import JobController
        run_id="a"*32
        atomic_json(self.settings.metadata_dir/(run_id+".json"),{"run_id":run_id,"status":"RUNNING"})
        controller=JobController(self.settings)
        self.assertEqual(controller.read(run_id)["status"],"INTERRUPTED")
        self.assertIsNone(controller.read("../../secret"))

    def test_failed_run_cannot_offer_previous_report(self):
        import pipeline
        self.settings.report_path.parent.mkdir(parents=True)
        self.settings.report_path.write_bytes(b"old-report")
        atomic_json(self.settings.report_receipt,{"status":"COMPLETED","run_id":"a"*32})
        self.settings.input_dir.mkdir(parents=True)
        self.settings.input_path.write_bytes(b"not-a-zip")
        run_id="b"*32
        with RunLock(self.settings.lock_path,run_id),self.assertRaises(InvalidArchive):
            pipeline.execute_locked(self.settings,run_id,sha256_file(self.settings.input_path))
        self.assertFalse(self.settings.report_path.exists())
        self.assertEqual(json.loads(self.settings.report_receipt.read_text())["status"],"PENDING")
        self.assertEqual((self.settings.output_dir/"report_history"/run_id/self.settings.report_filename).read_bytes(),b"old-report")

if __name__=="__main__":unittest.main(verbosity=2)
