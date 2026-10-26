"""Contrôle sans charger Qwen ni appeler from_pretrained."""
import argparse
import importlib
from importlib import metadata
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from config import SETTINGS
from file_manager import readiness

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--paths-only",action="store_true")
    args=parser.parse_args()
    failures=[]
    def check(label,condition,detail=""):
        print(("PASS" if condition else "FAIL"),label,detail)
        if not condition:failures.append(label)
    check("Python >= 3.10",sys.version_info>=(3,10),sys.executable+" "+sys.version.split()[0])
    errors=readiness(SETTINGS)
    for error in errors:check(error,False)
    if not errors:check("CSV / modèle / entrées / sorties accessibles",True)
    if not args.paths_only:
        minima={"streamlit":"1.44","numpy":"1.23","pandas":"1.5","Pillow":"9.0","PyMuPDF":"1.23",
                "openpyxl":"3.1","packaging":"21.0","torch":"2.0","transformers":"4.57","accelerate":"0.30","psutil":"5.9"}
        try:
            from packaging.version import Version
            for package,minimum in minima.items():
                try:
                    version=metadata.version(package)
                    check(package,Version(version)>=Version(minimum),version)
                except metadata.PackageNotFoundError:check(package,False,"absent")
        except ImportError:check("packaging",False,"absent")
        try:
            import torch
            check("CUDA",torch.cuda.is_available(),str(torch.version.cuda))
            names=[torch.cuda.get_device_name(i) for i in range(torch.cuda.device_count())]
            check("Un seul H100 visible",len(names)==1 and "H100" in names[0],str(names))
            check("BF16",torch.cuda.is_available() and torch.cuda.is_bf16_supported())
            if names:print("GPU mémoire totale GiB",torch.cuda.get_device_properties(0).total_memory/1024**3)
        except Exception as exc:check("PyTorch / pilote CUDA",False,str(exc))
        try:
            from transformers import AutoProcessor, AutoModelForImageTextToText
            try:from transformers.integrations.finegrained_fp8 import FineGrainedFP8Config
            except ImportError:from transformers import FineGrainedFP8Config
            check("FineGrainedFP8Config(dequantize=True)",FineGrainedFP8Config(dequantize=True).dequantize is True)
        except Exception as exc:check("Classes Transformers du notebook",False,str(exc))
        try:
            import fitz
            check("PyMuPDF",hasattr(fitz,"Matrix") and hasattr(fitz,"open"),getattr(fitz,"VersionBind",""))
        except Exception as exc:check("PyMuPDF",False,str(exc))
        try:
            import streamlit as st
            check("API Streamlit",all(hasattr(st,n) for n in ("fragment","status","download_button","cache_resource")))
        except Exception as exc:check("Streamlit",False,str(exc))
    print("Aucun poids modèle chargé. La compatibilité complète reste validée au benchmark.")
    return bool(failures)

if __name__=="__main__":sys.exit(main())
