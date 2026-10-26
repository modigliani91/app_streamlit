"""Propriétaire unique du moteur et, à travers lui, de Qwen/processor."""
import importlib
import os
import threading
from importlib import metadata
from exceptions import EnvironmentNotReady

_CORE = None
_IMPORT_LOCK = threading.Lock()
_POISONED = False
_LOADED_REPOSITORY_SIGNATURE = None

def get_core():
    global _CORE
    with _IMPORT_LOCK:
        if _POISONED:
            raise EnvironmentNotReady("Le contexte du modèle est inutilisable. Redémarrez l'App avant de reprendre.")
        if _CORE is None:
            if os.environ.get("KYC_CORE_CPU_TEST") == "1":
                raise EnvironmentNotReady("Le mode de test CPU interdit le traitement réel.")
            os.environ.update(HF_HUB_OFFLINE="1", TRANSFORMERS_OFFLINE="1", HF_HUB_DISABLE_TELEMETRY="1", DO_NOT_TRACK="1")
            try:
                _CORE = importlib.import_module("validated_core")
                packages = ["numpy","pandas","Pillow","PyMuPDF","openpyxl","packaging","torch","transformers","accelerate","psutil"]
                _CORE.INSTALLED_VERSIONS = {p:metadata.version(p) for p in packages}
            except Exception as exc:
                _CORE = None
                raise EnvironmentNotReady("Le moteur KYC ne peut pas être importé. Vérifiez l'environnement du notebook et les journaux.") from exc
        return _CORE

def require_h100(core):
    """Validation du matériel uniquement; aucun poids chargé ici."""
    torch=core.torch
    if torch is None or not torch.cuda.is_available():
        raise EnvironmentNotReady("Le GPU requis pour le traitement n'est pas disponible (CUDA indisponible).")
    if torch.cuda.device_count()!=1 or "H100" not in torch.cuda.get_device_name(0):
        raise EnvironmentNotReady("Un seul GPU NVIDIA H100 doit être visible par cette application.")
    if not torch.cuda.is_bf16_supported():
        raise EnvironmentNotReady("Le GPU visible ne prend pas en charge BF16.")

def guard_model(core, callback):
    """Décorateur temporaire autour du chargeur original, sans second modèle."""
    original=core.ensure_model_loaded
    def load():
        global _POISONED, _LOADED_REPOSITORY_SIGNATURE
        if _POISONED:
            raise EnvironmentNotReady("Redémarrez l'App pour réinitialiser le modèle.")
        was_ready=core.MODEL_READY
        if not was_ready:
            callback()
        elif _LOADED_REPOSITORY_SIGNATURE != core.RAW_MODEL_INFO.get("model_signature"):
            raise EnvironmentNotReady("Le dépôt du modèle a changé depuis son chargement. Redémarrez l'App.")
        try:
            original()
            _LOADED_REPOSITORY_SIGNATURE=core.RAW_MODEL_INFO.get("model_signature")
            if not was_ready:callback(True)
        except BaseException:
            # A partial load or failed validation can leave large tensors alive.
            # Never retry a second load in the same process.
            if not core.MODEL_READY: _POISONED=True
            raise
    return load

def finish_run(core):
    global _POISONED
    if core.RUNTIME_STATE.get("fatal_cuda"): _POISONED=True
    # Retain only model/processor and tiny diagnostics between jobs.
    for key in ("REFERENCE","ALL_CUSTOMERS","INVENTORY","SELECTED_DOCUMENTS","CUSTOMERS_SELECTED",
                "SELECTED_CUSTOMER_IDS","SELECTIONS_BY_CUSTOMER","CUSTOMER_RESULTS","CUSTOMER_SNAPSHOTS",
                "ATTEMPTS","PERFORMANCE","ERRORS"):
        value=getattr(core,key,None)
        if hasattr(value,"clear"): value.clear()
    core.REFERENCE=None
    core.STRUCTURE_DOCUMENT_CONTEXT.clear()
    core.cleanup_document()
