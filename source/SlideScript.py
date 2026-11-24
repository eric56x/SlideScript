#!/usr/bin/env python3
# -*- coding: utf-8 -*-
__app_name__ = "SlideScript"
__version__ = "1.1.0"

import os, re, sys, time, base64, urllib.request, urllib.error, subprocess, multiprocessing
from datetime import datetime
from pathlib import Path

if __name__ == "__main__":
    multiprocessing.freeze_support()

if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    _THIS_DIR = Path(__file__).resolve().parent
    SCRIPT_DIR = _THIS_DIR.parent if _THIS_DIR.name.lower() == "source" else _THIS_DIR
    RUNTIME_PYW = SCRIPT_DIR / "runtime" / "pythonw.exe"
    RUNTIME_PY = SCRIPT_DIR / "runtime" / "python.exe"
    current_exe = Path(sys.executable).resolve()
    in_portable_rt = current_exe in {p.resolve() for p in (RUNTIME_PYW, RUNTIME_PY) if p.exists()}
    if not in_portable_rt and (RUNTIME_PYW.exists() or RUNTIME_PY.exists()):
        target_exe = RUNTIME_PYW if RUNTIME_PYW.exists() else RUNTIME_PY
        subprocess.Popen([str(target_exe), str(Path(__file__).resolve()), *sys.argv[1:]])
        sys.exit(0)

_WIN_CA_PEM = SCRIPT_DIR / "source" / "win_ca_bundle.pem"
if not _WIN_CA_PEM.exists():
    _WIN_CA_PEM = SCRIPT_DIR / "win_ca_bundle.pem"
if _WIN_CA_PEM.exists():
    os.environ["REQUESTS_CA_BUNDLE"] = str(_WIN_CA_PEM)
    os.environ["SSL_CERT_FILE"] = str(_WIN_CA_PEM)

import io
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

sys.path = [p for p in sys.path if p not in ("", str(SCRIPT_DIR))]

import json
import shutil
import ctypes
import tempfile
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

if os.name == "nt":
    import winreg

# Support du mode portable (.\runtime\, .\ffmpeg\, .\hf_cache\ à côté du script/.exe)
RUNTIME_DIR = SCRIPT_DIR / "runtime"
FFMPEG_DIR = SCRIPT_DIR / "ffmpeg"
HF_CACHE_DIR = SCRIPT_DIR / "hf_cache"
HF_CACHE_DIR.mkdir(parents=True, exist_ok=True)

if FFMPEG_DIR.exists():
    os.environ["PATH"] = str(FFMPEG_DIR) + os.pathsep + os.environ.get("PATH", "")
os.environ["HF_HOME"] = str(HF_CACHE_DIR)
os.environ["HF_HUB_CACHE"] = str(HF_CACHE_DIR)
os.environ["HF_HUB_DISABLE_SYMLINKS"] = "1"
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
import shutil as _sh_sym
if os.name == "nt":
    def _win_safe_symlink(src, dst, target_is_directory=False, *, dir_fd=None):
        src_p = Path(dst).parent / src if not os.path.isabs(src) else Path(src)
        if src_p.is_dir():
            _sh_sym.copytree(src_p, dst, dirs_exist_ok=True)
        elif src_p.exists():
            _sh_sym.copy2(src_p, dst)
    os.symlink = _win_safe_symlink
os.environ["TORCH_HOME"] = str(HF_CACHE_DIR / "torch")
if (RUNTIME_DIR / "Lib" / "site-packages").exists():
    sys.path.insert(0, str(RUNTIME_DIR / "Lib" / "site-packages"))
    
try:
    from docx import Document
    from docx.shared import RGBColor, Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches as PptxInches
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from PIL import Image, ImageChops, ImageDraw, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

EXTENSIONS_MEDIA = {".mp4", ".mkv", ".mov", ".avi", ".mp3", ".wav", ".m4a", ".flac", ".ogg"}
CONFIG_FILE = SCRIPT_DIR / "slidescript_config.json"
FONT_NAME = "Aptos"
APP_VERSION = f"v{__version__}"
ENGINE_BACKEND = "whisperx"

# Charte graphique unifiée (Thème Bleu & Barre supérieure sombre)
THEME_TOP_BG = "#1E2229"
THEME_BLUE = "#0078B8"
THEME_BLUE_HOVER = "#006399"
THEME_BLUE_DARK = "#005A8D"
THEME_TROUGH_BG = "#DCE1E7"
THEME_BTN_PRIMARY_BG = "#525864"
THEME_BTN_PRIMARY_ACTIVE = "#3E434D"
THEME_BTN_SEC_BG = "#D8DCE2"
THEME_BTN_SEC_ACTIVE = "#C6CBD3"

REQUIRED_HF_REPOS = [
    "pyannote/speaker-diarization-community-1",
    "pyannote/segmentation-3.0",
]

WHISPER_MODEL_CATALOG = [
    {"id": "tiny",        "category": "whisper",  "label_fr": "tiny (~75 Mo)",                                                                             "label_en": "tiny (~75 MB)",                                                                           "repos": ["Systran/faster-whisper-tiny"],                                                      "gated": False, "available": True},
    {"id": "tiny.en",     "category": "whisper",  "label_fr": "tiny.en (~75 Mo)",                                                                          "label_en": "tiny.en (~75 MB)",                                                                        "repos": ["Systran/faster-whisper-tiny.en"],                                                   "gated": False, "available": True},
    {"id": "small",       "category": "whisper",  "label_fr": "small (~464 Mo)",                                                                           "label_en": "small (~464 MB)",                                                                         "repos": ["Systran/faster-whisper-small"],                                                     "gated": False, "available": True},
    {"id": "small.en",    "category": "whisper",  "label_fr": "small.en (~464 Mo)",                                                                        "label_en": "small.en (~464 MB)",                                                                      "repos": ["Systran/faster-whisper-small.en"],                                                  "gated": False, "available": True},
    {"id": "medium",      "category": "whisper",  "label_fr": "medium (~1,46 Go)",                                                                         "label_en": "medium (~1.46 GB)",                                                                       "repos": ["Systran/faster-whisper-medium"],                                                    "gated": False, "available": True},
    {"id": "medium.en",   "category": "whisper",  "label_fr": "medium.en (~1,46 Go)",                                                                      "label_en": "medium.en (~1.46 GB)",                                                                    "repos": ["Systran/faster-whisper-medium.en"],                                                 "gated": False, "available": True},
    {"id": "large-v3",    "category": "whisper",  "label_fr": "large (~2,90 Go — large-v3)",                                                               "label_en": "large (~2.90 GB — large-v3)",                                                             "repos": ["Systran/faster-whisper-large-v3"],                                                  "gated": False, "available": True},
    {"id": "large.en",    "category": "whisper",  "label_fr": "large.en (non publié par OpenAI)",                                                          "label_en": "large.en (not released by OpenAI)",                                                       "repos": [],                                                                                   "gated": False, "available": False},
    {"id": "pyannote",    "category": "speaker",  "label_fr": "Identification des locuteurs (pyannote)", "label_en": "Speaker identification (pyannote)", "repos": ["pyannote/speaker-diarization-community-1", "pyannote/segmentation-3.0"],            "gated": True,  "available": True},
    {"id": "wav2vec2-fr", "category": "align",    "label_fr": "Alignement audio français — rapide",       "label_en": "French audio alignment — fast",       "repos": ["bofenghuang/asr-wav2vec2-ctc-french"],                                     "gated": False, "available": True},
]

SPEAKER_COLORS = [
    RGBColor(211, 47, 47) if DOCX_AVAILABLE else None,
    RGBColor(25, 118, 210) if DOCX_AVAILABLE else None,
    RGBColor(56, 142, 60) if DOCX_AVAILABLE else None,
    RGBColor(123, 31, 162) if DOCX_AVAILABLE else None,
    RGBColor(245, 124, 0) if DOCX_AVAILABLE else None,
    RGBColor(0, 131, 143) if DOCX_AVAILABLE else None,
]

TRANSCRIPT_TS_RE = re.compile(r"Transcript:\s*\[\s*(\d+(?:\.\d+)?)\s*-->\s*(\d+(?:\.\d+)?)\s*\]")
WHISPERX_PROG_RE = re.compile(r"^Progress:\s*(\d+(?:\.\d+)?)%\.\.\.$")
TQDM_DL_RE = re.compile(
    r"^\s*(\d+)%\|[^|]*\|\s*([0-9.]+[kMGT]?B?)/([0-9.]+[kMGT]?B?)\s*\[[^\],]+,\s*([^\]]+)\]"
)

I18N = {
    "fr": {
        "win_title": "SlideScript v1.1.0 — Transcription, Diarisation & Extraction de Slides",
        "ui_lang_label": "Interface :", "app_ver": "v1.1.0",
        "btn_models_mgr": "⚙ Modèles installés & Mises à jour",
        "dlg_mod_title": "Modèles installés & Mises à jour (cache local)",
        "dlg_mod_desc": "Cochez les modèles à installer localement ou décochez ceux à supprimer pour libérer de l'espace disque :",
        "dlg_mod_apply": "Appliquer & Redémarrer",
        "dlg_mod_cancel": "Annuler",
        "dlg_mod_none": "Aucune modification sélectionnée.",
        "dlg_mod_del": "Suppression de {name}...",
        "dlg_mod_dl": "Téléchargement de {name} en cours (connexion Internet requise)...",
        "dlg_mod_restart": "Opérations terminées. Redémarrage de l'application...",
        "sec1": "1. Sélection des médias à transcrire",
        "btn_1file": "Choisir un fichier",
        "btn_mfiles": "Choisir plusieurs fichiers",
        "btn_folder": "Choisir un dossier",
        "chk_sub": "Inclure les sous-dossiers",
        "no_sel": "Aucune sélection pour le moment.",
        "sel_1": "1 fichier sélectionné : {name}",
        "sel_n": "{n} fichiers sélectionnés dans {parent}",
        "sel_dir": "Dossier : {folder} ({n} fichier(s) média trouvé(s))",
        "sec2": "2. Paramètres de transcription",
        "lbl_fmt": "Formats de sortie :",
        "fmt_txt": ".txt",
        "fmt_docx": ".docx",
        "fmt_docx_slides": ".docx avec slides",
        "fmt_pptx": ".pptx avec transcripts dans les notes",
        "lbl_lang": "Langue :",
        "lbl_model": "Modèle Whisper :",
        "model_eff_hint": "→ {eff} sera utilisé",
        "model_missing_hint": "⚠️ Modèle '{eff}' absent — installez-le via '⚙ Modèles installés & Mises à jour'",
        "diarize_missing_hint": "⚠️ Modèle locuteurs absent — installez-le via '⚙ Modèles installés & Mises à jour'",
        "chk_offline": "Mode hors-ligne (cache local)",
        "chk_diarize": "Identifier les locuteurs",
        "lbl_cpu_cores": "Cœurs CPU :",
        "cpu_hidden_gpu": "Cœurs CPU : masqué (GPU actif)",
        "warn_cpu_med": "⚠️ Charge CPU élevée",
        "warn_cpu_high": "⚠️ Charge CPU très élevée (risque de ralentissement du PC)",
        "hw_cuda": "Matériel détecté : GPU NVIDIA CUDA actif (accélération matérielle)",
        "hw_cpu": "Matériel détecté : CPU uniquement (batch_size=16)",
        "sec3": "3. Progression",
        "btn_start": "Démarrer la transcription",
        "btn_start_n": "Démarrer la transcription ({n} fichier{s})",
        "btn_pause": "Pause",
        "btn_resume": "Reprendre",
        "btn_stop": "Arrêter",
        "btn_token": "Configurer token Hugging Face",
        "btn_diag": "Diagnostic système",
        "btn_clear": "Effacer le journal",
        "lbl_tot_prog": "Avancement total :",
        "tot_fmt": "{done} / {total} fichier(s) traité(s) ({pct} %)",
        "cur_none": "Fichier en cours : Aucun",
        "cur_fmt": "Fichier en cours : {name}",
        "step_wait": "En attente",
        "step_vad": "Détection de l'activité vocale (VAD)...",
        "step_tr_init": "Transcription audio en cours...",
        "step_tr_pct": "Transcription ({pct} % — {cur} / {tot})",
        "step_align": "Alignement temporel des mots...",
        "step_diar": "Détection et identification des locuteurs...",
        "step_slides": "Extraction et détection des slides vidéo...",
        "step_doc": "Génération des livrables...",
        "step_done": "Fichier terminé",
        "step_skip": "Déjà traité (ignoré)",
        "step_err": "Échec (code {rc})",
        "step_paused": "EN PAUSE (CPU libéré)",
        "step_resumed": "Reprise en cours...",
        "step_stopped": "Traitement arrêté",
        "step_batch_done": "Lot terminé",
        "sec_queue": "4. File des fichiers à traiter",
        "sec_queue_fmt": "4. File des fichiers à traiter (Écoulé : {elapsed} | Restant est. : {remaining} | Fin estimée : {eta_end})",
        "col_file": "Fichier",
        "col_model": "Modèle",
        "col_status": "Statut",
        "col_task": "Tâches",
        "col_started": "Début",
        "col_finished": "Fin",
        "st_pending": "En attente",
        "st_running": "En cours ({pct} %)",
        "st_completed": "Terminé ({dur})",
        "st_skipped": "Déjà traité",
        "st_failed": "Échec", "st_cancelled": "Annulé", "st_cancelled": "Annulé",
        "sec_log": "Journal d'avancement",
        "dl_line": "[{name}] ⬇ Téléchargement modèle : {pct} % — {cur} / {tot} ({spd})",
        "dlg_no_sel_t": "Aucune sélection",
        "dlg_no_sel_m": "Veuillez d'abord sélectionner un fichier ou un dossier.",
        "dlg_no_fmt_t": "Aucun format sélectionné",
        "dlg_no_fmt_m": "Veuillez cocher au moins un format de sortie (txt, docx, docx avec slides ou pptx).",
        "dlg_pre_t": "Prérequis bloquant",
        "dlg_pre_m": "Un prérequis indispensable manque pour lancer le traitement.\nConsultez la correction indiquée dans le journal.",
        "dlg_quit_t": "Transcription en cours",
        "dlg_quit_m": "Une transcription est actuellement en cours.\nVoulez-vous vraiment l'arrêter immédiatement et quitter ?"
    },
    "en": {
        "win_title": "SlideScript v1.1.0 — Transcription, Speaker Diarization & Slide Extraction",
        "ui_lang_label": "Interface:", "app_ver": "v1.1.0",
        "btn_models_mgr": "⚙ Installed Models & Updates",
        "dlg_mod_title": "Installed Models & Cache Update",
        "dlg_mod_desc": "Check models to install locally or uncheck models to uninstall and free disk space:",
        "dlg_mod_apply": "Apply & Restart",
        "dlg_mod_cancel": "Cancel",
        "dlg_mod_none": "No changes selected.",
        "dlg_mod_del": "Uninstalling {name}...",
        "dlg_mod_dl": "Downloading {name} (Internet connection required)...",
        "dlg_mod_restart": "All changes applied. Restarting application...",
        "sec1": "1. Media Selection",
        "btn_1file": "Select a file",
        "btn_mfiles": "Select multiple files",
        "btn_folder": "Select a folder",
        "chk_sub": "Include subfolders",
        "no_sel": "No media selected yet.",
        "sel_1": "1 file selected: {name}",
        "sel_n": "{n} files selected in {parent}",
        "sel_dir": "Folder: {folder} ({n} media file(s) found)",
        "sec2": "2. Transcription Settings",
        "lbl_fmt": "Output formats:",
        "fmt_txt": ".txt",
        "fmt_docx": ".docx",
        "fmt_docx_slides": ".docx with slides",
        "fmt_pptx": ".pptx with transcripts in notes",
        "lbl_lang": "Language:",
        "lbl_model": "Whisper model:",
        "model_eff_hint": "→ {eff} will be used",
        "model_missing_hint": "⚠️ Model '{eff}' missing — install via '⚙ Installed Models & Updates'",
        "diarize_missing_hint": "⚠️ Speaker model missing — install via '⚙ Installed Models & Updates'",
        "chk_offline": "Offline mode (local cache)",
        "chk_diarize": "Identify speakers",
        "lbl_cpu_cores": "CPU cores:",
        "cpu_hidden_gpu": "CPU cores: hidden (GPU active)",
        "warn_cpu_med": "⚠️ High CPU load",
        "warn_cpu_high": "⚠️ Very high CPU load (may slow down PC)",
        "hw_cuda": "Hardware detected: NVIDIA CUDA GPU active (hardware acceleration)",
        "hw_cpu": "Hardware detected: CPU only (batch_size=16)",
        "sec3": "3. Progress",
        "btn_start": "Start transcription",
        "btn_start_n": "Start transcription ({n} file{s})",
        "btn_pause": "Pause",
        "btn_resume": "Resume",
        "btn_stop": "Stop",
        "btn_token": "Configure Hugging Face token",
        "btn_diag": "System diagnostics",
        "btn_clear": "Clear log",
        "lbl_tot_prog": "Overall progress:",
        "tot_fmt": "{done} / {total} file(s) processed ({pct}%)",
        "cur_none": "Current file: None",
        "cur_fmt": "Current file: {name}",
        "step_wait": "Waiting",
        "step_vad": "Voice Activity Detection (VAD)...",
        "step_tr_init": "Transcribing audio...",
        "step_tr_pct": "Transcription ({pct}% — {cur} / {tot})",
        "step_align": "Aligning word timestamps...",
        "step_diar": "Detecting & identifying speakers...",
        "step_slides": "Extracting & detecting video slides...",
        "step_doc": "Generating output files...",
        "step_done": "File completed",
        "step_skip": "Already processed (skipped)",
        "step_err": "Failed (code {rc})",
        "step_paused": "PAUSED (CPU freed)",
        "step_resumed": "Resuming...",
        "step_stopped": "Processing stopped",
        "step_batch_done": "Batch completed",
        "sec_queue": "4. File Queue & Status",
        "sec_queue_fmt": "4. File Queue & Status (Elapsed: {elapsed} | Est. remaining: {remaining} | Est. finish: {eta_end})",
        "col_file": "File",
        "col_model": "Model",
        "col_status": "Status",
        "col_task": "Tasks",
        "col_started": "Started",
        "col_finished": "Finished",
        "st_pending": "Pending",
        "st_running": "Running ({pct}%)",
        "st_completed": "Completed ({dur})",
        "st_skipped": "Already done",
        "st_failed": "Failed", "st_cancelled": "Cancelled", "st_cancelled": "Cancelled",
        "sec_log": "Progress Log",
        "dl_line": "[{name}] ⬇ Downloading model: {pct}% — {cur} / {tot} ({spd})",
        "dlg_no_sel_t": "No selection",
        "dlg_no_sel_m": "Please select a file or folder first.",
        "dlg_no_fmt_t": "No output format selected",
        "dlg_no_fmt_m": "Please check at least one output format (txt, docx, docx with slides or pptx).",
        "dlg_pre_t": "Blocking prerequisite",
        "dlg_pre_m": "A required prerequisite is missing.\nCheck the suggested fix in the log.",
        "dlg_quit_t": "Transcription in progress",
        "dlg_quit_m": "A transcription is currently running.\nDo you really want to stop it immediately and exit?"
    }
}

def create_button_icons(master):
    """Génère des icônes vectorielles nettes (anti-aliasées 4x) assorties à la charte graphique."""
    if not PIL_AVAILABLE:
        return {}
    icons = {}
    scale = 4
    sz = 14 * scale

    im_play = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    d = ImageDraw.Draw(im_play)
    d.polygon([(12, 6), (12, sz - 6), (sz - 6, sz // 2)], fill=(255, 255, 255, 255))
    icons["play"] = ImageTk.PhotoImage(im_play.resize((14, 14), Image.LANCZOS), master=master)

    im_pause = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    d = ImageDraw.Draw(im_pause)
    d.rectangle([12, 8, 22, sz - 8], fill=(17, 24, 39, 255))
    d.rectangle([sz - 22, 8, sz - 12, sz - 8], fill=(17, 24, 39, 255))
    icons["pause"] = ImageTk.PhotoImage(im_pause.resize((14, 14), Image.LANCZOS), master=master)

    im_stop = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    d = ImageDraw.Draw(im_stop)
    d.rounded_rectangle([10, 10, sz - 10, sz - 10], radius=5, fill=(17, 24, 39, 255))
    icons["stop"] = ImageTk.PhotoImage(im_stop.resize((14, 14), Image.LANCZOS), master=master)

    im_diag = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    d = ImageDraw.Draw(im_diag)
    d.ellipse([8, 8, 38, 38], outline=(17, 24, 39, 255), width=5)
    d.line([(33, 33), (48, 48)], fill=(17, 24, 39, 255), width=6)
    icons["diag"] = ImageTk.PhotoImage(im_diag.resize((14, 14), Image.LANCZOS), master=master)

    im_clear = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    d = ImageDraw.Draw(im_clear)
    d.line([(42, 8), (26, 26)], fill=(17, 24, 39, 255), width=5)
    d.polygon([(18, 24), (34, 36), (22, 50), (8, 38)], fill=(17, 24, 39, 255))
    icons["clear"] = ImageTk.PhotoImage(im_clear.resize((14, 14), Image.LANCZOS), master=master)

    # 9. Logo Hugging Face N&B / niveaux de gris (16x16, 571 octets)
    import base64
    from io import BytesIO
    hf_b64 = (
        "iVBORw0KGgoAAAANSUhEUgAAAA8AAAAPCAYAAAA71pVKAAADAElEQVR42l2TP0hrZwDFz/cn90vy3fwzYBMbg38afA9KxBc7iFRL26Fg6VYnEYQ3ukhGhTeo8AbFJU6dSheh0KXgoLyGroWIQ0GQpBQ01xqDUfP33vvd+3VRKD3jOT84yzkEz6pUKqRQKGgAKBaLr6WUnzPGUp7nXbXb7d8PDg7++j9H/mtsbGwkE4nE+2Qy+b1pmrFAIADXdfH4+Nh8eHj4qdFobJVKpd4LT9bX1+lzeTyTyZzmcrk3lFJwzn0AUEoRQghxHAfVavW3y8vLb03TtNfW1jQdGxsjpVLJT6fT+7lc7k04HHaCwSCUUlQpRTnnRAgBKaUzOTn55fj4+LtSqeSXy2XKTk5OdLFYfD0yMnKYSqXIxcUF39/fx8zMDOLxOM7Pz3F4eIjZ2Vkai8V0t9vNT01N/bi9vf1EAUBK+U06nTZubm50rVZDu92GbdvwfR/9fh/NZhP1ep3U63WkUqmYaZpfAAAHgFAolOOcwzAMLC4uYmlpCZRSKKWwsLCAubk5dDodaK21YRhaCDEFALxSqZDT01MOAJFIBEII7O3twXVdzM/P4/j4GPl8HisrK+j3+1BKEcaYUalUCC0UCtpxnL8ty8LV1ZW+vb1FNptFuVzGzs4Ozs7OMDExgevra1iWBcuyYNt2tVAoaLK1tfWJaZq/JhKJVwD0YDAgUkq0Wi1Uq1VMT0/DMAx0Oh2EQiENgNzf3//pOM53FMB0NBp9JYRwGWNkaGgIjuMgm81ieXkZ8XgcjuMgmUyCMUaEEG40Gv3Udd08DwQCGcaYrbXuhUKhWDAYpEII3N3deYQQ0mq19PDwMKOUghDiu67b5pzLQCDwMfc874ExJoQQKhKJ+M9ro0IIppSCEALhcBi+7/uGYfhPT0/C8zzhed4jB/Bzo9HISCnf9nq9TDAYpN1uVzebzQ+U0lvf9z/yPO8rKSXt9/t+r9drdDqdHwaDwS/k5VWrq6vx0dHRz4QQKdu2rd3d3Q8v2ebm5tdCiLRt2//UarU/jo6OHgHgX+uqWXfLlct3AAAAAElFTkSuQmCC"
    )
    im_hf = Image.open(BytesIO(base64.b64decode(hf_b64)))
    icons["hf"] = ImageTk.PhotoImage(im_hf, master=master)

    return icons


def create_hf_badge_icons(master):
    """Génère deux variantes colorées du logo Hugging Face sur fond transparent : vert (présent) et orange + ! rouge (manquant)."""
    if not PIL_AVAILABLE:
        return {}
    import base64
    from io import BytesIO
    hf_b64 = (
        "iVBORw0KGgoAAAANSUhEUgAAAA8AAAAPCAYAAAA71pVKAAADAElEQVR42l2TP0hrZwDFz/cn90vy3fwzYBMbg38afA9KxBc7iFRL26Fg6VYnEYQ3ukhGhTeo8AbFJU6dSheh0KXgoLyGroWIQ0GQpBQ01xqDUfP33vvd+3VRKD3jOT84yzkEz6pUKqRQKGgAKBaLr6WUnzPGUp7nXbXb7d8PDg7++j9H/mtsbGwkE4nE+2Qy+b1pmrFAIADXdfH4+Nh8eHj4qdFobJVKpd4LT9bX1+lzeTyTyZzmcrk3lFJwzn0AUEoRQghxHAfVavW3y8vLb03TtNfW1jQdGxsjpVLJT6fT+7lc7k04HHaCwSCUUlQpRTnnRAgBKaUzOTn55fj4+LtSqeSXy2XKTk5OdLFYfD0yMnKYSqXIxcUF39/fx8zMDOLxOM7Pz3F4eIjZ2Vkai8V0t9vNT01N/bi9vf1EAUBK+U06nTZubm50rVZDu92GbdvwfR/9fh/NZhP1ep3U63WkUqmYaZpfAAAHgFAolOOcwzAMLC4uYmlpCZRSKKWwsLCAubk5dDodaK21YRhaCDEFALxSqZDT01MOAJFIBEII7O3twXVdzM/P4/j4GPl8HisrK+j3+1BKEcaYUalUCC0UCtpxnL8ty8LV1ZW+vb1FNptFuVzGzs4Ozs7OMDExgevra1iWBcuyYNt2tVAoaLK1tfWJaZq/JhKJVwD0YDAgUkq0Wi1Uq1VMT0/DMAx0Oh2EQiENgNzf3//pOM53FMB0NBp9JYRwGWNkaGgIjuMgm81ieXkZ8XgcjuMgmUyCMUaEEG40Gv3Udd08DwQCGcaYrbXuhUKhWDAYpEII3N3deYQQ0mq19PDwMKOUghDiu67b5pzLQCDwMfc874ExJoQQKhKJ+M9ro0IIppSCEALhcBi+7/uGYfhPT0/C8zzhed4jB/Bzo9HISCnf9nq9TDAYpN1uVzebzQ+U0lvf9z/yPO8rKSXt9/t+r9drdDqdHwaDwS/k5VWrq6vx0dHRz4QQKdu2rd3d3Q8v2ebm5tdCiLRt2//UarU/jo6OHgHgX+uqWXfLlct3AAAAAElFTkSuQmCC"
    )
    base_gray = Image.open(BytesIO(base64.b64decode(hf_b64))).convert("RGBA")
    g_px = base_gray.load()
    out = {}

    # Variante 1 : Logo Hugging Face vert (#15803D) seul sur fond transparent (22x18)
    im_ok = Image.new("RGBA", (22, 18), (0, 0, 0, 0))
    px_ok = im_ok.load()
    for y in range(15):
        for x in range(15):
            r, g, b, a = g_px[x, y]
            if a > 20:
                t = ((r + g + b) / 3.0) / 255.0
                nr = int(21 * (1 - t) + 187 * t)
                ng = int(128 * (1 - t) + 247 * t)
                nb = int(61 * (1 - t) + 208 * t)
                px_ok[x + 1, y + 1] = (nr, ng, nb, a)
    out["hf_ok"] = ImageTk.PhotoImage(im_ok, master=master)

    # Variante 2 : Logo Hugging Face orange (#D97706) + point d'exclamation rouge (#DC2626) à côté (22x18)
    im_warn = Image.new("RGBA", (22, 18), (0, 0, 0, 0))
    px_w = im_warn.load()
    for y in range(15):
        for x in range(15):
            r, g, b, a = g_px[x, y]
            if a > 20:
                t = ((r + g + b) / 3.0) / 255.0
                nr = int(217 * (1 - t) + 254 * t)
                ng = int(119 * (1 - t) + 215 * t)
                nb = int(6 * (1 - t) + 170 * t)
                px_w[x + 1, y + 1] = (nr, ng, nb, a)
    dw = ImageDraw.Draw(im_warn)
    dw.rectangle([18, 2, 20, 11], fill=(220, 38, 38, 255))
    dw.rectangle([18, 13, 20, 16], fill=(220, 38, 38, 255))
    out["hf_missing"] = ImageTk.PhotoImage(im_warn, master=master)
    return out


class PillProgressBar(tk.Canvas):
    """Barre de progression bleue/orange arrondie avec pourcentage centré optionnel."""
    def __init__(self, parent, height=18, show_text=False, maximum=100.0, fill_color=None, **kwargs):
        bg_parent = kwargs.pop("bg", "#F0F0F0")
        super().__init__(parent, height=height, bg=bg_parent, highlightthickness=0, bd=0, **kwargs)
        self._max = float(maximum)
        self._val = 0.0
        self._show_text = show_text
        self._fill_color = fill_color or THEME_BLUE
        self.bind("<Configure>", lambda e: self._redraw())

    def configure(self, cnf=None, **kw):
        if "value" in kw:
            self._val = float(kw.pop("value"))
        if "maximum" in kw:
            self._max = max(1.0, float(kw.pop("maximum")))
        if "fill_color" in kw:
            self._fill_color = kw.pop("fill_color")
        if cnf and "value" in cnf:
            self._val = float(cnf.pop("value"))
        if cnf and "maximum" in cnf:
            self._max = max(1.0, float(cnf.pop("maximum")))
        if cnf and "fill_color" in cnf:
            self._fill_color = cnf.pop("fill_color")
        if kw or cnf:
            super().configure(cnf, **kw)
        self._redraw()

    config = configure

    def __setitem__(self, key, value):
        if key == "value":
            self._val = float(value)
            self._redraw()
        elif key == "maximum":
            self._max = max(1.0, float(value))
            self._redraw()
        else:
            super().__setitem__(key, value)

    def __getitem__(self, key):
        if key == "value":
            return self._val
        if key == "maximum":
            return self._max
        return super().__getitem__(key)

    def _draw_pill(self, x1, y1, x2, y2, r, fill, outline=""):
        if x2 - x1 < 2 * r:
            r = max(1, (x2 - x1) / 2)
        pts = [
            x1 + r, y1,
            x2 - r, y1,
            x2, y1,
            x2, y1 + r,
            x2, y2 - r,
            x2, y2,
            x2 - r, y2,
            x1 + r, y2,
            x1, y2,
            x1, y2 - r,
            x1, y1 + r,
            x1, y1
        ]
        return self.create_polygon(pts, smooth=True, fill=fill, outline=outline)

    def _redraw(self):
        self.delete("all")
        w = max(20, self.winfo_width())
        h = max(10, self.winfo_height())
        r = h / 2.0
        self._draw_pill(1, 1, w - 1, h - 1, r, fill=THEME_TROUGH_BG, outline="#CBD2D9")
        ratio = max(0.0, min(1.0, self._val / self._max)) if self._max > 0 else 0.0
        if ratio > 0.005:
            fill_w = max(h, (w - 2) * ratio)
            self._draw_pill(1, 1, min(w - 1, 1 + fill_w), h - 1, r, fill=self._fill_color)
        if self._show_text:
            pct_int = int(round(ratio * 100))
            txt_color = "#FFFFFF" if ratio >= 0.48 else "#1F2937"
            self.create_text(w / 2.0, h / 2.0, text=f"{pct_int} %", fill=txt_color, font=("Segoe UI", 8, "bold"))

def get_app_icon_path() -> str:
    candidates = []
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, "SlideScript.ico"))
    candidates.append(os.path.join(str(SCRIPT_DIR), "SlideScript.ico"))
    for p in candidates:
        if os.path.exists(p):
            return p
    return ""

def detect_nvidia_cuda() -> bool:
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    if not shutil.which("nvidia-smi"):
        return False
    try:
        smi = subprocess.run(["nvidia-smi"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5, creationflags=creationflags)
        if smi.returncode != 0:
            return False
    except Exception:
        return False

    runtime_py = RUNTIME_DIR / "python.exe"
    venv_py = SCRIPT_DIR / "whisperx-env" / "Scripts" / "python.exe"
    py_exe = runtime_py if runtime_py.exists() else (venv_py if venv_py.exists() else None)
    if not py_exe:
        return False

    try:
        res = subprocess.run(
            [str(py_exe), "-c", "import torch; print('1' if torch.cuda.is_available() else '0')"],
            capture_output=True, text=True, timeout=15, creationflags=creationflags
        )
        return res.stdout.strip() == "1"
    except Exception:
        return False

def check_hf_repos_access(token: str):
    results = {}
    blocked = False
    network_error = False

    for repo in REQUIRED_HF_REPOS:
        file_url = f"https://huggingface.co/{repo}/resolve/main/config.yaml"
        file_req = urllib.request.Request(file_url, method="HEAD", headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(file_req, timeout=6):
                results[repo] = "OK"
        except urllib.error.HTTPError as e:
            blocked = True
            if e.code == 401:
                results[repo] = "Refusé (401 : Jeton invalide)"
            elif e.code == 403:
                results[repo] = "Bloqué (403 : Conditions non acceptées)"
            else:
                results[repo] = f"HTTP {e.code}"
        except Exception:
            network_error = True
            results[repo] = "Non vérifié (Hors-ligne)"

    if blocked:
        return {"status": "blocked", "results": results}
    if network_error:
        return {"status": "offline", "results": results}
    return {"status": "ok", "results": results}

def get_media_duration(media_path: Path) -> float:
    ffprobe_cmd = shutil.which("ffprobe")
    if not ffprobe_cmd:
        return 0.0
    try:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        res = subprocess.run(
            [
                ffprobe_cmd, "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                str(media_path)
            ],
            capture_output=True, text=True, timeout=10, creationflags=creationflags
        )
        val = res.stdout.strip()
        return float(val) if val else 0.0
    except Exception:
        return 0.0

def has_video_stream(media_path: Path) -> bool:
    ffprobe_cmd = shutil.which("ffprobe")
    if not ffprobe_cmd:
        return media_path.suffix.lower() in {".mp4", ".mkv", ".mov", ".avi"}
    try:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        res = subprocess.run(
            [
                ffprobe_cmd, "-v", "error",
                "-select_streams", "v:0",
                "-show_entries", "stream=codec_type",
                "-of", "default=noprint_wrappers=1:nokey=1",
                str(media_path)
            ],
            capture_output=True, text=True, timeout=10, creationflags=creationflags
        )
        return "video" in res.stdout.lower()
    except Exception:
        return False

def format_short_duration(seconds: float) -> str:
    if not seconds or seconds < 0:
        return "00:00"
    total = int(seconds)
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}" if h > 0 else f"{m:02d}:{s:02d}"

def format_elapsed_buzz(seconds: float) -> str:
    total = max(0, int(seconds))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    if h > 0:
        return f"{h}h {m:02d}m"
    return f"{m}m {s:02d}s"

_DPAPI_ENTROPY = b"SlideScript::HF_TOKEN::DPAPI::v1"
_REG_APP_KEY = r"Software\SlideScript"
_REG_TOKEN_VAL = "HuggingFaceTokenDPAPI"

def _dpapi_protect(plaintext: str) -> bytes:
    from ctypes import wintypes
    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]
    raw = plaintext.encode("utf-8")
    in_buf = (ctypes.c_byte * len(raw)).from_buffer_copy(raw)
    ent_buf = (ctypes.c_byte * len(_DPAPI_ENTROPY)).from_buffer_copy(_DPAPI_ENTROPY)
    in_blob = DATA_BLOB(len(raw), in_buf)
    ent_blob = DATA_BLOB(len(_DPAPI_ENTROPY), ent_buf)
    out_blob = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptProtectData(ctypes.byref(in_blob), "SlideScript_HF_TOKEN", ctypes.byref(ent_blob), None, None, 0x01, ctypes.byref(out_blob)):
        raise OSError("CryptProtectData failed")
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)

def _dpapi_unprotect(cipher_bytes: bytes) -> str:
    from ctypes import wintypes
    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]
    in_buf = (ctypes.c_byte * len(cipher_bytes)).from_buffer_copy(cipher_bytes)
    ent_buf = (ctypes.c_byte * len(_DPAPI_ENTROPY)).from_buffer_copy(_DPAPI_ENTROPY)
    in_blob = DATA_BLOB(len(cipher_bytes), in_buf)
    ent_blob = DATA_BLOB(len(_DPAPI_ENTROPY), ent_buf)
    out_blob = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(in_blob), None, ctypes.byref(ent_blob), None, None, 0x01, ctypes.byref(out_blob)):
        return ""
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData).decode("utf-8", errors="ignore").strip()
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)

def get_hf_token(config_data: dict) -> str:
    env_val = os.environ.get("HF_TOKEN", "").strip()
    if env_val.startswith("hf_"):
        return env_val
    if os.name == "nt":
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_APP_KEY, 0, winreg.KEY_READ) as key:
                blob, reg_type = winreg.QueryValueEx(key, _REG_TOKEN_VAL)
                if reg_type == winreg.REG_BINARY and blob:
                    dec = _dpapi_unprotect(blob)
                    if dec.startswith("hf_"):
                        os.environ["HF_TOKEN"] = dec
                        return dec
        except OSError:
            pass
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ) as key:
                reg_val, _ = winreg.QueryValueEx(key, "HF_TOKEN")
                reg_val = str(reg_val).strip()
                if reg_val.startswith("hf_"):
                    save_hf_token(reg_val, config_data)
                    return reg_val
        except OSError:
            pass
    return ""

def save_hf_token(token: str, config_data: dict):
    token = token.strip()
    os.environ["HF_TOKEN"] = token
    config_data.pop("hf_token", None)
    if os.name == "nt":
        try:
            protected_blob = _dpapi_protect(token)
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _REG_APP_KEY, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, _REG_TOKEN_VAL, 0, winreg.REG_BINARY, protected_blob)
        except Exception:
            pass

BLOCKING_ERRORS = [
    {
        "signatures": ["401 client error", "403 client error", "gatedrepoerror", "cannot access gated repo", "invalid user token", "repositorynotfounderror"],
        "titre": "Accès refusé aux dépôts pyannote sur Hugging Face",
        "etape": "Étape 3 — Configuration de l'authentification et des 2 dépôts Hugging Face",
        "cause": (
            "WhisperX exige que les conditions d'utilisation soient acceptées sur DEUX dépôts Hugging Face :\n"
            "• https://huggingface.co/pyannote/speaker-diarization-community-1\n"
            "• https://huggingface.co/pyannote/segmentation-3.0"
        ),
        "action": (
            "1. Connectez-vous sur Hugging Face avec le compte propriétaire de votre jeton.\n"
            "2. Ouvrez les DEUX liens ci-dessus et cliquez sur 'Agree and access repository' sur chacun.\n"
            "3. Cliquez ensuite sur '🔑 Configurer token Hugging Face'."
        )
    },
    {
        "signatures": ["localentrynotfounderror", "cannot find the requested files in the cached path", "outgoing traffic has been disabled"],
        "titre": "Modèle absent du cache local alors que le mode hors-ligne est actif",
        "etape": "Étape 4 — Validation fonctionnelle (Premier téléchargement)",
        "cause": "Les modèles Whisper ou pyannote ne sont pas encore téléchargés en cache local.",
        "action": (
            "1. Ouvrez '⚙ Modèles installés' en haut à droite pour cocher et installer le modèle requis,\n"
            "   puis relancez la transcription."
        )
    },
    {
        "signatures": ["filenotfounderror: [winerror 2]", "no such file or directory: 'ffmpeg'", "ffmpeg was not found"],
        "titre": "FFmpeg introuvable sur le système",
        "etape": "Étape 1 — Installation des prérequis système",
        "cause": "Impossible de décoder le fichier audio/vidéo car FFmpeg est absent.",
        "action": "1. Exécutez dans PowerShell : winget install Gyan.FFmpeg\n2. Relancez l'application."
    },
    {
        "signatures": ["outofmemoryerror", "cuda out of memory", "defaultcpuallocator: not enough memory", "memoryerror"],
        "titre": "Mémoire saturée pendant l'exécution du modèle",
        "etape": "Section 2 — Choix du modèle Whisper",
        "cause": "Le modèle choisi consomme plus de RAM/VRAM que disponible sur la machine.",
        "action": "Sélectionnez un modèle plus léger ('small' ou 'tiny') et relancez."
    },
    {
        "signatures": ["permissionerror: [errno 13]", "permissionerror: [winerror 32]"],
        "titre": "Fichier verrouillé ou accès refusé en écriture",
        "etape": "Section 2 — Précautions d'exploitation",
        "cause": "Le fichier cible est ouvert dans Word/PowerPoint ou le répertoire est protégé en écriture.",
        "action": "Fermez le document cible dans Word/PowerPoint s'il est ouvert."
    }
]

def diagnose_blocking_failure(log_text: str):
    tail = "\n".join(log_text.splitlines()[-40:]).lower()
    return [rule for rule in BLOCKING_ERRORS if any(sig in tail for sig in rule["signatures"])]

def suspend_or_resume_pid(pid: int, suspend: bool):
    handle = ctypes.windll.kernel32.OpenProcess(0x0800, False, pid)
    if handle:
        try:
            if suspend:
                ctypes.windll.ntdll.NtSuspendProcess(handle)
            else:
                ctypes.windll.ntdll.NtResumeProcess(handle)
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)

def set_process_cpu_cores_affinity(pid: int, num_cores: int):
    if os.name != "nt" or not pid:
        return
    try:
        total = os.cpu_count() or 8
        n = max(1, min(total, int(num_cores)))
        mask = (1 << n) - 1
        PROCESS_SET_INFORMATION = 0x0200
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_SET_INFORMATION, False, pid)
        if handle:
            try:
                ctypes.windll.kernel32.SetProcessAffinityMask(handle, ctypes.c_size_t(mask))
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
    except Exception:
        pass

def format_time(seconds):
    if seconds is None:
        return "00:00:00.000"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"

def speaker_label_and_color(raw_speaker):
    if not raw_speaker:
        return "Speaker 1", SPEAKER_COLORS[0]
    try:
        idx = int(raw_speaker.split("_")[-1])
        color_idx = max(0, idx - 1) % len(SPEAKER_COLORS)
        return f"Speaker {idx}", SPEAKER_COLORS[color_idx]
    except ValueError:
        return raw_speaker, SPEAKER_COLORS[0]

def build_merged_blocks(segments):
    merged_blocks = []
    last_raw_speaker = "SPEAKER_00"
    speaker_order_map = {}

    for seg in segments:
        raw_spk = seg.get("speaker") or last_raw_speaker
        last_raw_speaker = raw_spk
        text = seg.get("text", "").strip()
        start = seg.get("start", 0.0)
        end = seg.get("end", 0.0)
        if not text:
            continue

        if raw_spk not in speaker_order_map:
            speaker_order_map[raw_spk] = f"SPEAKER_{len(speaker_order_map) + 1:02d}"
        ordered_spk = speaker_order_map[raw_spk]

        if merged_blocks and merged_blocks[-1]["speaker"] == ordered_spk:
            merged_blocks[-1]["end"] = end
            merged_blocks[-1]["text"] += " " + text
        else:
            merged_blocks.append({"speaker": ordered_spk, "start": start, "end": end, "text": text})
    return merged_blocks

def _make_slide_thumb(img_path: Path):
    with Image.open(img_path) as im:
        w, h = im.size
        crop_box = (int(w * 0.05), int(h * 0.06), int(w * 0.95), int(h * 0.85))
        return im.crop(crop_box).convert("L").resize((160, 90))

def _frame_diff_ratio(img_a, img_b) -> float:
    diff = ImageChops.difference(img_a, img_b)
    hist = diff.histogram()
    total_px = sum(hist)
    if total_px == 0:
        return 0.0
    changed_px = sum(hist[32:])
    return changed_px / total_px

def extract_slides_from_video(media_path: Path, total_duration: float, temp_dir: Path, log_fn) -> list:
    ffmpeg_cmd = shutil.which("ffmpeg")
    if not ffmpeg_cmd or not PIL_AVAILABLE:
        log_fn("[WARN] Extraction de slides indisponible (FFmpeg ou Pillow manquant).\n", "error")
        return []

    if not has_video_stream(media_path):
        log_fn(f"[INFO] Aucune piste vidéo détectée dans {media_path.name} : extraction de slides ignorée.\n")
        return []

    frames_dir = temp_dir / "frames"
    frames_dir.mkdir(parents=True, exist_ok=True)

    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    pattern = str(frames_dir / "frame_%05d.png")
    cmd = [
        ffmpeg_cmd, "-y", "-i", str(media_path),
        "-vf", "fps=0.5,scale=iw:ih:flags=lanczos",
        "-compression_level", "3",
        pattern
    ]
    res = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creationflags)
    if res.returncode != 0:
        log_fn("[WARN] Échec de l'extraction d'images FFmpeg pour les slides.\n", "error")
        return []

    frame_files = sorted(frames_dir.glob("frame_*.png"))
    if not frame_files:
        return []

    step_sec = 2.0
    CHANGE_THRESHOLD = 0.12
    MIN_STABLE_FRAMES = 3

    transitions = []
    last_stable_thumb = None
    candidate_thumb = None
    candidate_start_idx = 0
    candidate_count = 0

    for idx, fpath in enumerate(frame_files):
        try:
            thumb = _make_slide_thumb(fpath)
        except Exception:
            continue

        if last_stable_thumb is None:
            last_stable_thumb = thumb
            # Prendre la 2e frame si disponible pour éviter un éventuel fondu d'ouverture
            init_img = frame_files[1] if len(frame_files) > 1 else fpath
            transitions.append((0.0, init_img, thumb))
            continue

        diff_from_stable = _frame_diff_ratio(last_stable_thumb, thumb)
        if diff_from_stable >= CHANGE_THRESHOLD:
            if candidate_thumb is None or _frame_diff_ratio(candidate_thumb, thumb) >= CHANGE_THRESHOLD:
                candidate_thumb = thumb
                candidate_start_idx = idx
                candidate_count = 1
            else:
                candidate_count += 1
                if candidate_count >= MIN_STABLE_FRAMES:
                    if _frame_diff_ratio(last_stable_thumb, thumb) >= CHANGE_THRESHOLD:
                        t_sec = candidate_start_idx * step_sec
                        # Utiliser fpath (la frame stabilisée après animation) au lieu de la 1re frame de transition
                        transitions.append((t_sec, fpath, thumb))
                        last_stable_thumb = thumb
                    candidate_thumb = None
                    candidate_count = 0
        else:
            candidate_thumb = None
            candidate_count = 0

    end_video = total_duration if total_duration > 0 else (len(frame_files) * step_sec)
    slides = []
    for i, (t_start, img_path, _) in enumerate(transitions):
        t_end = transitions[i + 1][0] if (i + 1) < len(transitions) else max(end_video, t_start + step_sec)
        slides.append({
            "index": i + 1,
            "start": t_start,
            "end": t_end,
            "image_path": img_path
        })

    log_fn(f"[INFO] {len(slides)} diapositive(s) unique(s) détectée(s) dans la vidéo.\n", "success")
    return slides

def assign_blocks_to_slides(slides: list, raw_segments: list) -> list:
    if not slides:
        return []

    speaker_order_map = {}
    last_raw_speaker = "SPEAKER_00"
    ordered_segments = []

    for seg in raw_segments:
        raw_spk = seg.get("speaker") or last_raw_speaker
        last_raw_speaker = raw_spk
        text = seg.get("text", "").strip()
        start = seg.get("start", 0.0)
        end = seg.get("end", 0.0)
        if not text:
            continue
        if raw_spk not in speaker_order_map:
            speaker_order_map[raw_spk] = f"SPEAKER_{len(speaker_order_map) + 1:02d}"
        ordered_segments.append({
            "speaker": speaker_order_map[raw_spk],
            "start": start,
            "end": end,
            "text": text
        })

    grouped = [{"slide": s, "blocks": []} for s in slides]
    for seg in ordered_segments:
        mid = (seg["start"] + seg["end"]) / 2.0
        target_idx = 0
        for i, item in enumerate(grouped):
            s = item["slide"]
            if s["start"] <= mid < s["end"]:
                target_idx = i
                break
            if mid >= s["start"]:
                target_idx = i

        slide_blocks = grouped[target_idx]["blocks"]
        if slide_blocks and slide_blocks[-1]["speaker"] == seg["speaker"]:
            slide_blocks[-1]["end"] = seg["end"]
            slide_blocks[-1]["text"] += " " + seg["text"]
        else:
            slide_blocks.append(dict(seg))

    return grouped

def save_as_docx(base_name, merged_blocks, out_path: Path, log_fn, card_fn):
    if not DOCX_AVAILABLE:
        log_fn("[ERREUR] Bibliothèque 'python-docx' absente.\n", "error")
        return False

    doc = Document()
    normal_style = doc.styles["Normal"]
    normal_style.font.name = FONT_NAME
    normal_style.font.size = Pt(11)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run(base_name)
    title_run.bold = True
    title_run.font.name = FONT_NAME
    title_run.font.size = Pt(14)

    for block in merged_blocks:
        label, color = speaker_label_and_color(block["speaker"])
        time_str = f"[{format_time(block['start'])} – {format_time(block['end'])}]"

        header_p = doc.add_paragraph()
        header_p.paragraph_format.space_before = Pt(8)
        header_p.paragraph_format.space_after = Pt(2)

        run_header = header_p.add_run(f"● {label}  {time_str}")
        run_header.bold = True
        run_header.font.name = FONT_NAME
        run_header.font.color.rgb = color

        formatted_text = block["text"].replace(". ", ".\n")
        body_p = doc.add_paragraph()
        body_p.paragraph_format.space_after = Pt(6)
        body_run = body_p.add_run(formatted_text)
        body_run.font.name = FONT_NAME

    try:
        doc.save(out_path)
        log_fn(f"[OK] Généré : {out_path}\n", "success")
        return True
    except PermissionError:
        log_fn(f"[ERREUR] Fichier verrouillé : {out_path}\n", "error")
        card_fn("Fichier Word verrouillé", "Précautions", f"Le fichier {out_path.name} est ouvert dans Word.", "Fermez ce document dans Word et relancez.")
        return False

def save_as_docx_with_slides(base_name, slides_grouped, merged_blocks, out_path: Path, log_fn, card_fn):
    if not DOCX_AVAILABLE:
        log_fn("[ERREUR] Bibliothèque 'python-docx' absente.\n", "error")
        return False

    if not slides_grouped:
        return save_as_docx(base_name, merged_blocks, out_path, log_fn, card_fn)

    doc = Document()
    normal_style = doc.styles["Normal"]
    normal_style.font.name = FONT_NAME
    normal_style.font.size = Pt(10.5)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run(f"{base_name} — Support & Transcription")
    title_run.bold = True
    title_run.font.name = FONT_NAME
    title_run.font.size = Pt(14)

    for entry in slides_grouped:
        slide = entry["slide"]
        blocks = entry["blocks"]

        s_hdr = doc.add_paragraph()
        s_hdr.paragraph_format.space_before = Pt(12)
        s_hdr.paragraph_format.space_after = Pt(4)
        s_hdr.paragraph_format.keep_with_next = True
        r_s = s_hdr.add_run(f"Slide {slide['index']}  [{format_time(slide['start'])} – {format_time(slide['end'])}]")
        r_s.bold = True
        r_s.font.name = FONT_NAME
        r_s.font.size = Pt(11.5)
        r_s.font.color.rgb = RGBColor(15, 76, 129)

        if slide["image_path"].exists():
            img_p = doc.add_paragraph()
            img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            img_p.paragraph_format.space_after = Pt(6)
            img_p.paragraph_format.keep_with_next = True
            run_img = img_p.add_run()
            run_img.add_picture(str(slide["image_path"]), width=Inches(6.3))

        if not blocks:
            empty_p = doc.add_paragraph()
            empty_p.paragraph_format.space_after = Pt(6)
            r_emp = empty_p.add_run("(Aucune intervention vocale sur cette diapositive)")
            r_emp.italic = True
            r_emp.font.name = FONT_NAME
            r_emp.font.color.rgb = RGBColor(120, 120, 120)
        else:
            for block in blocks:
                label, color = speaker_label_and_color(block["speaker"])
                time_str = f"[{format_time(block['start'])} – {format_time(block['end'])}]"

                header_p = doc.add_paragraph()
                header_p.paragraph_format.space_before = Pt(4)
                header_p.paragraph_format.space_after = Pt(1)
                run_header = header_p.add_run(f"● {label}  {time_str}")
                run_header.bold = True
                run_header.font.name = FONT_NAME
                run_header.font.color.rgb = color

                formatted_text = block["text"].replace(". ", ".\n")
                body_p = doc.add_paragraph()
                body_p.paragraph_format.space_after = Pt(4)
                body_run = body_p.add_run(formatted_text)
                body_run.font.name = FONT_NAME

    try:
        doc.save(out_path)
        log_fn(f"[OK] Généré : {out_path}\n", "success")
        return True
    except PermissionError:
        log_fn(f"[ERREUR] Fichier verrouillé : {out_path}\n", "error")
        card_fn("Fichier Word verrouillé", "Précautions", f"Le fichier {out_path.name} est ouvert dans Word.", "Fermez ce document dans Word et relancez.")
        return False

def save_as_pptx(base_name, slides_grouped, out_path: Path, log_fn, card_fn):
    if not PPTX_AVAILABLE:
        log_fn("[ERREUR] Bibliothèque 'python-pptx' absente.\n", "error")
        return False

    if not slides_grouped:
        log_fn(f"[INFO] Aucune diapositive extraite pour {base_name} : fichier .pptx non généré.\n")
        return True

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]

    for entry in slides_grouped:
        slide_info = entry["slide"]
        blocks = entry["blocks"]
        slide = prs.slides.add_slide(blank_layout)

        img_path = slide_info["image_path"]
        if img_path.exists():
            slide.shapes.add_picture(str(img_path), PptxInches(0), PptxInches(0), width=prs.slide_width)

        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.clear()

        p_title = tf.paragraphs[0]
        p_title.text = f"Slide {slide_info['index']} [{format_time(slide_info['start'])} – {format_time(slide_info['end'])}]"

        if not blocks:
            p_empty = tf.add_paragraph()
            p_empty.text = "(Aucune intervention vocale sur cette diapositive)"
        else:
            for block in blocks:
                label, _ = speaker_label_and_color(block["speaker"])
                time_str = f"[{format_time(block['start'])} – {format_time(block['end'])}]"
                p_spk = tf.add_paragraph()
                p_spk.text = f"\n● {label} {time_str} :"
                p_txt = tf.add_paragraph()
                p_txt.text = block["text"]

    try:
        prs.save(str(out_path))
        log_fn(f"[OK] Généré : {out_path}\n", "success")
        return True
    except PermissionError:
        log_fn(f"[ERREUR] Fichier PowerPoint verrouillé : {out_path}\n", "error")
        card_fn("Fichier PowerPoint verrouillé", "Précautions", f"Le fichier {out_path.name} est ouvert dans PowerPoint.", "Fermez cette présentation dans PowerPoint et relancez.")
        return False

def save_as_txt(base_name, merged_blocks, out_path: Path, log_fn, card_fn):
    lines = [base_name, "=" * len(base_name), ""]
    for block in merged_blocks:
        label, _ = speaker_label_and_color(block["speaker"])
        time_str = f"[{format_time(block['start'])} – {format_time(block['end'])}]"
        formatted_text = block["text"].replace(". ", ".\n")
        lines.append(f"● {label}  {time_str}")
        lines.append(formatted_text)
        lines.append("")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        log_fn(f"[OK] Généré : {out_path}\n", "success")
        return True
    except PermissionError:
        log_fn(f"[ERREUR] Accès refusé : {out_path}\n", "error")
        return False


def write_model_debug(msg: str):
    try:
        from datetime import datetime as _dt
        ts = _dt.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{ts}] {msg}\n"
        with open(SCRIPT_DIR / "debug_models.log", "a", encoding="utf-8") as _f:
            _f.write(line)
    except Exception:
        pass

def get_hf_cache_path() -> Path:
    return HF_CACHE_DIR if HF_CACHE_DIR.exists() else (Path.home() / ".cache" / "huggingface" / "hub")

def is_whisper_model_installed(model_id: str) -> bool:
    if not model_id or model_id == "large.en":
        return False
    cache = get_hf_cache_path()
    if model_id == "pyannote":
        d1 = cache / "models--pyannote--speaker-diarization-community-1"
        d2 = cache / "models--pyannote--segmentation-3.0"
        return d1.exists() and any(d1.rglob("*.bin")) and d2.exists() and any(d2.rglob("*.bin"))
    if model_id == "wav2vec2-fr":
        d_fr = cache / "models--jonatasgrosman--wav2vec2-large-xlsr-53-french"
        return d_fr.exists() and (any(d_fr.rglob("pytorch_model.bin")) or any(d_fr.rglob("model.safetensors")))
    repo_dir = cache / f"models--Systran--faster-whisper-{model_id}"
    return repo_dir.exists() and any(repo_dir.rglob("model.bin"))

def get_available_base_models() -> list:
    return ["tiny", "small", "medium", "large-v3"]


def fetch_remote_model_meta(repo_id: str, timeout: float = 6.0):
    """Interroge l'API Hugging Face pour récupérer (sha, taille_totale_octets) des fichiers utiles."""
    if not repo_id:
        return None, 0
    try:
        import json, urllib.request
        url = f"https://huggingface.co/api/models/{repo_id}"
        req = urllib.request.Request(url, headers={"User-Agent": "SlideScript/1.1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))
        sha = data.get("sha")
        ignore_ext = (".h5", ".msgpack", ".ot", ".tflite")
        ignore_prefix = ("coreml/", "openvino/")
        total_bytes = 0
        for sib in data.get("siblings", []):
            rname = sib.get("rfilename", "")
            if rname.endswith(ignore_ext) or rname.startswith(ignore_prefix):
                continue
            sz = sib.get("size")
            if isinstance(sz, int) and sz > 0:
                total_bytes += sz
            elif isinstance(sib.get("lfs"), dict) and isinstance(sib["lfs"].get("size"), int):
                total_bytes += sib["lfs"]["size"]
        return sha, total_bytes
    except Exception:
        return None, 0

def get_local_model_size_bytes(repo_id: str) -> int:
    if not repo_id:
        return 0
    folder_name = f"models--{repo_id.replace('/', '--')}"
    candidates = [HF_CACHE_DIR / folder_name, HF_CACHE_DIR / "hub" / folder_name]
    for d in candidates:
        if d.exists():
            td = (d / "blobs") if (d / "blobs").exists() else d
            total = 0
            for f in td.rglob("*"):
                if f.is_file() and not f.name.endswith(".lock"):
                    try:
                        total += f.stat().st_size
                    except OSError:
                        pass
            if total > 0:
                return total
    return 0

def format_size_human(b: int, lang: str = "fr") -> str:
    if b <= 0: return ""
    mb = b / (1024 * 1024)
    return (f"{mb/1024:.2f}".replace(".", "," if lang=="fr" else ".") + (" Go" if lang=="fr" else " GB")) if mb >= 1024 else f"{mb:.0f} " + ("Mo" if lang=="fr" else "MB")

class ModelManagerDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent_app = parent
        self.title(parent.tr("dlg_mod_title") + "  [SlideScript v1.1.0]")
        write_model_debug("=== Ouverture ModelManagerDialog v4.6-DEBUG ===")
        self.geometry("900x660")
        self.minsize(850, 560)
        self.transient(parent)
        self.grab_set()
        self.active_tasks = {}
        self.task_rows_by_mid = {}
        self.any_modified = False
        self.row_widgets = {}
        self.hf_badge_labels = []
        self.hf_icons = create_hf_badge_icons(self)
        self._poller_active = False
        self.updates_available = set()
        self._update_check_proc = None
        self._update_check_file = None

        if hasattr(parent, "_apply_window_icon"):
            self.after(50, lambda: parent._apply_window_icon(self))

        is_fr = (parent.ui_lang == "fr")
        frame = ttk.Frame(self, padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        if getattr(parent, "is_running", False):
            desc_txt = (
                "Consultation en cours de transcription — Les actions d'installation, mise à jour et désinstallation sont temporairement désactivées jusqu'à la fin du traitement :"
                if is_fr else
                "Read-only view during active transcription — Install, update, and uninstall actions are temporarily disabled until processing finishes:"
            )
        else:
            desc_txt = (
                "État des modèles dans .\\hf_cache\\ — Lancez directement une installation, mise à jour ou désinstallation (opérations parallèles supportées) :"
                if is_fr else
                "Model status in .\\hf_cache\\ — Trigger installation, update, or uninstallation directly (parallel tasks supported):"
            )
        self.lbl_mgr_desc = ttk.Label(frame, text=desc_txt, wraplength=790, justify=tk.LEFT, font=("Segoe UI", 9, "bold"))
        self.lbl_mgr_desc.pack(anchor=tk.W, pady=(0, 8))

        whisper_frame = ttk.LabelFrame(frame, text=("Modèles de transcription (Whisper)" if is_fr else "Transcription Models (Whisper)"), padding=8)
        whisper_frame.pack(fill=tk.X, pady=(0, 8))
        align_frame = ttk.LabelFrame(frame, text=("Modèles d\'alignement forcé (optionnel — améliorer les horodatages)" if is_fr else "Forced Alignment Models (optional — improve timestamps)"), padding=8)
        align_frame.pack(fill=tk.X, pady=(0, 8))
        diar_frame = ttk.LabelFrame(frame, text=("Modèles d\'identification des locuteurs" if is_fr else "Speaker Identification Models"), padding=8)
        diar_frame.pack(fill=tk.X, pady=(0, 8))
        lbl_key = "label_fr" if is_fr else "label_en"

        for item in WHISPER_MODEL_CATALOG:
            mid = item["id"]
            installed = is_whisper_model_installed(mid)
            target_box = whisper_frame if item.get("category") == "whisper" else (align_frame if mid == "wav2vec2-fr" else diar_frame)

            row = ttk.Frame(target_box)
            row.pack(fill=tk.X, pady=2)

            ind_lbl = tk.Label(
                row,
                text=("✔" if installed else "○"),
                fg=("#15803D" if installed else "#9CA3AF"),
                font=("Segoe UI", 10, "bold"),
                width=2,
                anchor=tk.CENTER
            )
            ind_lbl.pack(side=tk.LEFT, padx=(0, 4))

            name_lbl = ttk.Label(
                row,
                text=item[lbl_key],
                foreground=("#111827" if item["available"] else "#9CA3AF"),
                font=("Segoe UI", 9, "bold" if installed else "normal")
            )
            name_lbl.pack(side=tk.LEFT)

            btn_uninst = None
            btn_act = None
            if item["available"]:
                btn_uninst = tk.Button(
                    row,
                    text=("🗑 Désinstaller" if is_fr else "🗑 Uninstall"),
                    command=lambda it=item: self.start_model_task(it, mode="uninstall"),
                    width=13,
                    bg="#FEE2E2",
                    fg="#991B1B",
                    activebackground="#FECACA",
                    activeforeground="#7F1D1D",
                    disabledforeground="#9CA3AF",
                    font=("Segoe UI", 8, "bold"),
                    relief=tk.FLAT,
                    padx=6,
                    pady=2,
                    cursor="hand2"
                )
                btn_uninst.pack(side=tk.RIGHT, padx=(6, 0))

                btn_act = tk.Button(
                    row,
                    text="",
                    command=lambda it=item: self.start_model_task(it, mode="update" if is_whisper_model_installed(it["id"]) else "install"),
                    width=15,
                    font=("Segoe UI", 8, "bold"),
                    relief=tk.FLAT,
                    padx=6,
                    pady=2,
                    cursor="hand2"
                )
                btn_act.pack(side=tk.RIGHT, padx=(6, 0))

            if item.get("gated"):
                badge_lbl = tk.Label(row, cursor="hand2", bd=0, padx=2, pady=0)
                badge_lbl.pack(side=tk.RIGHT, padx=(6, 4))
                badge_lbl.bind("<Button-1>", lambda e: self._on_configure_token_click())
                self.hf_badge_labels.append(badge_lbl)

            self.row_widgets[mid] = {
                "item": item,
                "ind_lbl": ind_lbl,
                "name_lbl": name_lbl,
                "btn_act": btn_act,
                "btn_uninst": btn_uninst,
            }
            self._refresh_single_row_state(mid)

        self._refresh_hf_token_badges()
        self.after(180, self._start_remote_update_check)

        tasks_title = "Opérations en cours / récentes" if is_fr else "Active & Recent Operations"
        self.tasks_box = ttk.LabelFrame(frame, text=tasks_title, padding=8)
        self.tasks_box.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.tasks_canvas = tk.Canvas(self.tasks_box, height=130, highlightthickness=0, bd=0)
        self.tasks_scroll = ttk.Scrollbar(self.tasks_box, orient=tk.VERTICAL, command=self.tasks_canvas.yview)
        self.tasks_inner = ttk.Frame(self.tasks_canvas)
        self.tasks_inner.bind("<Configure>", lambda e: self.tasks_canvas.configure(scrollregion=self.tasks_canvas.bbox("all")))
        self._tasks_win_id = self.tasks_canvas.create_window((0, 0), window=self.tasks_inner, anchor="nw")
        self.tasks_canvas.bind("<Configure>", lambda e: self.tasks_canvas.itemconfigure(self._tasks_win_id, width=e.width))
        self.tasks_canvas.configure(yscrollcommand=self.tasks_scroll.set)
        self.tasks_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tasks_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.lbl_no_tasks = ttk.Label(
            self.tasks_inner,
            text=("Aucune opération lancée. Cliquez sur « Installer », « Mettre à jour » ou « Désinstaller » sur un ou plusieurs modèles."
                  if is_fr else
                  "No active tasks. Click « Install », « Update », or « Uninstall » on one or more models."),
            foreground="#6B7280",
            font=("Segoe UI", 8, "italic")
        )
        self.lbl_no_tasks.pack(anchor=tk.W, pady=6)

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)
        tk.Button(
            btn_row,
            text=("📄 Ouvrir debug_models.log" if is_fr else "📄 Open debug_models.log"),
            command=self._open_debug_log,
            bg="#E5E7EB",
            fg="#111827",
            font=("Segoe UI", 8),
            relief=tk.FLAT,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT)
        self.btn_close = tk.Button(
            btn_row,
            text=("Fermer" if is_fr else "Close"),
            command=self.on_close,
            bg=THEME_BTN_SEC_BG,
            fg="#111827",
            activebackground=THEME_BTN_SEC_ACTIVE,
            activeforeground="#111827",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            padx=14,
            pady=4,
            cursor="hand2"
        )
        self.btn_close.pack(side=tk.RIGHT)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._apply_dynamic_labels()
        self._start_remote_check()

    def _open_debug_log(self):
        lf = SCRIPT_DIR / "debug_models.log"
        if lf.exists():
            os.startfile(str(lf))

    def _force_rmtree(self, target_dir: Path):
        import stat
        if not target_dir.exists() and not target_dir.is_symlink():
            return
        for p in sorted(target_dir.rglob("*"), reverse=True):
            try:
                if p.is_symlink() or p.is_file():
                    os.chmod(p, stat.S_IWRITE)
                    p.unlink(missing_ok=True)
                elif p.is_dir():
                    p.rmdir()
            except Exception:
                pass
        def _on_rm_error(func, path, exc_info):
            try:
                os.chmod(path, stat.S_IWRITE)
                func(path)
            except Exception:
                pass
        if target_dir.exists():
            shutil.rmtree(target_dir, onerror=_on_rm_error)
        if target_dir.exists() and os.name == "nt":
            try:
                subprocess.run(
                    ["cmd.exe", "/c", "rmdir", "/s", "/q", str(target_dir)],
                    stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NO_WINDOW, timeout=10
                )
            except Exception:
                pass
        write_model_debug(f"_force_rmtree({target_dir.name}) -> exists={target_dir.exists()}")

    def _refresh_single_row_state(self, mid: str, force_installed: bool = None):
        w = self.row_widgets.get(mid)
        if not w:
            return
        item = w["item"]
        if not item["available"]:
            return
        is_fr = (self.parent_app.ui_lang == "fr")
        installed = is_whisper_model_installed(mid) if force_installed is None else force_installed
        busy_mid = mid in self.active_tasks
        locked_by_run = bool(getattr(self.parent_app, "is_running", False))
        disable_actions = busy_mid or locked_by_run

        w["ind_lbl"].config(text=("✔" if installed else "○"), fg=("#15803D" if installed else "#9CA3AF"))
        lbl_key = "label_fr" if is_fr else "label_en"
        raw_lbl = item[lbl_key]
        base_name = re.sub(r'\s*\(~?[^\)]*(?:Mo|Go|MB|GB)[^\)]*\)', '', raw_lbl).strip()
        repo_id = item.get("repo") or item.get("repo_id") or item.get("hf_repo") or item.get("model_id") or ("bofenghuang/asr-wav2vec2-ctc-french" if mid == "wav2vec2-fr" else ("pyannote/speaker-diarization-community-1" if "pyannote" in mid or "diar" in mid else ""))
        fb_size = (item.get("size_fr") if is_fr else item.get("size_en")) or (re.search(r'~?[0-9.,]+\s*(?:Mo|Go|MB|GB)', raw_lbl).group(0) if re.search(r'~?[0-9.,]+\s*(?:Mo|Go|MB|GB)', raw_lbl) else ("~380 Mo" if (mid == "wav2vec2-fr" and is_fr) else ("~380 MB" if mid == "wav2vec2-fr" else ("~35 Mo" if is_fr else "~35 MB"))))
        fb_clean = fb_size.lstrip("~").strip()
        has_upd = mid in getattr(self, "updates_available", set())
        rem_bytes = getattr(self, "remote_sizes_bytes", {}).get(mid, 0)
        rem_str = format_size_human(rem_bytes, "fr" if is_fr else "en") if rem_bytes > 0 else fb_clean
        if installed:
            loc_str = format_size_human(get_local_model_size_bytes(repo_id), "fr" if is_fr else "en") or fb_clean
            sp = f" ({loc_str})" if loc_str else ""
            display_lbl = f"{base_name}{sp}" + ((f" — Version disponible en ligne (~{rem_str})" if is_fr else f" — Online update available (~{rem_str})") if has_upd else "")
        else:
            display_lbl = f"{base_name} (~{rem_str})" if rem_str else base_name
        w["name_lbl"].config(text=display_lbl, font=("Segoe UI", 9, "bold" if installed else "normal"))
        btn_act, btn_uninst = w["btn_act"], w["btn_uninst"]
        if btn_act:
            if installed:
                btn_dis = disable_actions or (not has_upd)
                btn_act.config(text=("🔄 Mettre à jour" if is_fr else "🔄 Update"), bg=(("#93C5FD" if locked_by_run else THEME_BLUE) if has_upd else "#E5E7EB"), fg=("#FFFFFF" if has_upd else "#9CA3AF"), activebackground=(THEME_BLUE_HOVER if has_upd else "#E5E7EB"), activeforeground=("#FFFFFF" if has_upd else "#9CA3AF"), disabledforeground=("#FFFFFF" if (has_upd and locked_by_run) else "#9CA3AF"), cursor=("arrow" if btn_dis else "hand2"), state=(tk.DISABLED if btn_dis else tk.NORMAL))
            else:
                btn_act.config(
                    text=("⬇ Installer" if is_fr else "⬇ Install"),
                    bg=("#93C5FD" if locked_by_run else THEME_BLUE),
                    fg="#FFFFFF",
                    activebackground=THEME_BLUE_HOVER,
                    activeforeground="#FFFFFF",
                    disabledforeground="#FFFFFF",
                    cursor=("arrow" if disable_actions else "hand2"),
                    state=(tk.DISABLED if disable_actions else tk.NORMAL)
                )
        if btn_uninst:
            if installed and not disable_actions:
                btn_uninst.config(state=tk.NORMAL, bg="#FEE2E2", fg="#991B1B", cursor="hand2")
            else:
                btn_uninst.config(state=tk.DISABLED, bg="#F3F4F6", fg="#9CA3AF", cursor="arrow")

    def _refresh_hf_token_badges(self):
        tok = get_hf_token(self.parent_app.config_data)
        has_tok = bool(tok and tok.startswith("hf_"))
        bg_default = self.cget("bg")
        for lbl in self.hf_badge_labels:
            if has_tok and "hf_ok" in self.hf_icons:
                lbl.config(image=self.hf_icons["hf_ok"], text="", bg=bg_default)
            elif not has_tok and "hf_missing" in self.hf_icons:
                lbl.config(image=self.hf_icons["hf_missing"], text="", bg=bg_default)
            else:
                lbl.config(text=("HF" if has_tok else "HF!"), fg=("#15803D" if has_tok else "#D97706"), bg=bg_default)

    def _on_configure_token_click(self):
        self.parent_app.prompt_and_save_token()
        self._refresh_hf_token_badges()

    def _ensure_valid_hf_token_for_gated(self) -> str:
        tok = get_hf_token(self.parent_app.config_data)
        if tok:
            status_info = check_hf_repos_access(tok)
            if status_info["status"] == "ok":
                self._refresh_hf_token_badges()
                return tok
            if status_info["status"] == "blocked":
                details = "\n".join([f"• {repo} : {st}" for repo, st in status_info["results"].items()])
                self.parent_app.print_hf_setup_procedure(
                    f"Votre jeton Hugging Face est enregistré, mais l'accès aux dépôts requis est refusé :\n{details}"
                )
        else:
            self.parent_app.print_hf_setup_procedure(
                "Le téléchargement ou la mise à jour des modèles d'identification des locuteurs (pyannote) nécessite un jeton Hugging Face valide."
            )

        if self.parent_app.prompt_and_save_token():
            self._refresh_hf_token_badges()
            new_tok = get_hf_token(self.parent_app.config_data)
            if new_tok:
                st2 = check_hf_repos_access(new_tok)
                if st2["status"] == "ok":
                    return new_tok
                if st2["status"] == "blocked":
                    details = "\n".join([f"• {repo} : {st}" for repo, st in st2["results"].items()])
                    self.parent_app.print_hf_setup_procedure(f"Accès toujours refusé sur Hugging Face :\n{details}")
        return ""

    def _kill_pid_tree(self, pid: int):
        if not pid:
            return
        try:
            if os.name == "nt":
                subprocess.Popen(
                    ["taskkill", "/F", "/T", "/PID", str(pid)],
                    stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                os.kill(pid, 15)
        except Exception:
            pass

    def _get_repos_downloaded_bytes(self, repo_dirs: list) -> int:
        total = 0
        for repo_dir in repo_dirs:
            if not repo_dir.exists():
                continue
            for f in repo_dir.rglob("*"):
                if f.is_file() and not f.is_symlink() and f.suffix != ".lock":
                    try:
                        total += f.stat().st_size
                    except OSError:
                        pass
        return total

    def start_model_task(self, item: dict, mode: str):
        import traceback
        mid = item["id"]
        write_model_debug(f"CLIC start_model_task(mid={mid}, mode={mode})")
        try:
            self._start_model_task_impl(item, mode)
        except Exception as exc:
            write_model_debug(f"EXCEPTION start_model_task({mid}):\n{traceback.format_exc()}")
            messagebox.showerror("Erreur interne", f"Erreur lors du lancement de {mid} :\n{exc}", parent=self)

    def _start_model_task_impl(self, item: dict, mode: str):
        mid = item["id"]
        if mid in self.active_tasks or getattr(self.parent_app, "is_running", False):
            return
        is_fr = (self.parent_app.ui_lang == "fr")
        hf_tok = ""
        if mode in ("install", "update") and item.get("gated"):
            hf_tok = self._ensure_valid_hf_token_for_gated()
            if not hf_tok:
                messagebox.showerror(
                    "Token Hugging Face requis" if is_fr else "Hugging Face Token Required",
                    "Un jeton Hugging Face valide (avec conditions acceptées sur les 2 dépôts pyannote) est requis avant d'installer ce modèle."
                    if is_fr else
                    "A valid Hugging Face token (with terms accepted on both pyannote repositories) is required before installing this model.",
                    parent=self
                )
                return

        if self.lbl_no_tasks.winfo_ismapped():
            self.lbl_no_tasks.pack_forget()

        prev_row = self.task_rows_by_mid.pop(mid, None)
        if prev_row and prev_row.winfo_exists():
            prev_row.destroy()

        task_row = ttk.Frame(self.tasks_inner, padding=(2, 3))
        task_row.pack(fill=tk.X, pady=2)
        self.task_rows_by_mid[mid] = task_row

        top_line = ttk.Frame(task_row)
        top_line.pack(fill=tk.X)

        if mode == "uninstall":
            init_title = f"🗑 Désinstallation : {mid}" if is_fr else f"🗑 Uninstalling: {mid}"
        elif mode == "update":
            init_title = f"🔄 Mise à jour : {mid}" if is_fr else f"🔄 Updating: {mid}"
        else:
            init_title = f"⬇ Installation : {mid}" if is_fr else f"⬇ Installing: {mid}"

        lbl_title = ttk.Label(top_line, text=init_title, font=("Segoe UI", 8, "bold"), foreground=(THEME_BLUE if mode != "uninstall" else "#D97706"))
        lbl_title.pack(side=tk.LEFT)

        btn_stop_task = tk.Button(
            top_line,
            text=("⏹ Arrêter" if is_fr else "⏹ Stop"),
            command=lambda m=mid: self._cancel_single_task(m),
            bg="#FEE2E2",
            fg="#991B1B",
            relief=tk.FLAT,
            font=("Segoe UI", 7, "bold"),
            padx=6,
            pady=0,
            cursor="hand2"
        )
        if mode != "uninstall":
            btn_stop_task.pack(side=tk.RIGHT, padx=(6, 0))

        lbl_pct = ttk.Label(top_line, text="0 %" if mode != "uninstall" else "100 %", font=("Segoe UI", 8, "bold"), foreground=(THEME_BLUE if mode != "uninstall" else "#D97706"))
        lbl_pct.pack(side=tk.RIGHT)

        bar_color = "#D97706" if mode == "uninstall" else THEME_BLUE
        prog = PillProgressBar(task_row, height=13, show_text=False, maximum=100, fill_color=bar_color)
        prog.pack(fill=tk.X, pady=(2, 1))
        if mode == "uninstall":
            prog["value"] = 100

        lbl_sub = ttk.Label(task_row, text="", foreground="#4B5563", font=("Segoe UI", 8))
        lbl_sub.pack(anchor=tk.W)

        cache_dir = get_hf_cache_path()
        cache_dir.mkdir(parents=True, exist_ok=True)
        repos = item.get("repos", [])
        repo_dirs = [cache_dir / f"models--{r.replace('/', '--')}" for r in repos]

        ctx = {
            "item": item,
            "mode": mode,
            "cancel": False,
            "was_installed": is_whisper_model_installed(mid),
            "repo_dirs": repo_dirs,
            "baseline_bytes": self._get_repos_downloaded_bytes(repo_dirs),
            "total_bytes": 0,
            "t_start": time.time(),
            "proc": None,
            "info_proc": None,
            "err_file": None,
            "info_file": None,
            "uninst_step": 0,
            "lbl_title": lbl_title,
            "lbl_pct": lbl_pct,
            "prog": prog,
            "lbl_sub": lbl_sub,
            "btn_stop": btn_stop_task,
        }

        if mode != "uninstall":
            runtime_py = RUNTIME_DIR / "python.exe"
            venv_py = SCRIPT_DIR / "whisperx-env" / "Scripts" / "python.exe"
            py_exe = runtime_py if runtime_py.exists() else (venv_py if venv_py.exists() else Path(sys.executable))
            env = os.environ.copy()
            env["HF_HUB_CACHE"] = str(cache_dir)
            env["HF_HUB_DISABLE_SYMLINKS"] = "1"
            env["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
            env["HF_HUB_OFFLINE"] = "0"
            env["TRANSFORMERS_OFFLINE"] = "0"
            env["HF_HUB_ENABLE_HF_TRANSFER"] = "0"
            env["HF_HUB_DISABLE_PROGRESS_BARS"] = "1"
            env["TQDM_DISABLE"] = "1"
            env["PYTHONUNBUFFERED"] = "1"
            pem_path = SCRIPT_DIR / "source" / "win_ca_bundle.pem"
            if not pem_path.exists():
                pem_path = SCRIPT_DIR / "win_ca_bundle.pem"
            if pem_path.exists():
                env["REQUESTS_CA_BUNDLE"] = str(pem_path)
                env["SSL_CERT_FILE"] = str(pem_path)
            if hf_tok:
                env["HF_TOKEN"] = hf_tok
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

            # 1) Sous-processus léger pour récupérer la taille totale (sans pipe bloquant)
            info_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
            info_tmp.close()
            ctx["info_file"] = Path(info_tmp.name)
            info_code = (
                "import os, sys; "
                "os.environ['HF_HUB_OFFLINE'] = '0'; "
                "from huggingface_hub import HfApi; api = HfApi(); "
                "tot = sum(sum(getattr(s, 'size', 0) or 0 for s in (api.model_info(r, files_metadata=True).siblings or [])) for r in sys.argv[1:]); "
                "print(tot)"
            )
            with open(ctx["info_file"], "w", encoding="utf-8") as f_out:
                ctx["info_proc"] = subprocess.Popen(
                    [str(py_exe), "-c", info_code, *repos],
                    stdin=subprocess.DEVNULL, stdout=f_out, stderr=subprocess.DEVNULL,
                    env=env, creationflags=creationflags
                )

            # 2) Sous-processus principal de téléchargement (sortie stderr vers fichier temp, mode en ligne explicite)
            err_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".log")
            err_tmp.close()
            ctx["err_file"] = Path(err_tmp.name)
            dl_code = (
                "import os, sys; "
                "os.environ['HF_HUB_OFFLINE'] = '0'; "
                "os.environ['TRANSFORMERS_OFFLINE'] = '0'; "
                "os.environ['HF_HUB_DISABLE_PROGRESS_BARS'] = '1'; "
                "import shutil; os.symlink = lambda s, d, *a, **k: shutil.copy2(os.path.join(os.path.dirname(d), s) if not os.path.isabs(s) else s, d); from huggingface_hub import snapshot_download; "
                "tok = os.environ.get('HF_TOKEN') or None; "
                "[snapshot_download(repo_id=r, cache_dir=sys.argv[1], token=tok, local_files_only=False) for r in sys.argv[2:]]"
            )
            with open(ctx["err_file"], "w", encoding="utf-8") as f_err:
                ctx["proc"] = subprocess.Popen(
                    [str(py_exe), "-c", dl_code, str(cache_dir), *repos],
                    stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=f_err,
                    env=env, creationflags=creationflags
                )
            lbl_sub.config(text=("Connexion à Hugging Face en cours..." if is_fr else "Connecting to Hugging Face..."))

        self.active_tasks[mid] = ctx
        self._refresh_single_row_state(mid)
        self._update_close_button_label()
        self._ensure_poller()

    def _cancel_single_task(self, mid: str):
        ctx = self.active_tasks.get(mid)
        if not ctx or ctx["cancel"]:
            return
        is_fr = (self.parent_app.ui_lang == "fr")
        ctx["cancel"] = True
        if ctx["btn_stop"].winfo_exists():
            ctx["btn_stop"].config(state=tk.DISABLED, text=("Arrêt..." if is_fr else "Stopping..."))
        ctx["lbl_title"].config(
            text=(f"⏹ Arrêt de {mid} en cours..." if is_fr else f"⏹ Stopping {mid}..."),
            foreground="#B45309"
        )
        if ctx.get("info_proc") and ctx["info_proc"].poll() is None:
            self._kill_pid_tree(ctx["info_proc"].pid)
        if ctx.get("proc") and ctx["proc"].poll() is None:
            self._kill_pid_tree(ctx["proc"].pid)

    def _ensure_poller(self):
        if not self._poller_active:
            self._poller_active = True
            self.after(120, self._poll_all_tasks)

    def _cleanup_temp_files(self, ctx: dict):
        for k in ("info_file", "err_file"):
            fp = ctx.get(k)
            if fp and fp.exists():
                try:
                    fp.unlink()
                except OSError:
                    pass

    def _poll_all_tasks(self):
        if not self.winfo_exists():
            return
        is_fr = (self.parent_app.ui_lang == "fr")
        completed = []

        for mid, ctx in list(self.active_tasks.items()):
            mode = ctx["mode"]

            if mode == "uninstall":
                step = ctx["uninst_step"] + 1
                ctx["uninst_step"] = step
                if step == 1:
                    ctx["prog"].config(value=70)
                    ctx["lbl_pct"].config(text="70 %")
                    ctx["lbl_sub"].config(text=("Suppression des fichiers du cache local..." if is_fr else "Deleting files from local cache..."))
                elif step == 2:
                    for rd in ctx["repo_dirs"]:
                        self._force_rmtree(rd)
                    ctx["prog"].config(value=30)
                    ctx["lbl_pct"].config(text="30 %")
                else:
                    ctx["prog"].config(value=0)
                    ctx["lbl_pct"].config(text="0 %")
                    ctx["lbl_title"].config(
                        text=(f"✔ {mid} désinstallé" if is_fr else f"✔ {mid} uninstalled"),
                        foreground="#15803D"
                    )
                    ctx["lbl_sub"].config(text=("Espace disque libéré." if is_fr else "Disk space freed."))
                    self.any_modified = True
                    completed.append((mid, False))
                continue

            # Mode install / update
            info_p = ctx.get("info_proc")
            if info_p and info_p.poll() is not None:
                if info_p.returncode == 0 and ctx.get("info_file") and ctx["info_file"].exists():
                    txt_sz = ctx["info_file"].read_text(encoding="utf-8", errors="ignore").strip()
                    if txt_sz.isdigit():
                        ctx["total_bytes"] = int(txt_sz)
                ctx["info_proc"] = None

            proc = ctx.get("proc")
            if ctx["cancel"]:
                if proc and proc.poll() is None:
                    self._kill_pid_tree(proc.pid)
                    continue
                for rd in ctx["repo_dirs"]:
                    if rd.exists():
                        for inc in rd.rglob("*.incomplete"):
                            try:
                                inc.unlink()
                            except OSError:
                                pass
                        if not ctx["was_installed"] and not is_whisper_model_installed(mid):
                            self._force_rmtree(rd)
                if ctx["btn_stop"].winfo_exists():
                    ctx["btn_stop"].pack_forget()
                ctx["prog"].config(value=0)
                ctx["lbl_pct"].config(text="")
                ctx["lbl_title"].config(
                    text=(f"⏹ {mid} : interrompu (acquis préservés)" if is_fr else f"⏹ {mid}: stopped (previous files preserved)"),
                    foreground="#B45309"
                )
                ctx["lbl_sub"].config(text="")
                self._cleanup_temp_files(ctx)
                completed.append((mid, None))
                continue

            if proc and proc.poll() is None:
                cur_bytes = self._get_repos_downloaded_bytes(ctx["repo_dirs"])
                dl_session = max(0, cur_bytes - ctx["baseline_bytes"])
                elapsed = max(0.2, time.time() - ctx["t_start"])
                speed_bps = dl_session / elapsed
                speed_mb = speed_bps / (1024 * 1024)
                total_bytes = ctx["total_bytes"]
                if total_bytes > 0:
                    pct = min(99.0, (cur_bytes / total_bytes) * 100.0)
                    cur_mb = cur_bytes / (1024 * 1024)
                    tot_mb = total_bytes / (1024 * 1024)
                    rem_bytes = max(0, total_bytes - cur_bytes)
                    eta_str = ("finalisation SHA256..." if rem_bytes == 0 else (format_elapsed_buzz(rem_bytes / speed_bps) if speed_bps > 10240 else "calcul..."))
                    metrics_txt = f"{cur_mb:,.1f} Mo / {tot_mb:,.1f} Mo  •  Débit : {speed_mb:.1f} Mo/s  •  Restant est. : {eta_str}".replace(",", " ")
                else:
                    pct = min(95.0, (dl_session / (1024 * 1024)) * 0.1)
                    cur_mb = cur_bytes / (1024 * 1024)
                    metrics_txt = f"{cur_mb:,.1f} Mo téléchargés  •  Débit : {speed_mb:.1f} Mo/s".replace(",", " ")
                ctx["prog"].config(value=pct)
                ctx["lbl_pct"].config(text=f"{int(pct)} %")
                ctx["lbl_sub"].config(text=metrics_txt)
            else:
                rc = proc.returncode if proc else 1
                if ctx["btn_stop"].winfo_exists():
                    ctx["btn_stop"].pack_forget()
                if rc == 0:
                    ctx["prog"].config(value=100)
                    ctx["lbl_pct"].config(text="100 %")
                    ctx["lbl_title"].config(
                        text=(f"✔ {mid} prêt dans le cache local" if is_fr else f"✔ {mid} ready in local cache"),
                        foreground="#15803D"
                    )
                    self.any_modified = True
                    self.updates_available.discard(mid)
                    self._cleanup_temp_files(ctx)
                    completed.append((mid, True))
                else:
                    err_msg = "Erreur réseau"
                    if ctx.get("err_file") and ctx["err_file"].exists():
                        raw_err = ctx["err_file"].read_text(encoding="utf-8", errors="ignore").strip()
                        write_model_debug(f"ECHEC {mid} (rc={rc}):\n{raw_err}")
                        lines = raw_err.splitlines()
                        if lines:
                            err_msg = lines[-1]
                    ctx["lbl_title"].config(text=f"❌ {mid} : échec", foreground="#B91C1C")
                    ctx["lbl_sub"].config(text=err_msg)
                    self._cleanup_temp_files(ctx)
                    completed.append((mid, None))

        for mid, inst_flag in completed:
            self.active_tasks.pop(mid, None)
            if inst_flag is not None:
                self._refresh_single_row_state(mid, force_installed=inst_flag)
            else:
                self._refresh_single_row_state(mid)
            self.parent_app.refresh_model_display()
            self._update_close_button_label()

        if self.active_tasks:
            self.after(200, self._poll_all_tasks)
        else:
            self._poller_active = False

    def _update_close_button_label(self):
        is_fr = (self.parent_app.ui_lang == "fr")
        if self.active_tasks:
            self.btn_close.config(state=tk.DISABLED, text=("Opérations en cours..." if is_fr else "Operations running..."))
        elif self.any_modified:
            self.btn_close.config(
                state=tk.NORMAL,
                bg=THEME_BLUE,
                fg="#FFFFFF",
                text=("✔ Appliquer & Redémarrer" if is_fr else "✔ Apply & Restart")
            )
        else:
            self.btn_close.config(
                state=tk.NORMAL,
                bg=THEME_BTN_SEC_BG,
                fg="#111827",
                text=("Fermer" if is_fr else "Close")
            )

    def _start_remote_update_check(self):
        if not self.winfo_exists():
            return
        cache_dir = get_hf_cache_path()
        installed_pairs = []
        for item in WHISPER_MODEL_CATALOG:
            mid = item["id"]
            if not item.get("available") or not is_whisper_model_installed(mid):
                continue
            for repo in item.get("repos", []):
                ref_file = cache_dir / f"models--{repo.replace('/', '--')}" / "refs" / "main"
                if ref_file.exists():
                    try:
                        local_sha = ref_file.read_text(encoding="utf-8").strip()
                        if local_sha:
                            installed_pairs.append(f"{mid}|{repo}|{local_sha}")
                    except OSError:
                        pass
        if not installed_pairs:
            return

        runtime_py = RUNTIME_DIR / "python.exe"
        venv_py = SCRIPT_DIR / "whisperx-env" / "Scripts" / "python.exe"
        py_exe = runtime_py if runtime_py.exists() else (venv_py if venv_py.exists() else Path(sys.executable))
        if not py_exe.exists():
            return

        env = os.environ.copy()
        env["HF_HUB_OFFLINE"] = "0"
        env["TRANSFORMERS_OFFLINE"] = "0"
        pem_path = SCRIPT_DIR / "source" / "win_ca_bundle.pem"
        if not pem_path.exists():
            pem_path = SCRIPT_DIR / "win_ca_bundle.pem"
        if pem_path.exists():
            env["REQUESTS_CA_BUNDLE"] = str(pem_path)
            env["SSL_CERT_FILE"] = str(pem_path)
        hf_tok = get_hf_token(self.parent_app.config_data)
        if hf_tok:
            env["HF_TOKEN"] = hf_tok

        chk_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
        chk_tmp.close()
        self._update_check_file = Path(chk_tmp.name)

        chk_code = (
            "import os, sys; os.environ['HF_HUB_OFFLINE'] = '0'; "
            "from huggingface_hub import HfApi; api = HfApi(); tok = os.environ.get('HF_TOKEN') or None; "
            "out = [e.split('|')[0] for e in sys.argv[1:] if getattr(api.model_info(e.split('|')[1], token=tok), 'sha', None) not in (None, e.split('|')[2])]; "
            "print(','.join(sorted(set(out))))"
        )
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        with open(self._update_check_file, "w", encoding="utf-8") as f_out:
            self._update_check_proc = subprocess.Popen(
                [str(py_exe), "-c", chk_code, *installed_pairs],
                stdin=subprocess.DEVNULL, stdout=f_out, stderr=subprocess.DEVNULL,
                env=env, creationflags=creationflags
            )
        self.after(250, self._poll_remote_update_check)

    def _poll_remote_update_check(self):
        if not self.winfo_exists():
            return
        proc = self._update_check_proc
        if not proc:
            return
        if proc.poll() is None:
            self.after(250, self._poll_remote_update_check)
            return
        try:
            if proc.returncode == 0 and self._update_check_file and self._update_check_file.exists():
                raw = self._update_check_file.read_text(encoding="utf-8", errors="ignore").strip()
                if raw:
                    found_mids = [m.strip() for m in raw.split(",") if m.strip()]
                    self.updates_available.update(found_mids)
                    for mid in found_mids:
                        self._refresh_single_row_state(mid)
        finally:
            if self._update_check_file and self._update_check_file.exists():
                try:
                    self._update_check_file.unlink()
                except OSError:
                    pass
            self._update_check_proc = None
            self._update_check_file = None

    def on_close(self):
        if self.active_tasks:
            return
        if self._update_check_proc and self._update_check_proc.poll() is None:
            self._kill_pid_tree(self._update_check_proc.pid)
        if self._update_check_file and self._update_check_file.exists():
            try:
                self._update_check_file.unlink()
            except OSError:
                pass
        if self.any_modified:
            self.parent_app.restart_application()
        else:
            self.destroy()

    def _start_remote_check(self):
        import threading
        self._remote_meta = {}
        def _worker():
            items = [
                ("tiny", "Systran/faster-whisper-tiny"),
                ("tiny.en", "Systran/faster-whisper-tiny.en"),
                ("small", "Systran/faster-whisper-small"),
                ("small.en", "Systran/faster-whisper-small.en"),
                ("medium", "Systran/faster-whisper-medium"),
                ("medium.en", "Systran/faster-whisper-medium.en"),
                ("large-v3", "Systran/faster-whisper-large-v3"),
                ("pyannote", "pyannote/speaker-diarization-community-1"),
                ("wav2vec2-fr", "bofenghuang/asr-wav2vec2-ctc-french"),
            ]
            for key, repo in items:
                sha, r_bytes = fetch_remote_model_meta(repo)
                local_sha = None
                ref_file = HF_CACHE_DIR / f"models--{repo.replace('/', '--')}" / "refs" / "main"
                if ref_file.exists():
                    try:
                        local_sha = ref_file.read_text(encoding="utf-8").strip()
                    except Exception:
                        pass
                has_update = bool(sha and local_sha and sha != local_sha)
                self._remote_meta[key] = (has_update, r_bytes)
            try:
                from PySide6.QtCore import QTimer
                QTimer.singleShot(0, self._apply_dynamic_labels)
            except Exception:
                pass
        threading.Thread(target=_worker, daemon=True).start()

    def _apply_dynamic_labels(self):
        is_fr = (getattr(self, "lang", "fr") == "fr")
        repo_map = {
            "tiny": ("tiny", "Systran/faster-whisper-tiny", "75 Mo"),
            "tiny.en": ("tiny.en", "Systran/faster-whisper-tiny.en", "75 Mo"),
            "small": ("small", "Systran/faster-whisper-small", "464 Mo"),
            "small.en": ("small.en", "Systran/faster-whisper-small.en", "464 Mo"),
            "medium": ("medium", "Systran/faster-whisper-medium", "1,46 Go"),
            "medium.en": ("medium.en", "Systran/faster-whisper-medium.en", "1,46 Go"),
            "large-v3": ("large (large-v3)", "Systran/faster-whisper-large-v3", "2,90 Go"),
            "pyannote": ("Identification des locuteurs (pyannote)" if is_fr else "Speaker identification (pyannote)", "pyannote/speaker-diarization-community-1", "35 Mo"),
            "wav2vec2-fr": ("Alignement audio français — rapide" if is_fr else "French audio alignment — fast", "bofenghuang/asr-wav2vec2-ctc-french", "380 Mo"),
        }
        rows_dict = getattr(self, "model_rows", None) or getattr(self, "_rows", None) or {}
        for key, (base_lbl, repo, fb_sz) in repo_map.items():
            row = rows_dict.get(key)
            if not row:
                continue
            lbl_w = row.get("label") if isinstance(row, dict) else getattr(row, "label", None)
            btn_w = row.get("install_btn") or row.get("action_btn") if isinstance(row, dict) else getattr(row, "action_btn", None)
            installed = is_hf_model_cached(repo)
            loc_bytes = get_local_model_size_bytes(repo) if installed else 0
            loc_str = format_size_human(loc_bytes, "fr" if is_fr else "en") or fb_sz
            has_upd, rem_bytes = getattr(self, "_remote_meta", {}).get(key, (False, 0))
            rem_str = format_size_human(rem_bytes, "fr" if is_fr else "en") or fb_sz

            if lbl_w is not None:
                if installed:
                    if has_upd:
                        lbl_w.setText(f"{base_lbl} ({loc_str}) — Version disponible en ligne (~{rem_str})" if is_fr else f"{base_lbl} ({loc_str}) — Online update available (~{rem_str})")
                    else:
                        lbl_w.setText(f"{base_lbl} ({loc_str})")
                else:
                    lbl_w.setText(f"{base_lbl} (~{rem_str})")

            if btn_w is not None and installed:
                btn_w.setText("Mettre à jour" if is_fr else "Update")
                if has_upd:
                    btn_w.setEnabled(True)
                    btn_w.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold;")
                else:
                    btn_w.setEnabled(False)
                    btn_w.setStyleSheet("")


class HFTokenDialog(tk.Toplevel):
    def __init__(self, parent, current_token: str = ""):
        super().__init__(parent)
        self.title("Configuration Hugging Face (HF_TOKEN & 2 dépôts requis)")
        self.geometry("620x340")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.saved_token = None
        self.existing_token = current_token.strip()
        self.showing_masked_preview = bool(self.existing_token)

        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Prérequis Hugging Face pour l'identification des locuteurs (pyannote)", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 8))
        info = (
            "1. Acceptez les conditions d'utilisation sur LES DEUX dépôts Hugging Face ci-dessous :\n"
            "   • pyannote/speaker-diarization-community-1\n"
            "   • pyannote/segmentation-3.0\n"
            "2. Créez un jeton de type 'Read' (hf_...) avec le même compte et collez-le ci-dessous."
        )
        ttk.Label(frame, text=info, justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 10))

        link_row1 = ttk.Frame(frame)
        link_row1.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(link_row1, text="1a. Dépôt speaker-diarization-community-1", command=lambda: webbrowser.open("https://huggingface.co/pyannote/speaker-diarization-community-1")).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(link_row1, text="1b. Dépôt segmentation-3.0", command=lambda: webbrowser.open("https://huggingface.co/pyannote/segmentation-3.0")).pack(side=tk.LEFT)

        link_row2 = ttk.Frame(frame)
        link_row2.pack(fill=tk.X, pady=(0, 12))
        ttk.Button(link_row2, text="2. Créer / Copier mon jeton (Read)", command=lambda: webbrowser.open("https://huggingface.co/settings/tokens")).pack(side=tk.LEFT)

        ttk.Label(frame, text="Votre jeton HF_TOKEN (hf_...) :").pack(anchor=tk.W)
        initial_display = self._format_masked(self.existing_token) if self.existing_token else ""
        self.var_token = tk.StringVar(value=initial_display)
        self.entry_token = ttk.Entry(frame, textvariable=self.var_token, show="", width=74)
        self.entry_token.pack(fill=tk.X, pady=(4, 14))
        self.entry_token.bind("<FocusIn>", self._on_entry_focus_in)
        self.entry_token.bind("<Key>", self._on_entry_key)

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)
        tk.Button(
            btn_row, text="💾 Vérifier & Enregistrer", command=self.on_save,
            bg=THEME_BLUE, fg="#FFFFFF", activebackground=THEME_BLUE_HOVER, activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=12, pady=4, cursor="hand2"
        ).pack(side=tk.RIGHT, padx=(8, 0))
        tk.Button(
            btn_row, text="Annuler", command=self.destroy,
            bg=THEME_BTN_SEC_BG, fg="#111827", activebackground=THEME_BTN_SEC_ACTIVE, activeforeground="#111827",
            font=("Segoe UI", 9), relief=tk.FLAT, padx=12, pady=4, cursor="hand2"
        ).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda e: self.on_save())
        self.bind("<Escape>", lambda e: self.destroy())

    def _format_masked(self, tok: str) -> str:
        if len(tok) <= 5:
            return tok
        return tok[:5] + ("•" * (len(tok) - 5))

    def _on_entry_focus_in(self, event=None):
        if self.showing_masked_preview:
            self.showing_masked_preview = False
            self.var_token.set("")
            self.entry_token.config(show="•")

    def _on_entry_key(self, event=None):
        if self.showing_masked_preview:
            self.showing_masked_preview = False
            self.var_token.set("")
            self.entry_token.config(show="•")

    def on_save(self):
        if self.showing_masked_preview and self.existing_token:
            self.saved_token = self.existing_token
            self.destroy()
            return
        val = self.var_token.get().strip()
        if not val.startswith("hf_") or len(val) < 20 or "•" in val:
            messagebox.showwarning("Format de jeton invalide", "Un jeton Hugging Face valide doit commencer par 'hf_' et comporter au moins 20 caractères.", parent=self)
            return
        self.saved_token = val
        self.destroy()

def enable_win32_drag_and_drop(widget, callback):
    """Active le glisser-déposer natif Windows (WM_DROPFILES) sur un widget Tkinter sans dépendance externe."""
    if os.name != "nt":
        return
    try:
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32

        GWL_WNDPROC = -4
        WM_DROPFILES = 0x0233
        LRESULT = ctypes.c_ssize_t
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

        if ctypes.sizeof(ctypes.c_void_p) == 8:
            SetWindowLongPtr = user32.SetWindowLongPtrW
            SetWindowLongPtr.argtypes = [wintypes.HWND, ctypes.c_int, WNDPROC]
            SetWindowLongPtr.restype = ctypes.c_void_p
            CallWindowProc = user32.CallWindowProcW
            CallWindowProc.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            CallWindowProc.restype = LRESULT
        else:
            SetWindowLongPtr = user32.SetWindowLongW
            SetWindowLongPtr.argtypes = [wintypes.HWND, ctypes.c_int, WNDPROC]
            SetWindowLongPtr.restype = ctypes.c_void_p
            CallWindowProc = user32.CallWindowProcW
            CallWindowProc.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            CallWindowProc.restype = LRESULT

        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragFinish.argtypes = [wintypes.HANDLE]

        widget.update_idletasks()
        hwnd = widget.winfo_id()
        shell32.DragAcceptFiles(hwnd, True)

        old_proc = [None]

        def py_wndproc(h, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                hdrop = wintypes.HANDLE(wparam)
                count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
                dropped_paths = []
                for i in range(count):
                    length = shell32.DragQueryFileW(hdrop, i, None, 0)
                    buf = ctypes.create_unicode_buffer(length + 1)
                    shell32.DragQueryFileW(hdrop, i, buf, length + 1)
                    dropped_paths.append(buf.value)
                shell32.DragFinish(hdrop)
                if dropped_paths:
                    widget.after(10, lambda paths=dropped_paths: callback(paths))
                return 0
            return CallWindowProc(old_proc[0], h, msg, wparam, lparam)

        proc_cb = WNDPROC(py_wndproc)
        old_proc[0] = SetWindowLongPtr(hwnd, GWL_WNDPROC, proc_cb)
        widget._win32_dnd_ref = (proc_cb, old_proc)
    except Exception:
        pass

class SlideScriptApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("1120x860")
        self.minsize(960, 720)

        self.selected_files = []
        self.current_media_path = None
        self.queue_item_ids = {}
        self.cancelled_files = set()
        self.active_processed_paths = set()
        self._drag_start_iid = None
        self.media_durations = {}
        self.media_has_video = {}
        self.batch_start_time = None
        self.live_speed_multiplier = 1.0
        self.hf_token_ok = False
        self.current_process = None
        self.is_running = False
        self.is_paused = False
        self.stop_requested = False
        self.inline_log_active = False
        self.has_cuda_gpu = False
        self.ui_icons = {}
        self.last_effective_model = None
        self.current_step_key = "step_wait"
        self.current_step_kwargs = {}

        self.config_data = self.load_config()
        self.ui_lang = self.config_data.get("ui_lang", "fr")
        if self.ui_lang not in I18N:
            self.ui_lang = "fr"

        self.build_ui()
        self.apply_ui_language()
        self.after(150, lambda: self._apply_window_icon(self))
        self.after(250, lambda: enable_win32_drag_and_drop(self.frame_queue, self.on_files_dropped))
        self.after(250, lambda: enable_win32_drag_and_drop(self.tree_queue, self.on_files_dropped))
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.after(200, self.startup_checks)

    def _apply_window_icon(self, win):
        try:
            ico_candidates = [
                Path(__file__).resolve().parent / "SlideScript.ico",
                SCRIPT_DIR / "SlideScript.ico",
                SCRIPT_DIR / "source" / "SlideScript.ico",
                Path(sys.executable).resolve().parent / "SlideScript.ico",
            ]
            for ico in ico_candidates:
                if ico.exists():
                    win.iconbitmap(default=str(ico))
                    win.iconbitmap(str(ico))
                    return
        except Exception:
            pass

    
    def tr(self, key: str, **kwargs) -> str:
        text = I18N[self.ui_lang].get(key, I18N["fr"].get(key, key))
        return text.format(**kwargs) if kwargs else text

    def startup_checks(self):
        def _hw_check():
            cuda_ok = detect_nvidia_cuda()
            def _apply_hw():
                self.has_cuda_gpu = cuda_ok
                if cuda_ok:
                    self.lbl_hw_status.config(text=self.tr("hw_cuda"), foreground="#FFFFFF")
                    self.cpu_frame.pack_forget()
                    self.lbl_cpu_gpu_hint.config(text=self.tr("cpu_hidden_gpu"))
                    self.lbl_cpu_gpu_hint.pack(side=tk.RIGHT)
                else:
                    self.lbl_hw_status.config(text=self.tr("hw_cpu"), foreground="#E5E7EB")
                    self.lbl_cpu_gpu_hint.pack_forget()
                    self.cpu_frame.pack(side=tk.LEFT)
            self.after(0, _apply_hw)
        threading.Thread(target=_hw_check, daemon=True).start()


        self.run_preflight_check(verbose=False)

    def set_hf_button_alert(self, is_alert: bool):
        def _apply():
            self.hf_token_ok = not is_alert
            base_txt = self.tr("btn_token")
            self.btn_token.config(
                text=f" {base_txt} (OK)" if self.hf_token_ok else f" {base_txt}",
                bg="#B91C1C" if is_alert else "#E5E7EB",
                fg="#FFFFFF" if is_alert else "#111827",
                activebackground="#991B1B" if is_alert else "#D1D5DB",
                activeforeground="#FFFFFF" if is_alert else "#111827"
            )
        self.after(0, _apply)

    def print_hf_setup_procedure(self, reason: str):
        self.print_diagnostic_card(
            "Configuration requise du token Hugging Face (identification des locuteurs)",
            "Prérequis Hugging Face (HF_TOKEN & 2 dépôts pyannote)",
            reason,
            "Procédure exacte à exécuter :\n"
            "1. Créez un compte gratuit ou connectez-vous sur https://huggingface.co\n"
            "2. Ouvrez LES DEUX dépôts suivants et cliquez sur 'Agree and access repository' sur chacun :\n"
            "   • https://huggingface.co/pyannote/speaker-diarization-community-1\n"
            "   • https://huggingface.co/pyannote/segmentation-3.0\n"
            "3. Allez sur https://huggingface.co/settings/tokens → 'Create new token' (type 'Read') et copiez la clé (hf_...).\n"
            "4. Cliquez sur le bouton '🔑 Configurer token Hugging Face' en haut à droite, collez la clé et validez."
        )

    def prompt_and_save_token(self) -> bool:
        current = get_hf_token(self.config_data)
        dlg = HFTokenDialog(self, current_token=current)
        self.wait_window(dlg)
        if dlg.saved_token:
            save_hf_token(dlg.saved_token, self.config_data)
            self.save_config()
            masked = f"{dlg.saved_token[:5]}..."
            self.log(f"[OK] Jeton Hugging Face ({masked}) enregistré. Vérification des accès...\n", "success")
            self.verify_hf_access_in_background(dlg.saved_token, announce_ok=True)
            return True
        if not get_hf_token(self.config_data):
            self.set_hf_button_alert(True)
        return bool(get_hf_token(self.config_data))

    def verify_hf_access_in_background(self, token: str, announce_ok: bool = True):
        def _check():
            status_info = check_hf_repos_access(token)
            if status_info["status"] == "blocked":
                self.set_hf_button_alert(True)
                details = "\n".join([f"• {repo} : {st}" for repo, st in status_info["results"].items()])
                self.print_hf_setup_procedure(
                    f"Votre jeton Hugging Face est enregistré, mais l'accès aux modèles requis est refusé :\n{details}"
                )
            elif status_info["status"] == "ok":
                self.set_hf_button_alert(False)
                if announce_ok:
                    self.log("[OK] Jeton valide : accès confirmé aux 2 dépôts Hugging Face requis.\n", "success")
            else:
                hf_cache = get_hf_cache_path()
                diar_cached = (hf_cache / "models--pyannote--speaker-diarization-community-1").exists()
                self.set_hf_button_alert(not diar_cached)
        threading.Thread(target=_check, daemon=True).start()

    def load_config(self):
        defaults = {
            "ui_lang": "fr",
            "out_txt": False,
            "out_docx": True,
            "out_docx_slides": False,
            "out_pptx": False,
            "langue": "en",
            "modele": "tiny",
            "sous_dossiers": False,
            "offline": True,
            "diarize": True,
            "cpu_threads": max(1, (os.cpu_count() or 8) // 2),
            "last_dir": ""
        }
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    loaded.pop("hf_token", None)
                    defaults.update(loaded)
            except Exception:
                pass
        return defaults

    def save_config(self):
        self.config_data["ui_lang"] = self.ui_lang
        self.config_data["out_txt"] = self.var_out_txt.get()
        self.config_data["out_docx"] = self.var_out_docx.get()
        self.config_data["out_docx_slides"] = self.var_out_docx_slides.get()
        self.config_data["out_pptx"] = self.var_out_pptx.get()
        self.config_data["langue"] = self.var_langue.get().strip() or "en"
        self.config_data["modele"] = self.var_modele.get()
        self.config_data["sous_dossiers"] = self.var_sous_dossiers.get()
        self.config_data["offline"] = True
        self.config_data["diarize"] = self.var_diarize.get()
        self.config_data.pop("hf_token", None)
        try:
            self.config_data["cpu_threads"] = int(self.var_cpu_threads.get())
        except ValueError:
            self.config_data["cpu_threads"] = max(1, (os.cpu_count() or 8) // 2)
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config_data, f, indent=2)
        except Exception:
            pass

    def build_output_filepath(self, media: Path, model_tag: str, with_slides: bool, ext: str, diarized: bool = None) -> Path:
        use_diar = self.var_diarize.get() if diarized is None else diarized
        tags = [model_tag]
        if use_diar:
            tags.append("speakers")
        if with_slides:
            tags.append("slides")
        suffix = "_".join(tags)
        return media.parent / f"{media.stem}_{suffix}.{ext}"

    def get_expected_outputs(self, media: Path, model_tag: str = None) -> list:
        m = model_tag or self.get_effective_model()
        outputs = []
        if self.var_out_txt.get():
            outputs.append(self.build_output_filepath(media, m, with_slides=False, ext="txt"))
        if self.var_out_docx.get():
            outputs.append(self.build_output_filepath(media, m, with_slides=False, ext="docx"))
        if self.var_out_docx_slides.get():
            outputs.append(self.build_output_filepath(media, m, with_slides=True, ext="docx"))
        if self.var_out_pptx.get():
            outputs.append(self.build_output_filepath(media, m, with_slides=True, ext="pptx"))
        return outputs

    def are_all_selected_outputs_done(self, media: Path) -> bool:
        expected = self.get_expected_outputs(media)
        if not expected:
            return False
        is_video = media.suffix.lower() in {".mp4", ".mkv", ".mov", ".avi"}
        check_list = [p for p in expected if is_video or p.suffix.lower() != ".pptx"]
        return bool(check_list) and all(p.exists() for p in check_list)

    def build_ui(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")

        style.configure("Buzz.Treeview", rowheight=22, font=("Segoe UI", 9))
        style.configure("Buzz.Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Missing.TCheckbutton", foreground="#6B7280", font=("Segoe UI", 9, "italic"))

        self.ui_icons = create_button_icons(self)

        top_bar = tk.Frame(self, bg=THEME_TOP_BG, padx=12, pady=7)
        top_bar.pack(fill=tk.X)



        self.lbl_hw_status = tk.Label(
            top_bar, text="", bg=THEME_TOP_BG, fg="#FFFFFF",
            font=("Segoe UI", 9, "bold")
        )
        self.lbl_hw_status.pack(side=tk.LEFT)

        lang_box = tk.Frame(top_bar, bg=THEME_TOP_BG)
        lang_box.pack(side=tk.RIGHT)

        self.btn_models_mgr = tk.Button(
            lang_box, command=self.open_model_manager, width=30,
            bg="#E5E7EB", fg="#111827", activebackground="#D1D5DB", activeforeground="#111827",
            font=("Segoe UI", 8, "bold"), relief=tk.FLAT, padx=8, pady=3, cursor="hand2"
        )
        self.btn_models_mgr.pack(side=tk.LEFT, padx=(0, 12))

        self.lbl_ui_lang = tk.Label(lang_box, text="Interface :", width=10, anchor=tk.E, bg=THEME_TOP_BG, fg="#E5E7EB", font=("Segoe UI", 8, "bold"))
        self.lbl_ui_lang.pack(side=tk.LEFT, padx=(0, 4))

        self.var_ui_lang = tk.StringVar(value=self.ui_lang.upper())
        cb_ui_lang = ttk.Combobox(lang_box, textvariable=self.var_ui_lang, values=["FR", "EN"], width=4, state="readonly", font=("Segoe UI", 8))
        cb_ui_lang.pack(side=tk.LEFT)
        cb_ui_lang.bind("<<ComboboxSelected>>", self.on_ui_lang_changed)

        main_frame = ttk.Frame(self, padding=(12, 8, 12, 12))
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.frame_src = ttk.LabelFrame(main_frame, padding=8)
        self.frame_src.pack(fill=tk.X, pady=(0, 6))

        btn_row = ttk.Frame(self.frame_src)
        btn_row.pack(fill=tk.X)

        self.btn_1file = ttk.Button(btn_row, command=self.select_single_file, image=self.ui_icons.get("file"), compound=tk.LEFT, width=19)
        self.btn_1file.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_mfiles = ttk.Button(btn_row, command=self.select_multiple_files, image=self.ui_icons.get("files"), compound=tk.LEFT, width=25)
        self.btn_mfiles.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_folder = ttk.Button(btn_row, command=self.select_folder, image=self.ui_icons.get("folder"), compound=tk.LEFT, width=19)
        self.btn_folder.pack(side=tk.LEFT, padx=(0, 12))

        self.var_sous_dossiers = tk.BooleanVar(value=self.config_data["sous_dossiers"])
        self.chk_sub = ttk.Checkbutton(btn_row, variable=self.var_sous_dossiers, command=self.save_config, width=24)
        self.chk_sub.pack(side=tk.LEFT)


        sel_info_row = ttk.Frame(self.frame_src)
        sel_info_row.pack(fill=tk.X, pady=(5, 0))

        self.lbl_selection = ttk.Label(sel_info_row, foreground="#444444")
        self.lbl_selection.pack(side=tk.LEFT)

        self.frame_opt = ttk.LabelFrame(main_frame, padding=8)
        self.frame_opt.pack(fill=tk.X, pady=(0, 6))

        opt_row_fmt = ttk.Frame(self.frame_opt)
        opt_row_fmt.pack(fill=tk.X, pady=(0, 6))

        self.lbl_fmt = ttk.Label(opt_row_fmt, width=17, anchor=tk.W)
        self.lbl_fmt.pack(side=tk.LEFT)

        self.var_out_txt = tk.BooleanVar(value=self.config_data["out_txt"])
        self.chk_out_txt = ttk.Checkbutton(opt_row_fmt, variable=self.var_out_txt, command=self.on_output_format_changed, width=7)
        self.chk_out_txt.pack(side=tk.LEFT, padx=(0, 16))

        self.var_out_docx = tk.BooleanVar(value=self.config_data["out_docx"])
        self.chk_out_docx = ttk.Checkbutton(opt_row_fmt, variable=self.var_out_docx, command=self.on_output_format_changed, width=9)
        self.chk_out_docx.pack(side=tk.LEFT, padx=(0, 16))

        self.var_out_docx_slides = tk.BooleanVar(value=self.config_data["out_docx_slides"])
        self.chk_out_docx_slides = ttk.Checkbutton(opt_row_fmt, variable=self.var_out_docx_slides, command=self.on_output_format_changed, width=18)
        self.chk_out_docx_slides.pack(side=tk.LEFT, padx=(0, 16))

        self.var_out_pptx = tk.BooleanVar(value=self.config_data["out_pptx"])
        self.chk_out_pptx = ttk.Checkbutton(opt_row_fmt, variable=self.var_out_pptx, command=self.on_output_format_changed, width=36)
        self.chk_out_pptx.pack(side=tk.LEFT)

        self.lbl_cpu_gpu_hint = ttk.Label(opt_row_fmt, text="", foreground="#4B5563", font=("Segoe UI", 8))

        opt_row1 = ttk.Frame(self.frame_opt)
        opt_row1.pack(fill=tk.X, pady=(0, 6))

        self.lbl_lang = ttk.Label(opt_row1, width=17, anchor=tk.W)
        self.lbl_lang.pack(side=tk.LEFT)
        self.var_langue = tk.StringVar(value=self.config_data["langue"])
        cb_langue = ttk.Combobox(opt_row1, textvariable=self.var_langue, values=["auto", "en", "fr", "es", "de", "it"], width=7)
        cb_langue.pack(side=tk.LEFT, padx=(0, 24))
        cb_langue.bind("<<ComboboxSelected>>", self.on_language_changed)
        cb_langue.bind("<FocusOut>", self.on_language_changed)

        self.lbl_model = ttk.Label(opt_row1, width=15, anchor=tk.W)
        self.lbl_model.pack(side=tk.LEFT)
        avail_models = get_available_base_models()
        saved_m = self.config_data.get("modele", "tiny")
        if saved_m.endswith(".en"):
            saved_m = saved_m[:-3]
        if saved_m not in avail_models:
            saved_m = "small" if "small" in avail_models else avail_models[0]
        self.var_modele = tk.StringVar(value=saved_m)
        cb_modele = ttk.Combobox(opt_row1, textvariable=self.var_modele, values=avail_models, width=10, state="readonly")
        cb_modele.pack(side=tk.LEFT)
        cb_modele.bind("<<ComboboxSelected>>", self.on_model_changed)

        self.lbl_model_effective = ttk.Label(opt_row1, text="", width=58, anchor=tk.W, foreground=THEME_BLUE, font=("Segoe UI", 8, "bold"))
        self.lbl_model_effective.pack(side=tk.LEFT, padx=(10, 0))

        opt_row2 = ttk.Frame(self.frame_opt)
        opt_row2.pack(fill=tk.X)



        self.var_diarize = tk.BooleanVar(value=self.config_data.get("diarize", True))
        self.chk_diarize = ttk.Checkbutton(opt_row2, variable=self.var_diarize, command=self.on_diarize_changed, width=22)
        self.chk_diarize.pack(side=tk.LEFT, padx=(0, 4))
        self.lbl_diarize_warn = ttk.Label(opt_row2, text="", foreground="#6B7280", font=("Segoe UI", 8, "italic"))
        self.lbl_diarize_warn.pack(side=tk.LEFT, padx=(0, 12))

        self.cpu_frame = ttk.Frame(opt_row2)
        self.cpu_frame.pack(side=tk.LEFT)

        self.lbl_cpu_cores = ttk.Label(self.cpu_frame, width=11, anchor=tk.W)
        self.lbl_cpu_cores.pack(side=tk.LEFT)

        total_cores = os.cpu_count() or 8
        half_cores = max(1, total_cores // 2)
        core_options = [str(c) for c in range(half_cores, total_cores + 1)]
        saved_cores = str(self.config_data.get("cpu_threads", half_cores))
        if saved_cores not in core_options:
            saved_cores = str(half_cores)

        self.var_cpu_threads = tk.StringVar(value=saved_cores)
        cb_cores = ttk.Combobox(self.cpu_frame, textvariable=self.var_cpu_threads, values=core_options, width=4, state="readonly")
        cb_cores.pack(side=tk.LEFT)
        cb_cores.bind("<<ComboboxSelected>>", self.on_cpu_cores_changed)

        self.lbl_cpu_warn = ttk.Label(self.cpu_frame, text="", width=54, anchor=tk.W, font=("Segoe UI", 8, "bold"))
        self.lbl_cpu_warn.pack(side=tk.LEFT, padx=(10, 0))

        frame_ctrl = ttk.Frame(main_frame)
        frame_ctrl.pack(fill=tk.X, pady=(0, 6))

        self.btn_start = tk.Button(
            frame_ctrl, command=self.start_processing, state=tk.NORMAL,
            image=self.ui_icons.get("play"),
            compound=tk.LEFT, width=255,
            bg=THEME_BTN_PRIMARY_BG, fg="#FFFFFF",
            activebackground=THEME_BTN_PRIMARY_ACTIVE, activeforeground="#FFFFFF",
            disabledforeground="#FFFFFF", font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 8))

        sec_btn_opts = dict(
            bg=THEME_BTN_SEC_BG, fg="#111827",
            activebackground=THEME_BTN_SEC_ACTIVE, activeforeground="#111827",
            disabledforeground="#6B7280", font=("Segoe UI", 9),
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        )

        self.btn_pause = tk.Button(
            frame_ctrl, command=self.toggle_pause, state=tk.DISABLED,
            image=self.ui_icons.get("pause"), compound=tk.LEFT, width=105, **sec_btn_opts
        )
        self.btn_pause.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_stop = tk.Button(
            frame_ctrl, command=self.stop_processing, state=tk.DISABLED,
            image=self.ui_icons.get("stop"), compound=tk.LEFT, width=105, **sec_btn_opts
        )
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_diag = tk.Button(
            frame_ctrl, command=lambda: self.run_preflight_check(verbose=True),
            image=self.ui_icons.get("diag"), compound=tk.LEFT, width=165, **sec_btn_opts
        )
        self.btn_diag.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_clear = tk.Button(
            frame_ctrl, command=self.clear_log,
            image=self.ui_icons.get("clear"), compound=tk.LEFT, width=155, **sec_btn_opts
        )
        self.btn_clear.pack(side=tk.LEFT)

        self.frame_prog = ttk.LabelFrame(main_frame, padding=8)
        self.frame_prog.pack(fill=tk.X, pady=(0, 6))

        row_tot = ttk.Frame(self.frame_prog)
        row_tot.pack(fill=tk.X, pady=(0, 3))
        self.lbl_tot_prog_title = ttk.Label(row_tot, font=("Segoe UI", 9, "bold"))
        self.lbl_tot_prog_title.pack(side=tk.LEFT)
        self.lbl_total_progress = ttk.Label(row_tot, font=("Segoe UI", 9, "bold"), foreground="#1F2937")
        self.lbl_total_progress.pack(side=tk.RIGHT)

        self.progress_total = PillProgressBar(self.frame_prog, height=16, show_text=False)
        self.progress_total.pack(fill=tk.X, pady=(0, 6))

        row_cur = ttk.Frame(self.frame_prog)
        row_cur.pack(fill=tk.X, pady=(0, 3))
        self.lbl_current_file = ttk.Label(row_cur, font=("Segoe UI", 9, "bold"))
        self.lbl_current_file.pack(side=tk.LEFT)
        self.lbl_current_step = ttk.Label(row_cur, font=("Segoe UI", 9, "bold"), foreground=THEME_BLUE)
        self.lbl_current_step.pack(side=tk.RIGHT)

        self.progress_file = PillProgressBar(self.frame_prog, height=18, show_text=True, maximum=100)
        self.progress_file.pack(fill=tk.X)

        self.frame_queue = ttk.LabelFrame(main_frame, padding=6)
        self.frame_queue.pack(fill=tk.BOTH, expand=False, pady=(0, 6))

        cols = ("file", "model", "status", "task", "started", "finished", "drag")
        self.tree_queue = ttk.Treeview(self.frame_queue, columns=cols, show="headings", height=6, style="Buzz.Treeview", selectmode="extended")
        self.tree_queue.tag_configure("even", background="#F4F5F7")
        self.tree_queue.tag_configure("odd", background="#FFFFFF")
        self.tree_queue.tag_configure("done", background="#E8F5E9", foreground="#1B5E20")
        self.tree_queue.tag_configure("running", background="#E0F2FE", foreground="#0369A1")
        self.tree_queue.tag_configure("failed", background="#FFEBEE", foreground="#B71C1C")
        self.tree_queue.tag_configure("cancelled", background="#F3F4F6", foreground="#9CA3AF")
        self.tree_queue.tag_configure("cancelled", background="#F3F4F6", foreground="#9CA3AF")

        self.tree_queue.column("file", width=245, anchor=tk.W)
        self.tree_queue.column("model", width=105, anchor=tk.W)
        self.tree_queue.column("status", width=115, anchor=tk.W)
        self.tree_queue.column("task", width=290, anchor=tk.W)
        self.tree_queue.column("started", width=128, anchor=tk.W)
        self.tree_queue.column("finished", width=128, anchor=tk.W)
        self.tree_queue.column("drag", width=32, minwidth=32, stretch=False, anchor=tk.CENTER)

        q_scroll = ttk.Scrollbar(self.frame_queue, orient=tk.VERTICAL, command=self.tree_queue.yview)
        self.tree_queue.configure(yscrollcommand=q_scroll.set)
        q_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_queue.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_queue.bind("<Control-c>", self.copy_selected_queue_rows)
        self.tree_queue.bind("<Control-C>", self.copy_selected_queue_rows)
        self.tree_queue.bind("<Delete>", self.toggle_cancel_selected_queue_rows)
        self.tree_queue.bind("<Alt-Up>", lambda e: self.move_selected_queue_rows(-1))
        self.tree_queue.bind("<Alt-Down>", lambda e: self.move_selected_queue_rows(1))
        self.tree_queue.bind("<Alt-Home>", lambda e: self.move_selected_queue_rows("top"))
        self.tree_queue.bind("<Alt-End>", lambda e: self.move_selected_queue_rows("bottom"))
        self.tree_queue.bind("<Motion>", self._on_queue_hover_cursor)
        self.tree_queue.bind("<ButtonPress-1>", self._on_queue_drag_start)
        self.tree_queue.bind("<B1-Motion>", self._on_queue_drag_motion)
        self.tree_queue.bind("<ButtonRelease-1>", self._on_queue_drag_end)
        self.tree_queue.bind("<Button-3>", self.show_queue_context_menu)

        self.frame_log = ttk.LabelFrame(main_frame, padding=6)
        self.frame_log.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(self.frame_log, wrap=tk.WORD, bg="#1E1E1E", fg="#E6E6E6", font=("Consolas", 9), height=6, state=tk.DISABLED)
        self.txt_log.tag_configure("error", foreground="#FF6B6B", font=("Consolas", 9, "bold"))
        self.txt_log.tag_configure("diag_title", foreground="#FFD166", font=("Consolas", 10, "bold"))
        self.txt_log.tag_configure("diag_body", foreground="#FFE8A3", font=("Consolas", 9))
        self.txt_log.tag_configure("success", foreground="#6BCB77", font=("Consolas", 9, "bold"))
        self.txt_log.tag_configure("dl_progress", foreground="#38BDF8", font=("Consolas", 9, "bold"))

        scrollbar = ttk.Scrollbar(self.frame_log, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def open_model_manager(self):
        if self.is_running:
            return
        if getattr(self, "_model_mgr_win", None) and self._model_mgr_win.winfo_exists():
            self._model_mgr_win.lift()
            self._model_mgr_win.focus_force()
            return
        self._model_mgr_win = ModelManagerDialog(self)

    def restart_application(self):
        self.save_config()
        clean_env = os.environ.copy()
        for k in list(clean_env.keys()):
            if k.startswith("_PYI_") or k in ("TCL_LIBRARY", "TK_LIBRARY", "_MEIPASS2"):
                clean_env.pop(k, None)
        try:
            if getattr(sys, "frozen", False):
                if os.name == "nt":
                    exe_str = subprocess.list2cmdline([sys.executable, *sys.argv[1:]])
                    subprocess.Popen(
                        f'cmd.exe /c ping -n 2 127.0.0.1 >nul & start "" {exe_str}',
                        shell=True,
                        env=clean_env,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    subprocess.Popen([sys.executable, *sys.argv[1:]], env=clean_env)
            else:
                subprocess.Popen([sys.executable, str(Path(__file__).resolve()), *sys.argv[1:]], env=clean_env)
        finally:
            self.destroy()

    def get_effective_model(self) -> str:
        lang = (self.var_langue.get().strip().lower() or "en")
        base_model = self.var_modele.get().strip()
        if lang == "en" and base_model in ("tiny", "small", "medium"):
            en_model = f"{base_model}.en"
            if is_whisper_model_installed(en_model):
                return en_model
        if not is_whisper_model_installed(base_model) and base_model in ("tiny", "small", "medium"):
            en_model = f"{base_model}.en"
            if is_whisper_model_installed(en_model):
                return en_model
        return base_model

    def format_model_display(self) -> str:
        eff = self.get_effective_model()
        return f"{eff[0].upper() + eff[1:]}"

    def estimate_step_durations(self, media: Path, live_tr_total_sec: float = None):
        dur = self.media_durations.get(str(media), 0.0)
        if dur <= 0:
            return None, None, None

        eff = self.get_effective_model()
        mult = self.live_speed_multiplier

        if self.has_cuda_gpu:
            tr_factor = 0.035 if ("medium" in eff or "large" in eff) else 0.02
            diar_factor = 0.04
            startup_overhead = 8.0
        else:
            try:
                cores = max(1, int(self.var_cpu_threads.get()))
            except ValueError:
                cores = max(1, (os.cpu_count() or 8) // 2)
            ct2_eff_cores = min(4, cores)
            ct2_scale = 4.0 / ct2_eff_cores
            pyt_scale = 6.0 / cores

            if "tiny" in eff:
                tr_factor = 0.08 * ct2_scale
            elif "medium" in eff:
                tr_factor = 0.62 * ct2_scale
            elif "large" in eff:
                tr_factor = 1.15 * ct2_scale
            else:
                tr_factor = 0.25 * ct2_scale

            if eff.endswith(".en"):
                tr_factor *= 0.88

            diar_factor = 0.50 * pyt_scale
            startup_overhead = 22.0

        if live_tr_total_sec is not None and live_tr_total_sec > 0:
            tr_sec = live_tr_total_sec
        else:
            tr_sec = (startup_overhead + (dur * tr_factor)) * mult

        diar_sec = (dur * diar_factor * mult) if self.var_diarize.get() else None

        want_slides = self.var_out_docx_slides.get() or self.var_out_pptx.get()
        has_vid = self.media_has_video.get(str(media), media.suffix.lower() in {".mp4", ".mkv", ".mov", ".avi"})
        slides_sec = max(8.0, dur * 0.015) if (want_slides and has_vid) else None
        return tr_sec, diar_sec, slides_sec

    def get_task_label(self, media: Path = None, live_tr_total_sec: float = None) -> str:
        want_diar = self.var_diarize.get()
        want_slides = self.var_out_docx_slides.get() or self.var_out_pptx.get()

        if media is not None:
            has_vid = self.media_has_video.get(str(media), media.suffix.lower() in {".mp4", ".mkv", ".mov", ".avi"})
            tr_sec, diar_sec, slides_sec = self.estimate_step_durations(media, live_tr_total_sec)
        else:
            has_vid = True
            tr_sec = diar_sec = slides_sec = None

        parts = []
        if tr_sec is not None:
            parts.append(f"Transcribe (est.: {format_elapsed_buzz(tr_sec)})")
        else:
            parts.append("Transcribe")

        if want_diar:
            if diar_sec is not None:
                parts.append(f"Diarize (est.: {format_elapsed_buzz(diar_sec)})")
            else:
                parts.append("Diarize")

        if want_slides and has_vid:
            if slides_sec is not None:
                parts.append(f"Slides (est.: {format_elapsed_buzz(slides_sec)})")
            else:
                parts.append("Slides")

        return " + ".join(parts)

    def get_total_remaining_estimate(self) -> float:
        total_rem = 0.0
        for media in self.selected_files:
            if str(media) in self.cancelled_files or self.are_all_selected_outputs_done(media):
                continue
            tr_s, diar_s, sld_s = self.estimate_step_durations(media)
            if tr_s is not None:
                file_est = tr_s + (diar_s or 0.0) + (sld_s or 0.0)
                if self.is_running and self.current_media_path and str(media) == str(self.current_media_path):
                    pct = float(self.progress_file["value"] or 0.0)
                    file_est = file_est * max(0.05, 1.0 - (pct / 100.0))
                total_rem += file_est
        return total_rem

    def get_current_file_remaining_estimate(self) -> float:
        if not self.is_running or not self.current_media_path:
            return 0.0
        tr_s, diar_s, sld_s = self.estimate_step_durations(self.current_media_path)
        if tr_s is None:
            return 0.0
        file_est = tr_s + (diar_s or 0.0) + (sld_s or 0.0)
        pct = float(self.progress_file["value"] or 0.0)
        return file_est * max(0.02, 1.0 - (pct / 100.0))

    def update_current_file_eta_cell(self):
        if not self.is_running or not self.current_media_path:
            return
        rem_cur = self.get_current_file_remaining_estimate()
        if rem_cur <= 0:
            return
        from datetime import timedelta
        eta_dt = datetime.now() + timedelta(seconds=rem_cur)
        eta_str = f"est.: {eta_dt.strftime('%Y-%m-%d %H:%M:%S')}"
        self.update_queue_row(self.current_media_path, finished=eta_str, scroll_to=False)

    def copy_selected_queue_rows(self, event=None):
        sel = self.tree_queue.selection()
        if not sel:
            return "break"
        headers = [self.tree_queue.heading(c, "text") for c in ("file", "model", "status", "task", "started", "finished")]
        lines = ["\t".join(headers)]
        for iid in sel:
            vals = self.tree_queue.item(iid, "values")
            if vals:
                lines.append("\t".join(str(v) for v in vals[:6]))
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        return "break"

    def copy_all_queue_rows(self):
        children = self.tree_queue.get_children()
        if not children:
            return
        headers = [self.tree_queue.heading(c, "text") for c in ("file", "model", "status", "task", "started", "finished")]
        lines = ["\t".join(headers)]
        for iid in children:
            vals = self.tree_queue.item(iid, "values")
            if vals:
                lines.append("\t".join(str(v) for v in vals[:6]))
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))

    def _is_media_movable(self, media: Path) -> bool:
        if self.is_running and self.current_media_path and str(media) == str(self.current_media_path):
            return False
        return True

    def _get_min_movable_index(self) -> int:
        return 0

    def _sync_tree_order_with_selected_files(self, keep_selected_paths: set = None):
        for idx, media in enumerate(self.selected_files):
            iid = self.queue_item_ids.get(str(media))
            if iid and self.tree_queue.exists(iid):
                self.tree_queue.move(iid, "", idx)
                tags = self.tree_queue.item(iid, "tags") or ()
                if "even" in tags or "odd" in tags:
                    self.tree_queue.item(iid, tags=("even" if idx % 2 == 0 else "odd",))
        if keep_selected_paths:
            to_sel = [self.queue_item_ids[p] for p in keep_selected_paths if p in self.queue_item_ids and self.tree_queue.exists(self.queue_item_ids[p])]
            if to_sel:
                self.tree_queue.selection_set(to_sel)
                self.tree_queue.see(to_sel[0])
        self.update_queue_header_timer()

    def move_selected_queue_rows(self, direction):
        sel_iids = set(self.tree_queue.selection())
        if not sel_iids:
            return "break"
        iid_to_media = {v: k for k, v in self.queue_item_ids.items()}
        sel_paths = {iid_to_media[i] for i in sel_iids if i in iid_to_media}
        movable_paths = {p for p in sel_paths if self._is_media_movable(Path(p))}
        if not movable_paths:
            return "break"

        min_idx = self._get_min_movable_index()
        files = list(self.selected_files)
        n = len(files)

        if direction == "top":
            head = files[:min_idx]
            tail = files[min_idx:]
            moved = [m for m in tail if str(m) in movable_paths]
            rest = [m for m in tail if str(m) not in movable_paths]
            self.selected_files = head + moved + rest
        elif direction == "bottom":
            head = files[:min_idx]
            tail = files[min_idx:]
            moved = [m for m in tail if str(m) in movable_paths]
            rest = [m for m in tail if str(m) not in movable_paths]
            self.selected_files = head + rest + moved
        elif direction == -1:
            for i in range(min_idx + 1, n):
                if str(files[i]) in movable_paths and str(files[i - 1]) not in movable_paths:
                    files[i], files[i - 1] = files[i - 1], files[i]
            self.selected_files = files
        elif direction == 1:
            for i in range(n - 2, min_idx - 1, -1):
                if str(files[i]) in movable_paths and str(files[i + 1]) not in movable_paths:
                    files[i], files[i + 1] = files[i + 1], files[i]
            self.selected_files = files

        self._sync_tree_order_with_selected_files(keep_selected_paths=movable_paths)
        return "break"

    def _on_queue_hover_cursor(self, event):
        region = self.tree_queue.identify_region(event.x, event.y)
        if region != "cell":
            self.tree_queue.configure(cursor="")
            return
        col = self.tree_queue.identify_column(event.x)
        iid = self.tree_queue.identify_row(event.y)
        if col == "#7" and iid:
            iid_to_media = {v: k for k, v in self.queue_item_ids.items()}
            m_str = iid_to_media.get(iid)
            if m_str and self._is_media_movable(Path(m_str)):
                self.tree_queue.configure(cursor="fleur")
                return
        self.tree_queue.configure(cursor="")

    def _on_queue_drag_start(self, event):
        iid = self.tree_queue.identify_row(event.y)
        if not iid:
            self._drag_start_iid = None
            return
        iid_to_media = {v: k for k, v in self.queue_item_ids.items()}
        m_str = iid_to_media.get(iid)
        if m_str and self._is_media_movable(Path(m_str)):
            self._drag_start_iid = iid
            self.tree_queue.configure(cursor="fleur")
        else:
            self._drag_start_iid = None

    def _on_queue_drag_motion(self, event):
        if not self._drag_start_iid:
            return
        target_iid = self.tree_queue.identify_row(event.y)
        if not target_iid or target_iid == self._drag_start_iid:
            return
        iid_to_media = {v: k for k, v in self.queue_item_ids.items()}
        src_str = iid_to_media.get(self._drag_start_iid)
        dst_str = iid_to_media.get(target_iid)
        if not src_str or not dst_str:
            return
        min_idx = self._get_min_movable_index()
        files = list(self.selected_files)
        try:
            src_idx = next(i for i, m in enumerate(files) if str(m) == src_str)
            dst_idx = next(i for i, m in enumerate(files) if str(m) == dst_str)
        except StopIteration:
            return
        if dst_idx < min_idx:
            dst_idx = min_idx
        if src_idx != dst_idx and src_idx >= min_idx:
            item = files.pop(src_idx)
            files.insert(dst_idx, item)
            self.selected_files = files
            self._sync_tree_order_with_selected_files(keep_selected_paths={src_str})

    def _on_queue_drag_end(self, event):
        self._drag_start_iid = None
        self._on_queue_hover_cursor(event)

    def _is_media_cancelable(self, media: Path) -> bool:
        if self.is_running and self.current_media_path and str(media) == str(self.current_media_path):
            return False
        iid = self.queue_item_ids.get(str(media))
        if iid and self.tree_queue.exists(iid):
            tags = self.tree_queue.item(iid, "tags") or ()
            if "done" in tags or "running" in tags:
                return False
        if self.are_all_selected_outputs_done(media):
            return False
        return True

    def set_selected_queue_rows_cancelled(self, cancel: bool):
        sel_iids = set(self.tree_queue.selection())
        if not sel_iids:
            return
        for idx, media in enumerate(self.selected_files):
            m_str = str(media)
            iid = self.queue_item_ids.get(m_str)
            if not iid or iid not in sel_iids:
                continue
            if not self._is_media_cancelable(media):
                continue
            if cancel:
                self.cancelled_files.add(m_str)
                self.update_queue_row(media, status=self.tr("st_cancelled"), tag="cancelled", scroll_to=False)
            else:
                self.cancelled_files.discard(m_str)
                self.active_processed_paths.discard(m_str)
                row_tag = "even" if idx % 2 == 0 else "odd"
                self.update_queue_row(media, status=self.tr("st_pending"), tag=row_tag, scroll_to=False)
        self.after(20, self.update_queue_header_timer)

    def toggle_cancel_selected_queue_rows(self, event=None):
        sel_iids = set(self.tree_queue.selection())
        if not sel_iids:
            return "break"
        iid_to_media = {v: k for k, v in self.queue_item_ids.items()}
        movable = [iid_to_media[i] for i in sel_iids if i in iid_to_media and self._is_media_cancelable(Path(iid_to_media[i]))]
        if not movable:
            return "break"
        any_not_cancelled = any(p not in self.cancelled_files for p in movable)
        self.set_selected_queue_rows_cancelled(cancel=any_not_cancelled)
        return "break"

    def show_queue_context_menu(self, event):
        iid = self.tree_queue.identify_row(event.y)
        if iid and iid not in self.tree_queue.selection():
            self.tree_queue.selection_set(iid)
        sel_iids = set(self.tree_queue.selection())
        iid_to_media = {v: k for k, v in self.queue_item_ids.items()}
        movable_paths = [iid_to_media[i] for i in sel_iids if i in iid_to_media and self._is_media_cancelable(Path(iid_to_media[i]))]

        menu = tk.Menu(self, tearoff=0)
        is_fr = (self.ui_lang == "fr")

        if movable_paths:
            has_active = any(p not in self.cancelled_files for p in movable_paths)
            has_cancelled = any(p in self.cancelled_files for p in movable_paths)

            if has_active:
                menu.add_command(
                    label="🚫 Annuler la sélection (Suppr)" if is_fr else "🚫 Cancel selected file(s) (Del)",
                    command=lambda: self.set_selected_queue_rows_cancelled(True)
                )
            if has_cancelled:
                menu.add_command(
                    label="♻️ Réactiver la sélection" if is_fr else "♻️ Re-enable selected file(s)",
                    command=lambda: self.set_selected_queue_rows_cancelled(False)
                )
            menu.add_separator()

        lbl_sel = "📋 Copier la sélection (Ctrl+C)" if is_fr else "📋 Copy selected rows (Ctrl+C)"
        lbl_all = "📑 Copier tout le tableau" if is_fr else "📑 Copy entire table"
        menu.add_command(label=lbl_sel, command=self.copy_selected_queue_rows)
        menu.add_command(label=lbl_all, command=self.copy_all_queue_rows)
        menu.tk_popup(event.x_root, event.y_root)

    def update_queue_header_timer(self, scroll_top: bool = False):
        from datetime import timedelta
        elapsed_sec = (time.time() - self.batch_start_time) if (self.is_running and self.batch_start_time) else 0.0
        rem_sec = self.get_total_remaining_estimate()
        probed_count = sum(1 for m in self.selected_files if str(m) in self.media_durations)
        total_count = len(self.selected_files)
        rem_str = format_elapsed_buzz(rem_sec)
        if total_count > 0 and probed_count < total_count:
            rem_str = f"{rem_str} ({probed_count}/{total_count}...)"
        if rem_sec > 0 and total_count > 0:
            fin_dt = datetime.now() + timedelta(seconds=rem_sec)
            eta_end_str = fin_dt.strftime("%d/%m %H:%M") if fin_dt.date() != datetime.now().date() else fin_dt.strftime("%H:%M")
        else:
            eta_end_str = "--:--"
        title_txt = self.tr(
            "sec_queue_fmt",
            elapsed=format_elapsed_buzz(elapsed_sec),
            remaining=rem_str,
            eta_end=eta_end_str
        )
        self.frame_queue.config(text=title_txt)
        if scroll_top and not self.is_running:
            self.tree_queue.yview_moveto(0.0)

    def _tick_queue_timer(self):
        if self.is_running:
            self.update_queue_header_timer()
            if int(time.time()) % 3 == 0:
                self.update_current_file_eta_cell()
            self.after(1000, self._tick_queue_timer)

    def update_diarize_status_display(self):
        diar_installed = is_whisper_model_installed("pyannote")
        if hasattr(self, "chk_diarize"):
            self.chk_diarize.configure(style="TCheckbutton" if diar_installed else "Missing.TCheckbutton")
        if hasattr(self, "lbl_diarize_warn"):
            if self.var_diarize.get() and not diar_installed:
                self.lbl_diarize_warn.config(text=self.tr("diarize_missing_hint"), foreground="#6B7280", font=("Segoe UI", 8, "italic"))
            else:
                self.lbl_diarize_warn.config(text="")

    def refresh_model_display(self):
        base_model = self.var_modele.get().strip()
        eff_model = self.get_effective_model()
        installed = is_whisper_model_installed(eff_model)
        if not installed:
            self.lbl_model_effective.config(
                text=self.tr("model_missing_hint", eff=eff_model),
                foreground="#6B7280",
                font=("Segoe UI", 8, "italic")
            )
        elif eff_model != base_model:
            self.lbl_model_effective.config(
                text=self.tr("model_eff_hint", eff=eff_model),
                foreground=THEME_BLUE,
                font=("Segoe UI", 8, "bold")
            )
        else:
            self.lbl_model_effective.config(text="")
        self.update_diarize_status_display()

        if eff_model == self.last_effective_model:
            return
        self.last_effective_model = eff_model

        if self.selected_files and not self.is_running:
            self.populate_queue_table(self.selected_files)
            return

        model_label = self.format_model_display()
        for media in self.selected_files:
            if self.current_media_path and str(media) == str(self.current_media_path):
                continue
            iid = self.queue_item_ids.get(str(media))
            if iid and self.tree_queue.exists(iid):
                tags = self.tree_queue.item(iid, "tags") or ()
                if "done" in tags or "failed" in tags:
                    continue
                vals = list(self.tree_queue.item(iid, "values"))
                if vals:
                    vals[1] = model_label
                    vals[3] = self.get_task_label(media)
                    self.tree_queue.item(iid, values=vals)
        self.update_queue_header_timer()
        if self.is_running:
            self.log(f"\n[MODÈLE] Modèle Whisper modifié en direct à '{eff_model}' (sera appliqué dès le fichier suivant).\n", "dl_progress")

    def refresh_all_queue_tasks(self):
        for media in self.selected_files:
            iid = self.queue_item_ids.get(str(media))
            if iid and self.tree_queue.exists(iid):
                tags = self.tree_queue.item(iid, "tags") or ()
                if "done" in tags or "failed" in tags:
                    continue
                vals = list(self.tree_queue.item(iid, "values"))
                if vals:
                    vals[3] = self.get_task_label(media)
                    self.tree_queue.item(iid, values=vals)
        self.update_current_file_eta_cell()
        self.update_queue_header_timer()

    def on_output_format_changed(self):
        self.save_config()
        if self.selected_files and not self.is_running:
            self.populate_queue_table(self.selected_files)

    def on_ui_lang_changed(self, event=None):
        chosen = self.var_ui_lang.get().strip().lower()
        if chosen in I18N:
            self.ui_lang = chosen
            self.save_config()
            self.apply_ui_language()

    def on_language_changed(self, event=None):
        self.save_config()
        self.refresh_model_display()

    def on_model_changed(self, event=None):
        self.save_config()
        self.refresh_model_display()

    def on_diarize_changed(self):
        self.save_config()
        self.update_diarize_status_display()
        if self.selected_files and not self.is_running:
            self.populate_queue_table(self.selected_files)
            return
        self.refresh_all_queue_tasks()

    def on_cpu_cores_changed(self, event=None):
        self.save_config()
        self.update_cpu_warning()
        self.refresh_all_queue_tasks()
        if self.is_running and self.current_process:
            try:
                chosen = int(self.var_cpu_threads.get())
                set_process_cpu_cores_affinity(self.current_process.pid, chosen)
                self.log(f"\n[CPU] Nombre de cœurs ajusté en direct à {chosen} cœur(s) (immédiat sur le processus actif + appliqué aux fichiers suivants).\n", "dl_progress")
            except Exception:
                pass

    def update_cpu_warning(self):
        total_cores = os.cpu_count() or 8
        half_cores = max(1, total_cores // 2)
        try:
            chosen = int(self.var_cpu_threads.get())
        except ValueError:
            chosen = half_cores

        if chosen <= half_cores:
            self.lbl_cpu_warn.config(text="")
        elif chosen >= max(half_cores + 2, int(total_cores * 0.85)):
            self.lbl_cpu_warn.config(text=self.tr("warn_cpu_high"), foreground="#B91C1C")
        else:
            self.lbl_cpu_warn.config(text=self.tr("warn_cpu_med"), foreground="#B45309")

    def apply_ui_language(self):
        self.title(self.tr("win_title"))

        self.btn_models_mgr.config(text=self.tr("btn_models_mgr"))
        self.lbl_ui_lang.config(text=self.tr("ui_lang_label"))
        self.lbl_hw_status.config(text=self.tr("hw_cuda") if self.has_cuda_gpu else self.tr("hw_cpu"))
        self.frame_src.config(text=self.tr("sec1"))
        self.btn_1file.config(text=" " + self.tr("btn_1file"))
        self.btn_mfiles.config(text=" " + self.tr("btn_mfiles"))
        self.btn_folder.config(text=" " + self.tr("btn_folder"))
        self.chk_sub.config(text=self.tr("chk_sub"))
        if hasattr(self, "lbl_app_version"): self.lbl_app_version.config(text=self.tr("app_ver"))

        if not self.selected_files:
            self.lbl_selection.config(text=self.tr("no_sel"))

        self.frame_opt.config(text=self.tr("sec2"))
        self.lbl_fmt.config(text=self.tr("lbl_fmt"))
        self.chk_out_txt.config(text=self.tr("fmt_txt"))
        self.chk_out_docx.config(text=self.tr("fmt_docx"))
        self.chk_out_docx_slides.config(text=self.tr("fmt_docx_slides"))
        self.chk_out_pptx.config(text=self.tr("fmt_pptx"))
        self.lbl_lang.config(text=self.tr("lbl_lang"))
        self.lbl_model.config(text=self.tr("lbl_model"))
        self.refresh_model_display()
        self.chk_diarize.config(text=self.tr("chk_diarize"))
        self.update_diarize_status_display()
        self.lbl_cpu_cores.config(text=self.tr("lbl_cpu_cores"))
        self.lbl_cpu_gpu_hint.config(text=self.tr("cpu_hidden_gpu"))
        self.update_cpu_warning()

        n = len(self.selected_files)
        if n > 0 and not self.is_running:
            self.btn_start.config(text=" " + self.tr("btn_start_n", n=n, s="s" if n > 1 else ""))
        else:
            self.btn_start.config(text=" " + self.tr("btn_start"))

        self.btn_pause.config(text=" " + (self.tr("btn_resume") if self.is_paused else self.tr("btn_pause")))
        self.btn_stop.config(text=" " + self.tr("btn_stop"))
        self.btn_diag.config(text=" " + self.tr("btn_diag"))
        self.btn_clear.config(text=" " + self.tr("btn_clear"))

        self.frame_prog.config(text=self.tr("sec3"))
        self.lbl_tot_prog_title.config(text=self.tr("lbl_tot_prog"))
        if not self.is_running:
            self.lbl_total_progress.config(text=self.tr("tot_fmt", done=0, total=n, pct=0))
            self.lbl_current_file.config(text=self.tr("cur_none"))
            self.lbl_current_step.config(text=self.tr("step_wait"))
        else:
            done_cnt = int(self.progress_total["value"] or 0)
            pct_tot = int((done_cnt / n) * 100) if n > 0 else 0
            self.lbl_total_progress.config(text=self.tr("tot_fmt", done=done_cnt, total=n, pct=pct_tot))
            if self.current_media_path:
                self.lbl_current_file.config(text=self.tr("cur_fmt", name=self.current_media_path.name))
            self.lbl_current_step.config(text=self.tr(self.current_step_key, **self.current_step_kwargs))
        self.update_queue_header_timer()
        self.tree_queue.heading("file", text=self.tr("col_file"), anchor=tk.W)
        self.tree_queue.heading("model", text=self.tr("col_model"), anchor=tk.W)
        self.tree_queue.heading("status", text=self.tr("col_status"), anchor=tk.W)
        self.tree_queue.heading("task", text=self.tr("col_task"), anchor=tk.W)
        self.tree_queue.heading("started", text=self.tr("col_started"), anchor=tk.W)
        self.tree_queue.heading("finished", text=self.tr("col_finished"), anchor=tk.W)
        self.tree_queue.heading("drag", text="", anchor=tk.CENTER)

        self.frame_log.config(text=self.tr("sec_log"))

    def populate_queue_table(self, files):
        for item in self.tree_queue.get_children():
            self.tree_queue.delete(item)
        self.queue_item_ids.clear()
        self.cancelled_files.clear()
        self.cancelled_files.clear()

        model_label = self.format_model_display()

        for i, media in enumerate(files):
            task_label = self.get_task_label(media)

            if str(media) in self.cancelled_files or self.are_all_selected_outputs_done(media):
                status = self.tr("st_skipped")
                tag = "done"
                expected = self.get_expected_outputs(media)
                existing = [p for p in expected if p.exists()]
                try:
                    ftime = datetime.fromtimestamp(existing[0].stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S") if existing else "-"
                except Exception:
                    ftime = "-"
            else:
                status = self.tr("st_pending")
                tag = "even" if i % 2 == 0 else "odd"
                ftime = ""

            iid = self.tree_queue.insert(
                "", tk.END,
                values=(media.name, model_label, status, task_label, "", ftime, "≡"),
                tags=(tag,)
            )
            self.queue_item_ids[str(media)] = iid

        self.tree_queue.yview_moveto(0.0)
        self.update_queue_header_timer(scroll_top=True)
        threading.Thread(target=self._probe_files_metadata_worker, args=(list(files),), daemon=True).start()

    def _probe_files_metadata_worker(self, files):
        total = len(files)
        for idx, media in enumerate(files, start=1):
            key = str(media)
            if key not in self.media_durations:
                self.media_durations[key] = get_media_duration(media)
            if key not in self.media_has_video:
                self.media_has_video[key] = has_video_stream(media)
            new_task = self.get_task_label(media)
            self.update_queue_row(media, status=None, task=new_task, scroll_to=False)
            is_last = (idx == total)
            self.after(0, lambda last=is_last: self.update_queue_header_timer(scroll_top=last))
        if total == 0:
            self.after(0, lambda: self.update_queue_header_timer(scroll_top=True))

    def update_queue_row(self, media: Path, status: str = None, finished: str = None, tag: str = None, started: str = None, task: str = None, scroll_to: bool = True):
        def _upd():
            iid = self.queue_item_ids.get(str(media))
            if not iid or not self.tree_queue.exists(iid):
                return
            vals = list(self.tree_queue.item(iid, "values"))
            if status is not None:
                vals[2] = status
            if task is not None:
                vals[3] = task
            if started is not None:
                vals[4] = started
            if finished is not None:
                vals[5] = finished
            if len(vals) > 6:
                if tag == "running":
                    vals[6] = ""
                elif tag is not None:
                    vals[6] = "≡"
            self.tree_queue.item(iid, values=vals)
            if tag:
                self.tree_queue.item(iid, tags=(tag,))
            if scroll_to and self.is_running:
                self.tree_queue.see(iid)
                try:
                    self.tree_queue.selection_set(iid)
                    self.tree_queue.focus(iid)
                except Exception:
                    pass
        self.after(0, _upd)

    def update_total_bar(self, completed: int, total: int):
        def _upd():
            pct = int((completed / total) * 100) if total > 0 else 0
            self.progress_total["maximum"] = max(total, 1)
            self.progress_total["value"] = completed
            self.lbl_total_progress.config(text=self.tr("tot_fmt", done=completed, total=total, pct=pct))
            if not self.is_running:
                if total > 0:
                    self.btn_start.config(state=tk.NORMAL, text=" " + self.tr("btn_start_n", n=total, s="s" if total > 1 else ""))
                else:
                    self.btn_start.config(state=tk.NORMAL, text=" " + self.tr("btn_start"))
        self.after(0, _upd)

    def update_file_bar(self, filename: str, step_text: str, pct: float):
        def _upd():
            clamped = max(0.0, min(100.0, pct))
            self.lbl_current_file.config(text=self.tr("cur_fmt", name=filename))
            self.lbl_current_step.config(text=step_text)
            self.progress_file["value"] = clamped
            if self.is_running and self.current_media_path and 0 < clamped < 100:
                iid = self.queue_item_ids.get(str(self.current_media_path))
                if iid and self.tree_queue.exists(iid):
                    vals = list(self.tree_queue.item(iid, "values"))
                    vals[2] = self.tr("st_running", pct=int(clamped))
                    self.tree_queue.item(iid, values=vals)
        self.after(0, _upd)

    def get_initial_dir(self) -> str:
        saved_dir = self.config_data.get("last_dir", "").strip()
        if saved_dir:
            p = Path(saved_dir)
            if p.exists() and p.is_dir():
                return str(p)
        return str(SCRIPT_DIR)

    def append_files_to_queue(self, new_paths: list):
        existing_set = {str(x.resolve()) if x.exists() else str(x) for x in self.selected_files}
        added = []
        for raw in new_paths:
            p = Path(raw)
            if p.is_dir():
                it = p.rglob("*") if self.var_sous_dossiers.get() else p.iterdir()
                for sub in sorted(it):
                    if sub.is_file() and sub.suffix.lower() in EXTENSIONS_MEDIA:
                        key = str(sub.resolve())
                        if key not in existing_set:
                            existing_set.add(key)
                            added.append(sub)
            elif p.is_file() and p.suffix.lower() in EXTENSIONS_MEDIA:
                key = str(p.resolve())
                if key not in existing_set:
                    existing_set.add(key)
                    added.append(p)
        if not added:
            return
        self.selected_files.extend(added)
        self.config_data["last_dir"] = str(added[0].parent)
        self.save_config()
        n_tot = len(self.selected_files)
        if n_tot == 1:
            self.lbl_selection.config(text=self.tr("sel_1", name=self.selected_files[0].name))
        else:
            self.lbl_selection.config(text=self.tr("sel_n", n=n_tot, parent=self.selected_files[0].parent))

        if self.is_running:
            model_label = self.format_model_display()
            start_idx = len(self.tree_queue.get_children())
            for offset, media in enumerate(added):
                task_label = self.get_task_label(media)
                if self.are_all_selected_outputs_done(media):
                    status = self.tr("st_skipped")
                    tag = "done"
                else:
                    status = self.tr("st_pending")
                    tag = "even" if (start_idx + offset) % 2 == 0 else "odd"
                row_vals = (media.name, model_label, status, task_label, "", "", "≡")
                iid = self.tree_queue.insert("", tk.END, values=row_vals, tags=(tag,))
                self.queue_item_ids[str(media)] = iid
            done_cnt = int(self.progress_total["value"] or 0)
            self.update_total_bar(done_cnt, n_tot)
            self.update_queue_header_timer()
            self.log(f"[FILE] +{len(added)} fichier(s) ajouté(s) en direct à la file d'attente (total : {n_tot}).\n", "success")
            threading.Thread(target=self._probe_files_metadata_worker, args=(list(added),), daemon=True).start()
        else:
            self.populate_queue_table(self.selected_files)
            self.update_total_bar(0, n_tot)

    def on_files_dropped(self, dropped_paths: list):
        self.append_files_to_queue(dropped_paths)

    def select_single_file(self):
        filetypes = [("Audio/Video", "*.mp4 *.mkv *.mov *.avi *.mp3 *.wav *.m4a *.flac *.ogg"), ("All files", "*.*")]
        path = filedialog.askopenfilename(initialdir=self.get_initial_dir(), title=self.tr("btn_1file"), filetypes=filetypes)
        if path:
            if self.is_running:
                self.append_files_to_queue([path])
            else:
                p = Path(path)
                self.config_data["last_dir"] = str(p.parent)
                self.save_config()
                self.selected_files = [p]
                self.lbl_selection.config(text=self.tr("sel_1", name=p.name))
                self.populate_queue_table(self.selected_files)
                self.update_total_bar(0, 1)

    def select_multiple_files(self):
        filetypes = [("Audio/Video", "*.mp4 *.mkv *.mov *.avi *.mp3 *.wav *.m4a *.flac *.ogg"), ("All files", "*.*")]
        paths = filedialog.askopenfilenames(initialdir=self.get_initial_dir(), title=self.tr("btn_mfiles"), filetypes=filetypes)
        if paths:
            if self.is_running:
                self.append_files_to_queue(list(paths))
            else:
                self.selected_files = [Path(p) for p in paths]
                self.config_data["last_dir"] = str(self.selected_files[0].parent)
                self.save_config()
                self.lbl_selection.config(text=self.tr("sel_n", n=len(self.selected_files), parent=self.selected_files[0].parent))
                self.populate_queue_table(self.selected_files)
                self.update_total_bar(0, len(self.selected_files))

    def select_folder(self):
        folder = filedialog.askdirectory(initialdir=self.get_initial_dir(), title=self.tr("btn_folder"))
        if folder:
            if self.is_running:
                self.append_files_to_queue([folder])
            else:
                d = Path(folder)
                self.config_data["last_dir"] = str(d)
                self.save_config()
                iterateur = d.rglob("*") if self.var_sous_dossiers.get() else d.iterdir()
                self.selected_files = sorted([f for f in iterateur if f.is_file() and f.suffix.lower() in EXTENSIONS_MEDIA])
                self.lbl_selection.config(text=self.tr("sel_dir", folder=d, n=len(self.selected_files)))
                self.populate_queue_table(self.selected_files)
                self.update_total_bar(0, len(self.selected_files))

    def log(self, text: str, tag: str = None, replace_last_line: bool = False):
        def append():
            self.txt_log.configure(state=tk.NORMAL)
            if replace_last_line and self.inline_log_active:
                self.txt_log.delete("end-2l", "end-1l")
            if tag:
                self.txt_log.insert(tk.END, text, tag)
            else:
                self.txt_log.insert(tk.END, text)
            self.inline_log_active = replace_last_line
            self.txt_log.see(tk.END)
            self.txt_log.configure(state=tk.DISABLED)
        self.after(0, append)

    def clear_log(self):
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.inline_log_active = False
        self.txt_log.configure(state=tk.DISABLED)

    def print_diagnostic_card(self, title: str, step: str, cause: str, action: str):
        self.log("\n┌──────────────────────────────────────────────────────────────────────────┐\n", "diag_title")
        self.log(f"│ 💡 CORRECTION SUGGÉRÉE : {title}\n", "diag_title")
        self.log("├──────────────────────────────────────────────────────────────────────────┤\n", "diag_title")
        self.log(f"│ 📍 Étape concernée : {step}\n", "diag_body")
        self.log("│ 🔎 Cause bloquante :\n", "diag_body")
        for line in cause.splitlines():
            self.log(f"│    {line}\n", "diag_body")
        self.log("│ 🛠 Action à effectuer :\n", "diag_body")
        for line in action.splitlines():
            self.log(f"│    {line}\n", "diag_body")
        self.log("└──────────────────────────────────────────────────────────────────────────┘\n\n", "diag_title")

    def run_preflight_check(self, verbose: bool = True) -> bool:
        blocking_issues = []
        expected_venv = SCRIPT_DIR / "whisperx-env"
        whisperx_exe = expected_venv / "Scripts" / "whisperx.exe"
        runtime_py = RUNTIME_DIR / "python.exe"

        if not runtime_py.exists() and (sys.version_info.major, sys.version_info.minor) != (3, 12):
            blocking_issues.append((
                f"Version Python incompatible ({sys.version_info.major}.{sys.version_info.minor})",
                "Étape 1 & Étape 2 — Python 3.12",
                f"Le script tourne sous {sys.executable}.",
                "Installez Python 3.12 et l'environnement whisperx-env."
            ))
        elif not runtime_py.exists() and not whisperx_exe.exists() and not shutil.which("whisperx"):
            blocking_issues.append((
                "Environnement 'whisperx-env' ou WhisperX introuvable",
                "Étape 2 — Création du projet",
                f"Environnement introuvable : {expected_venv}",
                "Créez whisperx-env et installez whisperx python-docx."
            ))

        if (self.var_out_docx.get() or self.var_out_docx_slides.get()) and not DOCX_AVAILABLE:
            blocking_issues.append((
                "Bibliothèque 'python-docx' absente",
                "Étape 2 — Dépendances Python",
                "Le module 'docx' n'est pas installé.",
                "Installez python-docx dans l'environnement actif."
            ))

        if self.var_out_pptx.get() and not PPTX_AVAILABLE:
            blocking_issues.append((
                "Bibliothèque 'python-pptx' absente",
                "Étape 2 — Dépendances Python",
                "Le module 'pptx' est requis pour générer des présentations PowerPoint.",
                "Installez python-pptx et recompilez avec --collect-all pptx --collect-all lxml."
            ))

        if not shutil.which("ffmpeg"):
            blocking_issues.append((
                "FFmpeg absent du PATH Windows",
                "Étape 1 — Prérequis système",
                "WhisperX ne peut pas lire les médias sans FFmpeg.",
                "Installez FFmpeg : winget install Gyan.FFmpeg"
            ))

        hf_token = get_hf_token(self.config_data)
        hf_access = None
        hf_cache = get_hf_cache_path()
        both_cached = is_whisper_model_installed("pyannote")
        fr_align_cached = is_whisper_model_installed("wav2vec2-fr")
        selected_lang = self.var_langue.get().strip().lower()
        eff_whisper = self.get_effective_model()
        whisper_cached = is_whisper_model_installed(eff_whisper)
        if (not whisper_cached) or (self.var_diarize.get() and not both_cached) or (selected_lang == "fr" and not fr_align_cached):
            missing_models = []
            if not whisper_cached:
                missing_models.append(f"Systran/faster-whisper-{eff_whisper} (modèle Whisper)")
            if self.var_diarize.get() and not both_cached:
                missing_models.append("pyannote/speaker-diarization-community-1 (requis pour l'identification des locuteurs)")
            if selected_lang == "fr" and not fr_align_cached:
                missing_models.append("bofenghuang/asr-wav2vec2-ctc-french (recommandé pour l'alignement audio français)")
            blocking_issues.append((
                "Installation ou mise à jour des modèles requise",
                "Cache local hors-ligne (.\\hf_cache\\)",
                "Modèle(s) manquant(s) dans le cache local :\n• " + "\n• ".join(missing_models),
                "Une installation ou mise à jour est requise pour télécharger les modèles attendus.\nCliquez sur le bouton '⚙ Modèles installés & Mises à jour' en haut à droite."
            ))
        if not blocking_issues:
            if verbose:
                self.has_cuda_gpu = detect_nvidia_cuda()
                installed_list = [m["id"] for m in WHISPER_MODEL_CATALOG if is_whisper_model_installed(m["id"])]
                self.log("[OK] Aucun problème bloquant détecté :\n", "success")
                self.log(f"  ✔ Matériel de calcul : {'GPU NVIDIA CUDA (--device cuda, float16)' if self.has_cuda_gpu else 'CPU uniquement (--device cpu, int8, batch_size=16)'}\n")
                self.log(f"  ✔ Python actif : {sys.executable} (v{sys.version_info.major}.{sys.version_info.minor})\n")
                self.log(f"  ✔ python-docx : {'OK' if DOCX_AVAILABLE else 'Absent'} | python-pptx : {'OK' if PPTX_AVAILABLE else 'Absent'} | Pillow : {'OK' if PIL_AVAILABLE else 'Absent'}\n")
                self.log(f"  ✔ FFmpeg : {shutil.which('ffmpeg')}\n")
                if hf_token:
                    self.log(f"  ✔ HF_TOKEN : Configuré ({hf_token[:5]}...)\n")
                if hf_access:
                    for repo, st in hf_access["results"].items():
                        self.log(f"  ✔ Dépôt HF ({repo}) : {st}\n")
                self.log(f"  ✔ Modèles Whisper en cache : {', '.join(installed_list) if installed_list else 'Aucun'}\n")
                self.log(f"  ✔ Cache local pyannote : {'Présent' if both_cached else 'Téléchargement au 1er lancement'}\n")
                self.log(f"  ✔ Cache local alignement Français (fr) : {'Présent' if fr_align_cached else 'Absent'}\n\n")
            return True

        self.log(f"[BLOQUANT] {len(blocking_issues)} prérequis manquant(s) :\n", "error")
        for title, step, cause, action in blocking_issues:
            self.print_diagnostic_card(title, step, cause, action)
        return False

    def start_processing(self):
        if not self.selected_files:
            messagebox.showwarning(self.tr("dlg_no_sel_t"), self.tr("dlg_no_sel_m"))
            return

        if not (self.var_out_txt.get() or self.var_out_docx.get() or self.var_out_docx_slides.get() or self.var_out_pptx.get()):
            messagebox.showwarning(self.tr("dlg_no_fmt_t"), self.tr("dlg_no_fmt_m"))
            return

        if not self.run_preflight_check(verbose=False):
            messagebox.showerror(self.tr("dlg_pre_t"), self.tr("dlg_pre_m"))
            return

        hf_token = get_hf_token(self.config_data)
        self.save_config()
        saved_cancelled = set(self.cancelled_files)
        self.populate_queue_table(self.selected_files)
        for m in self.selected_files:
            if str(m) in saved_cancelled and not self.are_all_selected_outputs_done(m):
                self.cancelled_files.add(str(m))
                self.update_queue_row(m, status=self.tr("st_cancelled"), tag="cancelled", scroll_to=False)
        self.is_running = True
        self.is_paused = False
        self.stop_requested = False
        self.batch_start_time = time.time()
        self.last_effective_model = self.get_effective_model()
        self.update_queue_header_timer(scroll_top=True)
        self.after(1000, self._tick_queue_timer)

        self.btn_start.config(state=tk.DISABLED, bg="#9CA3AF", cursor="arrow")
        self.btn_models_mgr.config(
            state=tk.DISABLED,
            bg="#6B7280",
            fg="#D1D5DB",
            disabledforeground="#D1D5DB",
            cursor="arrow"
        )
        if getattr(self, "_model_mgr_win", None) and self._model_mgr_win.winfo_exists():
            for _m in list(self._model_mgr_win.row_widgets.keys()):
                self._model_mgr_win._refresh_single_row_state(_m)
        self.btn_pause.config(state=tk.NORMAL, text=" " + self.tr("btn_pause"))
        self.btn_stop.config(state=tk.NORMAL)
        self.update_total_bar(0, len(self.selected_files))

        thread = threading.Thread(target=self.worker_loop, args=(list(self.selected_files), hf_token), daemon=True)
        thread.start()

    def toggle_pause(self):
        if not self.is_running or not self.current_process:
            return
        self.is_paused = not self.is_paused
        suspend_or_resume_pid(self.current_process.pid, self.is_paused)
        if self.is_paused:
            self.btn_pause.config(text=" " + self.tr("btn_resume"))
            self.current_step_key = "step_paused"; self.current_step_kwargs = {}; self.current_step_key = "step_paused"; self.current_step_kwargs = {}; self.current_step_key = "step_paused"; self.current_step_kwargs = {}; self.current_step_key = "step_paused"; self.current_step_kwargs = {}; self.lbl_current_step.config(text=self.tr("step_paused"))
            self.log("\n[PAUSE] Traitement suspendu (CPU libéré).\n", "diag_title")
        else:
            self.btn_pause.config(text=" " + self.tr("btn_pause"))
            self.current_step_key = "step_resumed"; self.current_step_kwargs = {}; self.current_step_key = "step_resumed"; self.current_step_kwargs = {}; self.current_step_key = "step_resumed"; self.current_step_kwargs = {}; self.current_step_key = "step_resumed"; self.current_step_kwargs = {}; self.lbl_current_step.config(text=self.tr("step_resumed"))
            self.log("[REPRISE] Traitement relancé.\n\n", "success")

    def kill_current_process_tree(self):
        if not self.current_process:
            return
        pid = self.current_process.pid
        if self.is_paused:
            suspend_or_resume_pid(pid, False)
            self.is_paused = False
        try:
            if os.name == "nt":
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                self.current_process.terminate()
        except Exception:
            pass

    def stop_processing(self):
        if not self.is_running:
            return
        self.stop_requested = True
        self.kill_current_process_tree()
        self.log("\n[ARRÊT] Interruption demandée.\n", "error")

    def on_closing(self):
        if self.is_running:
            if not messagebox.askyesno(self.tr("dlg_quit_t"), self.tr("dlg_quit_m")):
                return
            self.stop_requested = True
            self.kill_current_process_tree()
        self.destroy()

    def iter_process_lines(self, proc):
        buf = bytearray()
        while True:
            ch = proc.stdout.read(1)
            if not ch:
                if buf:
                    yield buf.decode("utf-8", errors="replace"), False
                break
            if ch in (b"\r", b"\n"):
                if buf:
                    yield buf.decode("utf-8", errors="replace"), (ch == b"\r")
                    buf.clear()
            else:
                buf.extend(ch)

    def worker_loop(self, files, hf_token):
        want_txt = self.var_out_txt.get()
        want_docx = self.var_out_docx.get()
        want_docx_slides = self.var_out_docx_slides.get()
        want_pptx = self.var_out_pptx.get()
        need_slides = want_docx_slides or want_pptx

        langue = self.var_langue.get().strip() or "en"
        modele = self.var_modele.get()
        use_diarize = self.var_diarize.get()
        try:
            selected_threads = int(self.var_cpu_threads.get())
        except ValueError:
            selected_threads = max(1, (os.cpu_count() or 8) // 2)

        runtime_py = RUNTIME_DIR / "python.exe"
        whisperx_exe = SCRIPT_DIR / "whisperx-env" / "Scripts" / "whisperx.exe"
        fast_runner = SCRIPT_DIR / "source" / "run_engine_fast.py"

        if runtime_py.exists() and fast_runner.exists():
            base_cmd = [str(runtime_py), str(fast_runner)]
        elif runtime_py.exists():
            base_cmd = [str(runtime_py), "-m", "whisperx"]
        elif whisperx_exe.exists():
            base_cmd = [str(whisperx_exe)]
        else:
            base_cmd = ["whisperx"]

        env = os.environ.copy()
        if hf_token:
            env["HF_TOKEN"] = hf_token
        env["HF_HUB_OFFLINE"] = "1"
        env["PYTHONUNBUFFERED"] = "1"

        use_cuda = detect_nvidia_cuda()
        cpu_count = os.cpu_count() or 8

        if not use_cuda:
            env["OMP_NUM_THREADS"] = str(selected_threads)
            env["MKL_NUM_THREADS"] = str(selected_threads)
            env["TORCH_NUM_THREADS"] = str(selected_threads)

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        total_files = len(files)
        completed_files = 0

        for idx, media in enumerate(files, start=1):
            if self.stop_requested:
                break

            dossier_sortie = media.parent
            json_temp = dossier_sortie / f"{media.stem}.json"

            if str(media) in self.cancelled_files or self.are_all_selected_outputs_done(media):
                self.log(f"[IGNORÉ] Déjà traité : {media.name}\n")
                self.update_file_bar(media.name, self.tr("step_skip"), 100.0)
                self.update_queue_row(media, self.tr("st_skipped"), tag="done")
                completed_files += 1
                self.update_total_bar(completed_files, total_files)
                continue

            self.current_media_path = media
            t_start = time.time()
            started_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            total_duration = get_media_duration(media)
            self.media_durations[str(media)] = total_duration
            self.media_has_video[str(media)] = has_video_stream(media)
            dur_str = format_short_duration(total_duration)

            self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_vad"), 5.0)
            self.update_queue_row(
                media,
                status=self.tr("st_running", pct=0),
                task=self.get_task_label(media),
                started=started_ts,
                tag="running",
                scroll_to=True
            )
            self.after(250, lambda m=media: self.update_queue_row(m, scroll_to=True))
            self.after(0, lambda m=media, ml=self.format_model_display(): (
                self.tree_queue.set(self.queue_item_ids[str(m)], "model", ml)
                if str(m) in self.queue_item_ids and self.tree_queue.exists(self.queue_item_ids[str(m)]) else None
            ))
            self.log(f"\n========================================\n[EN COURS ({idx}/{total_files})] {media.name}\n========================================\n")

            effective_model = self.get_effective_model()
            self.update_queue_row(media, status=None, task=self.get_task_label(media), scroll_to=False)
            self.after(0, self.update_current_file_eta_cell)
            if effective_model != modele:
                self.log(f"[INFO] Langue anglaise : utilisation automatique du modèle optimisé '{effective_model}'.\n")

            cmd = base_cmd + [str(media), "--model", effective_model]
            if langue.lower() != "auto":
                cmd += ["--language", langue]

            try:
                current_threads = int(self.var_cpu_threads.get())
            except ValueError:
                current_threads = selected_threads

            if use_cuda:
                self.log("[INFO] Accélération GPU NVIDIA CUDA activée (--device cuda, float16, batch_size=16).\n")
                cmd += [
                    "--device", "cuda",
                    "--compute_type", "float16",
                    "--batch_size", "16",
                ]
            else:
                file_env = env.copy()
                file_env["OMP_NUM_THREADS"] = str(current_threads)
                file_env["MKL_NUM_THREADS"] = str(current_threads)
                file_env["TORCH_NUM_THREADS"] = str(current_threads)
                self.log(f"[INFO] Mode CPU ({current_threads}/{cpu_count} cœurs actifs, batch_size=16).\n")
                cmd += [
                    "--device", "cpu",
                    "--compute_type", "int8",
                    "--batch_size", "16",
                    "--threads", str(current_threads),
                ]

            if use_diarize:
                cmd += ["--diarize"]
                if hf_token:
                    cmd += ["--hf_token", hf_token]

            cmd += [
                "--output_format", "json",
                "--output_dir", str(dossier_sortie)
            ]

            self.current_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=0,
                env=(env if use_cuda else file_env),
                creationflags=creationflags
            )
            if not use_cuda:
                set_process_cpu_cores_affinity(self.current_process.pid, current_threads)

            current_stage = "vad"
            captured_lines = []

            for line, is_cr in self.iter_process_lines(self.current_process):
                clean = line.strip()
                if not clean:
                    continue

                prog_match = WHISPERX_PROG_RE.match(clean)
                if prog_match:
                    pct_val = float(prog_match.group(1))
                    elapsed_tr = time.time() - t_start
                    est_total_tr = (elapsed_tr / (pct_val / 100.0)) if pct_val >= 3.0 else None
                    tr_s, diar_s, sld_s = self.estimate_step_durations(media, live_tr_total_sec=est_total_tr)
                    if tr_s and tr_s > 0:
                        tot_file_est = tr_s + (diar_s or 0.0) + (sld_s or 0.0)
                        tr_weight = (tr_s / tot_file_est) * 92.0
                    else:
                        tr_weight = 65.0 if use_diarize else 82.0
                    file_pct = 3.0 + ((pct_val / 100.0) * tr_weight)
                    formatted_prog = f"[{media.name}] 🎙 Transcription : {pct_val:.1f} %\n"
                    self.log(formatted_prog, tag="dl_progress", replace_last_line=True)
                    self.update_file_bar(media.name, f"Transcription ({pct_val:.0f} %)", file_pct)
                    if est_total_tr is not None:
                        self.update_queue_row(media, status=None, task=self.get_task_label(media, live_tr_total_sec=est_total_tr), scroll_to=False)
                    continue

                dl_match = TQDM_DL_RE.match(clean)
                if dl_match:
                    pct_dl, cur_sz, tot_sz, spd = dl_match.groups()
                    formatted_dl = self.tr("dl_line", name=media.name, pct=pct_dl, cur=cur_sz, tot=tot_sz, spd=spd) + "\n"
                    self.log(formatted_dl, tag="dl_progress", replace_last_line=True)
                    self.update_file_bar(media.name, f"⬇ {pct_dl} % — {cur_sz} / {tot_sz} ({spd})", float(pct_dl))
                    continue

                captured_lines.append(clean + "\n")
                self.log(clean + "\n", replace_last_line=False)

                lower_line = clean.lower()

                ts_match = TRANSCRIPT_TS_RE.search(clean)
                if ts_match:
                    current_stage = "transcribe"
                    end_sec = float(ts_match.group(2))
                    if total_duration > 0:
                        ratio = min(1.0, end_sec / total_duration)
                        max_tr_pct = 65.0 if use_diarize else 82.0
                        file_pct = 5.0 + (ratio * max_tr_pct)
                        cur_str = format_short_duration(end_sec)
                        self.update_file_bar(
                            media.name,
                            self.tr("step_tr_pct", pct=int(ratio * 100), cur=cur_str, tot=dur_str),
                            file_pct
                        )
                        self.update_queue_row(media, self.tr("st_running", pct=int(file_pct)), scroll_to=False)
                    else:
                        self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_tr_init"), 40.0)
                elif "detected language:" in lower_line:
                    lang_match = re.search(r"detected language:\s*([a-z]{2,3})(?:\s*\(([0-9.]+)\))?", lower_line)
                    if lang_match:
                        det_lang = lang_match.group(1).upper()
                        conf = f" ({int(float(lang_match.group(2)) * 100)} %)" if lang_match.group(2) else ""
                        self.update_file_bar(media.name, f"🌐 {det_lang}{conf}", 6.0)
                elif "performing voice activity detection" in lower_line:
                    current_stage = "vad"
                    self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.current_step_key = "step_vad"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_vad"), 4.0)
                elif "performing transcription" in lower_line:
                    current_stage = "transcribe"
                    self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.current_step_key = "step_tr_init"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_tr_init"), 8.0)
                elif "performing alignment" in lower_line:
                    current_stage = "align"
                    pct_align = 72.0 if use_diarize else 85.0
                    self.current_step_key = "step_align"; self.current_step_kwargs = {}; self.current_step_key = "step_align"; self.current_step_kwargs = {}; self.current_step_key = "step_align"; self.current_step_kwargs = {}; self.current_step_key = "step_align"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_align"), pct_align)
                    self.update_queue_row(media, self.tr("st_running", pct=int(pct_align)), scroll_to=False)
                elif "performing diarization" in lower_line or ("diarize" in lower_line and current_stage == "align"):
                    current_stage = "diarize"
                    self.current_step_key = "step_diar"; self.current_step_kwargs = {}; self.current_step_key = "step_diar"; self.current_step_kwargs = {}; self.current_step_key = "step_diar"; self.current_step_kwargs = {}; self.current_step_key = "step_diar"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_diar"), 82.0)
                    self.update_queue_row(media, self.tr("st_running", pct=82), scroll_to=False)

            self.current_process.wait()
            rc = self.current_process.returncode
            self.current_process = None

            if self.stop_requested:
                self.update_queue_row(media, self.tr("st_failed"), tag="failed")
                break

            if rc == 0 and json_temp.exists():
                with open(json_temp, "r", encoding="utf-8") as f:
                    data = json.load(f)
                merged_blocks = build_merged_blocks(data.get("segments", []))

                slides_temp_dir = None
                slides_grouped = []
                if need_slides:
                    self.current_step_key = "step_slides"; self.current_step_kwargs = {}; self.current_step_key = "step_slides"; self.current_step_kwargs = {}; self.current_step_key = "step_slides"; self.current_step_kwargs = {}; self.current_step_key = "step_slides"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_slides"), 90.0)
                    self.update_queue_row(media, self.tr("st_running", pct=90), scroll_to=False)
                    slides_temp_dir = Path(tempfile.mkdtemp(prefix="wx_slides_"))
                    raw_slides = extract_slides_from_video(media, total_duration, slides_temp_dir, self.log)
                    slides_grouped = assign_blocks_to_slides(raw_slides, data.get("segments", []))

                self.current_step_key = "step_doc"; self.current_step_kwargs = {}; self.current_step_key = "step_doc"; self.current_step_kwargs = {}; self.current_step_key = "step_doc"; self.current_step_kwargs = {}; self.current_step_key = "step_doc"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_doc"), 96.0)
                all_ok = True

                if want_txt:
                    out_p = self.build_output_filepath(media, effective_model, with_slides=False, ext="txt", diarized=use_diarize)
                    all_ok &= save_as_txt(media.stem, merged_blocks, out_p, self.log, self.print_diagnostic_card)
                if want_docx:
                    out_p = self.build_output_filepath(media, effective_model, with_slides=False, ext="docx", diarized=use_diarize)
                    all_ok &= save_as_docx(media.stem, merged_blocks, out_p, self.log, self.print_diagnostic_card)
                if want_docx_slides:
                    out_p = self.build_output_filepath(media, effective_model, with_slides=True, ext="docx", diarized=use_diarize)
                    all_ok &= save_as_docx_with_slides(media.stem, slides_grouped, merged_blocks, out_p, self.log, self.print_diagnostic_card)
                if want_pptx:
                    out_p = self.build_output_filepath(media, effective_model, with_slides=True, ext="pptx", diarized=use_diarize)
                    all_ok &= save_as_pptx(media.stem, slides_grouped, out_p, self.log, self.print_diagnostic_card)

                if slides_temp_dir and slides_temp_dir.exists():
                    shutil.rmtree(slides_temp_dir, ignore_errors=True)

                if all_ok:
                    for ext in [".json", ".srt", ".txt", ".vtt", ".tsv"]:
                        tmp = dossier_sortie / f"{media.stem}{ext}"
                        if tmp.exists():
                            tmp.unlink()
                    actual_elapsed = time.time() - t_start
                    tr_s, diar_s, sld_s = self.estimate_step_durations(media)
                    if tr_s is not None:
                        est_sum = (tr_s + (diar_s or 0.0) + (sld_s or 0.0)) / max(0.2, self.live_speed_multiplier)
                        if est_sum > 10.0:
                            measured_mult = actual_elapsed / est_sum
                            self.live_speed_multiplier = max(0.4, min(3.0, (self.live_speed_multiplier * 0.4) + (measured_mult * 0.6)))
                            self.after(0, self.refresh_all_queue_tasks)
                    elapsed_str = format_elapsed_buzz(actual_elapsed)
                    finished_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.current_step_key = "step_done"; self.current_step_kwargs = {}; self.current_step_key = "step_done"; self.current_step_kwargs = {}; self.current_step_key = "step_done"; self.current_step_kwargs = {}; self.current_step_key = "step_done"; self.current_step_kwargs = {}; self.update_file_bar(media.name, self.tr("step_done"), 100.0)
                    self.update_queue_row(media, self.tr("st_completed", dur=elapsed_str), finished=finished_ts, tag="done")
                    completed_files += 1
                    self.update_total_bar(completed_files, total_files)
                    self.after(0, self.update_queue_header_timer)
                else:
                    self.update_queue_row(media, self.tr("st_failed"), tag="failed")
            else:
                self.current_step_key = "step_err"; self.current_step_kwargs = {"rc": rc}; self.current_step_key = "step_err"; self.current_step_kwargs = {"rc": rc}; self.current_step_key = "step_err"; self.current_step_kwargs = {"rc": rc}; self.current_step_key = "step_err"; self.current_step_kwargs = {"rc": rc}; self.update_file_bar(media.name, self.tr("step_err", rc=rc), 0.0)
                self.update_queue_row(media, self.tr("st_failed"), tag="failed")
                self.log(f"\n[ÉCHEC BLOQUANT] WhisperX s'est arrêté avec le code {rc} sur : {media.name}\n", "error")
                diagnostics = diagnose_blocking_failure("".join(captured_lines))
                if diagnostics:
                    for diag in diagnostics:
                        self.print_diagnostic_card(diag["titre"], diag["etape"], diag["cause"], diag["action"])
                    if any("cache local" in d["titre"].lower() or "hugging face" in d["titre"].lower() for d in diagnostics):
                        self.log("[ARRÊT DU LOT] Traitement interrompu car un modèle requis manque dans le cache local.\n", "error")
                        break

        self.current_media_path = None
        self.is_running = False
        self.after(0, lambda: self.reset_ui_after_run(completed_files, total_files))

    def reset_ui_after_run(self, completed: int, total: int):
        self.btn_start.config(
            state=(tk.NORMAL if self.selected_files else tk.DISABLED),
            bg=THEME_BTN_PRIMARY_BG,
            cursor="hand2"
        )
        self.btn_models_mgr.config(
            state=tk.NORMAL,
            bg="#E5E7EB",
            fg="#111827",
            activebackground="#D1D5DB",
            activeforeground="#111827",
            cursor="hand2"
        )
        if getattr(self, "_model_mgr_win", None) and self._model_mgr_win.winfo_exists():
            for _m in list(self._model_mgr_win.row_widgets.keys()):
                self._model_mgr_win._refresh_single_row_state(_m)
        self.btn_pause.config(state=tk.DISABLED, text=" " + self.tr("btn_pause"))
        self.btn_stop.config(state=tk.DISABLED)
        self.update_queue_header_timer()
        if self.stop_requested:
            self.current_step_key = "step_stopped"; self.current_step_kwargs = {}; self.current_step_key = "step_stopped"; self.current_step_kwargs = {}; self.current_step_key = "step_stopped"; self.current_step_kwargs = {}; self.current_step_key = "step_stopped"; self.current_step_kwargs = {}; self.lbl_current_step.config(text=self.tr("step_stopped"))
        else:
            self.current_step_key = "step_batch_done"; self.current_step_kwargs = {}; self.current_step_key = "step_batch_done"; self.current_step_kwargs = {}; self.current_step_key = "step_batch_done"; self.current_step_kwargs = {}; self.current_step_key = "step_batch_done"; self.current_step_kwargs = {}; self.lbl_current_step.config(text=self.tr("step_batch_done"))
            self.log(f"\n✔ Traitement terminé ({completed}/{total} fichier(s) traité(s)).\n", "success")

if __name__ == "__main__":
    app = SlideScriptApp()
    app.mainloop()