# USER MANUAL, TECHNICAL SPECIFICATION & REFERENCE SOURCE CODE

- **Current Version** : `1.1.0`
- **Last Updated** : `2026-11-24`
- **Active Inference Engine** : `whisperx`
- **Supported Formats** : `.mp4, .mkv, .mov, .avi, .mp3, .wav, .m4a, .flac, .ogg`
- **Model Catalog** : `tiny, tiny.en, small, small.en, medium, medium.en, large-v3, large.en, pyannote, wav2vec2-fr`

SlideScript (v1.0.0) — Transcription, Speaker Diarization & Slide Extraction
100% self-contained zero-install portable Windows desktop application: high-accuracy speech transcription, automatic speaker identification, lossless PNG slide extraction, native drag-and-drop into the queue, live mid-batch file addition, drag-and-drop queue reordering, and multi-format export.

## 1. Application Purpose & Quick Start User Guide

SlideScript is a self-contained Windows desktop application designed to automatically convert audio and video recordings (meetings, conferences, webinars, training sessions, interviews) into structured, timestamped, speaker-attributed transcripts without requiring any system-wide installation or Internet access during processing.

### 1.1 Core Capabilities & Use Cases

Fast, High-Accuracy Transcription: Converts speech to text using Whisper models (CTranslate2), automatically leveraging NVIDIA CUDA GPU acceleration or dynamically scaling across available CPU cores.
Speaker Identification (Diarization): Automatically distinguishes and numbers each speaker (Speaker 1, Speaker 2, ...) and merges consecutive speech segments with exact timestamps.
Lossless Slide Capture & Synchronization: For presentation videos (.mp4, .mkv, .mov, .avi), extracts unique slides in lossless PNG format once transition animations settle and aligns each speaker's remarks under the corresponding slide.
Ready-to-Share Deliverables: Generates Plain Text (.txt), formatted Word (.docx), illustrated Word with embedded slides (.docx), and PowerPoint (.pptx) decks with speaker notes directly alongside source media files.

### 1.2 Step-by-Step Operating Instructions

Step 1 — Launch the Portable Executable: Double-click .\SlideScript.exe from the portable folder (local drive, USB stick, or network share). Hardware capabilities (NVIDIA CUDA GPU vs. CPU) are automatically detected and displayed in the top dark status bar.
Step 2 — One-Time Initial Setup (If Needed): Click '⚙ Installed Models & Updates' in the top dark status bar to install, update, or uninstall Whisper models (tiny, small, medium, large-v3), speaker identification (pyannote/speaker-diarization-community-1), and French alignment (wav2vec2-large-xlsr-53-french) in parallel. To configure your Hugging Face Read token (hf_...), click the Hugging Face icon badge next to the pyannote row after accepting terms on both required pyannote repositories.
Step 3 — Select or Drag-and-Drop Media Files (Section 1 & Section 4): Click 'Select a file', 'Select multiple files', or 'Select a folder' (check 'Include subfolders' to scan nested directories), or drag and drop files/folders directly from Windows Explorer into the Section 4 queue area—even while a batch is already running.
Step 4 — Configure Output Formats & Settings (Section 2): Select one or more output formats (.txt, .docx, .docx with slides, .pptx with transcripts in notes), choose the audio language (en, fr, auto, etc.) and Whisper model, and toggle 'Identify speakers' as needed (transcription always runs 100% offline from .\hf_cache\).
Step 5 — Start & Control Processing Live: Click 'Start transcription'. While running, both 'Start transcription' and '⚙ Installed Models & Updates' are disabled to prevent concurrent cache modifications, while you can suspend execution ('Pause'), cleanly abort the batch ('Stop'), drag and drop additional files into Section 4, reorder pending items via the '≡' handle, cancel/reactivate rows (right-click or Delete), adjust CPU cores or switch the Whisper model live, and copy queue rows via Ctrl+C.

| System Parameter | Reference Specification |
| --- | --- |
| Executable & Version | .\SlideScript.exe — Version 1.0 beta (Self-Contained v1.0.0 Architecture) |
| Target Platform | Windows 10 / Windows 11 (x64) — Portable without administrator privileges |
| Supported Input Formats | Video: .mp4, .mkv, .mov, .avi  |  Audio: .mp3, .wav, .m4a, .flac, .ogg |
| Generated Output Formats | TXT, standard DOCX, DOCX with lossless PNG slide captures, PPTX with speaker notes |
| Typical Disk Footprint | ~4.2 GB (lightweight English setup) to ~7.7 GB (with 3.52 GB French Wav2Vec2 alignment & Whisper models) |


## 2. Technical Architecture & Relative Directory Layout

Deployment relies on a 100% self-contained relative directory layout bundling the executable, Python/PyInstaller build & runtime environment, model cache, and source files (.\source\) inside the single SlideScript folder:

| Relative Path (from Executable) | Technical Role & Responsibilities |
| --- | --- |
| .\SlideScript.exe | Main PyInstaller executable (--onefile --noconsole) bundling the Tkinter/Canvas UI, native Win32 drag-and-drop, Win32/Tk icon, and DOCX/PPTX generators. |
| .\runtime\ | Standalone Python 3.12.10 distribution containing pip, PyInstaller, PyTorch, CTranslate2, WhisperX, pyannote.audio, huggingface_hub, Pillow, python-docx, and python-pptx. |
| .\ffmpeg\ | Standalone ffmpeg.exe and ffprobe.exe binaries prepended dynamically to the process PATH at startup. |
| .\hf_cache\ | Local Hugging Face cache (HF_HUB_CACHE=.\hf_cache and TORCH_HOME=.\hf_cache\torch) storing Faster-Whisper, Pyannote, and Wav2Vec2 models. |
| .\run_engine_fast.py | Optimized Python runner executing the WhisperX pipeline in an isolated subprocess with unbuffered output (PYTHONUNBUFFERED=1). |
| .\slidescript_config.json | JSON configuration file automatically created beside the executable if absent, persisting UI language, output formats, audio language, model, CPU cores, and last folder. |
| Windows Registry (HKCU) | Secure storage of HF_TOKEN under HKEY_CURRENT_USER\Software\SlideScript (DPAPI) (never written in plaintext to .\slidescript_config.json). |
| .\source\SlideScript.py | Master Python source code for the graphical user interface (v1.0.0). |
| .\source\SlideScript.ico | Lightweight multi-resolution application icon (~8.15 KB: 16×16, 32×32 BMP + 48×48 PNG) embedded during compilation. |
| .\source\build.ps1 | 100% relative PowerShell build script invoking .\runtime\python.exe -m PyInstaller to rebuild .\SlideScript.exe with zero external dependencies. |


### 2.1 Window Icon Handling, Enterprise SSL Trust & PyInstaller Isolation

Dual Windows Icon Application: SlideScript.ico (multi-resolution 16×16, 32×32 BMP and 48×48 PNG totaling 8.15 KB) is embedded via --icon and --add-data and applied asynchronously (self.after(150, ...)) using win.iconbitmap(), win.iconphoto() (PIL ImageTk), and Win32 WM_SETICON (0x0080) messages.
Clean Restart Without _MEIPASS Conflicts: When restarting automatically after installing or uninstalling models, internal PyInstaller environment variables (_PYI_*, _MEIPASS2, TCL_LIBRARY, TK_LIBRARY) are stripped before launching a detached delayed process.

## 3. Graphical User Interface (GUI) & Visual Design System

The main interface comprises a full-width dark hardware status bar followed by 5 vertically stacked functional sections, fully bilingual (English / French) with instant live switching (including active processing step indicators). At startup, the Progress Log height matches the Section 4 queue table (6 visible rows).

| UI Component | Hex Color Specification | Visual Design & Behavior |
| --- | --- | --- |
| Top Status Bar (top_bar) | #1E2229 (Dark Charcoal) | Full-width header: clean white hardware text without leading emoji on the left; harmonized light gray buttons (#E5E7EB, text #111827) for 'Configure Hugging Face token (OK)' (with 825-byte grayscale Hugging Face micro-logo) (switches to crimson alert #B91C1C / white text if missing/invalid) and '⚙ Installed Models & Updates', plus FR/EN selector on the right. |
| Progress Bars (PillProgressBar) | #0078B8 on #DCE1E7 trough | Custom Canvas pill-shaped rounded progress bar used in Section 3 (Overall progress and Current file progress with centered percentage) and in the Installed Models manager. |
| Primary Action Button (Start) | #525864 (Dark Gray) + #FFFFFF Icon | Flat button featuring a crisp anti-aliased white play triangle (create_button_icons) matching the bold white button label. |
| 4 Secondary Control Buttons | #D8DCE2 (Light Gray) + #111827 Icons | The four buttons to the right of Start (Pause, Stop, System diagnostics, Clear log) share a uniform light gray fill (#D8DCE2) and solid black vector icons (#111827). |
| Progress Log Area (txt_log) | #1E1E1E on #E6E6E6 (height=6) | Initial height aligned with Section 4 (6 visible rows at startup) for balanced visual layout between the queue and log panels. |


## 4. Hardware Management, Live Controls & Queue Tracking

NVIDIA CUDA vs. Adaptive CPU Execution: At startup, detect_nvidia_cuda() checks nvidia-smi and torch.cuda.is_available(). On CUDA GPUs, --device cuda --compute_type float16 --batch_size 16 is used and the CPU selector is hidden. On CPU-only machines, --device cpu --compute_type int8 --batch_size 16 is used with a core selector ranging from 50% to 100% of logical cores.
Native Windows Drag-and-Drop into Section 4: Section 4 (frame_queue and tree_queue) natively hooks Windows WM_DROPFILES messages via ctypes/shell32.dll (enable_win32_drag_and_drop), allowing users to drag and drop files or folders directly into the queue table before or during batch processing.
Live Mid-Batch File Addition: Section 1 file selection buttons and Section 4 drag-and-drop remain active while a batch is running: newly added media files are appended to the live queue, probed by ffprobe in the background, and immediately reflected in the total progress bar and estimated finish time.
Dynamic Queue Reordering (≡) & Top-of-Stack Selection: The rightmost column features a '≡' drag handle allowing users to reorder queued rows by drag-and-drop at any time. After each file completes, the worker automatically picks the topmost unprocessed file in the stack.
Row Cancellation & Reactivation: Right-clicking a row (or pressing Delete) toggles 'Cancelled' status—skipping execution and excluding it from remaining time calculations—or reactivates it on demand.
Live CPU Affinity & Mid-Batch Model Switching: Changing CPU cores during execution immediately applies SetProcessAffinityMask to the active process and updates remaining estimates. Changing the Whisper model mid-batch updates the Model and Tasks columns for all pending files and takes effect automatically on the next file.
Comprehensive Time & ETA Tracking: The Section 4 header continuously displays elapsed time, estimated remaining duration, and predicted batch completion time (Est. finish: DD/MM HH:MM). The 'Finished' column displays 'est.: YYYY-MM-DD HH:MM:SS' for the active file before recording its actual completion timestamp.

## 5. Processing Pipeline, Diarization & Lossless PNG Slide Extraction


## 1. VAD & Faster-Whisper Transcription: Automatically selects the specialized <model>.en variant when language is set to 'en' and the model is present in .\hf_cache\.


## 2. Word Alignment & Pyannote Diarization: Performs word-level phonetic alignment (wav2vec2-large-xlsr-53-french for French) and speaker diarization via pyannote/speaker-diarization-community-1 and segmentation-3.0.


## 3. Lossless PNG Video Slide Extraction: FFmpeg extracts frames every 2 seconds in lossless PNG format (-vf fps=0.5,scale=iw:ih:flags=lanczos). Central content regions (5%–95% width, 6%–85% height) are compared via Pillow, capturing the stabilized frame after 3 consistent frames (6 s) so transition animations have completed. Embedded at 6.3 inches wide in Word and full-bleed 16:9 in PowerPoint.


| Output Format | Filename (Speakers Checked) | Filename (No Speakers) | Deliverable Structure & Content |
| --- | --- | --- | --- |
| Plain Text (txt) | <name>_<model>_speakers.txt | <name>_<model>.txt | Timestamped transcript grouped by speaker block with sentence line breaks. |
| Standard Word (docx) | <name>_<model>_speakers.docx | <name>_<model>.docx | Styled Word document (Aptos font) with color-coded speaker headers and timestamps. |
| Word with Slides (docx) | <name>_<model>_speakers_slides.docx | <name>_<model>_slides.docx | Lossless PNG slide captures (6.3" wide) followed by synchronized speaker remarks. |
| PowerPoint (pptx) | <name>_<model>_speakers_slides.pptx | <name>_<model>_slides.pptx | 16:9 presentation with full-bleed PNG slides and synchronized transcript in speaker notes. |


## 6. Source Code Modification Guide & Self-Contained Build Procedure

The application is engineered to be maintained and rebuilt 100% self-contained on any Windows computer, without requiring a system-wide Python installation or external project folder: the Python 3.12.10 console interpreter, pip, and PyInstaller reside directly inside .\runtime\.

### 6.1 How to Modify the Source Code (.\source\SlideScript.py)


## 1. Editing Source Code: Open .\source\SlideScript.py in any UTF-8 code editor (VS Code, Notepad++, Notepad). Always save using UTF-8 without BOM.


## 2. Key Code Landmarks: Bilingual UI strings are centralized in the I18N dictionary ('fr' and 'en'). Queue orchestration, Win32 drag-and-drop (enable_win32_drag_and_drop), live appending (append_files_to_queue), and UI widgets reside in WhisperXApp.


## 3. Syntax Pre-Check: Before rebuilding, verify Python syntax using the bundled interpreter: .\runtime\python.exe -m py_compile .\source\SlideScript.py


## 4. Adding New Python Dependencies: To install an additional Python package into the portable runtime: .\runtime\python.exe -m pip install --no-warn-script-location <package_name>


### 6.2 One-Command Self-Contained Rebuild

Open PowerShell in the SlideScript root folder (or right-click .\source\build.ps1 → Run with PowerShell) and run:

| powershell -ExecutionPolicy Bypass -File .\source\build.ps1 |
| --- |


| Parameter | Relative Resolution / Value |
| --- | --- |
| $PortableRoot | Directory containing .\SlideScript.exe (parent of .\source\). |
| $SourceDir | Source directory: Join-Path $PortableRoot 'source' (i.e. .\source\). |
| $RuntimePy | Bundled console interpreter: Join-Path $PortableRoot 'runtime\python.exe'. |
| $GuiFile | Master Python source code: Join-Path $SourceDir 'SlideScript.py'. |
| $BuildIcon | Lightweight multi-resolution icon (8.15 KB): Join-Path $SourceDir 'SlideScript.ico'. |
| Build Command | powershell -ExecutionPolicy Bypass -File .\source\build.ps1 |


### 6.3 Complete Contents of .\source\build.ps1 (100% Self-Contained Build Script)


| $SourceDir    = $PSScriptRoot $PortableRoot = Split-Path $SourceDir -Parent $RuntimePy    = Join-Path $PortableRoot 'runtime\python.exe' $GuiFile      = Join-Path $SourceDir 'SlideScript.py' $BuildIcon    = Join-Path $SourceDir 'SlideScript.ico' $BuildTemp    = Join-Path $PortableRoot 'build_temp'  Get-Process -Name 'WhisperX_Transcription' -ErrorAction SilentlyContinue | Stop-Process -Force Start-Sleep -Milliseconds 300 Remove-Item -Force (Join-Path $PortableRoot 'SlideScript.exe') -ErrorAction SilentlyContinue  if (Test-Path $BuildTemp) { Remove-Item -Recurse -Force -LiteralPath $BuildTemp } New-Item -ItemType Directory -Path $BuildTemp -Force | Out-Null  $ArgsList = @(     '-m', 'PyInstaller',     '--clean', '--noconsole', '--onefile',     '--collect-all', 'pptx',     '--collect-all', 'lxml',     '--hidden-import=docx',     '--hidden-import=PIL',     '--hidden-import=PIL.Image',     '--hidden-import=PIL.ImageChops',     '--hidden-import=PIL.ImageDraw',     '--hidden-import=PIL.ImageStat',     '--hidden-import=PIL.ImageTk',     '--name', 'WhisperX_Transcription',     '--distpath', $PortableRoot,     '--workpath', (Join-Path $BuildTemp 'work'),     '--specpath', (Join-Path $BuildTemp 'spec') ) if (Test-Path $BuildIcon) { $ArgsList += @("--icon=$BuildIcon", "--add-data=$BuildIcon;.") } $ArgsList += $GuiFile  & $RuntimePy @ArgsList Remove-Item -Recurse -Force -LiteralPath $BuildTemp -ErrorAction SilentlyContinue |
| --- |


### 6.4 Parallel Model Manager, Enterprise SSL Trust & Automated Build Backups (v1.0.0)

- Parallel Model Operations & Live Telemetry: The '⚙ Installed Models & Updates' dialog supports concurrent installation, update, and uninstallation of Whisper, Pyannote, and French Wav2Vec2 models using non-blocking subprocesses polled by the Tkinter event loop, displaying real-time progress bars, MB/s transfer speed, ETA, and per-task 'Stop' buttons.
- Automatic Remote Update Detection: Upon opening the model manager, a lightweight background HfApi check compares local refs/main commit SHAs against Hugging Face; when a newer revision is available, the 'Update' button turns blue ('Update available').
- Strict Enterprise SSL Validation (win_ca_bundle.pem): Public root CA certificates from the Windows certificate store (Cert:\LocalMachine\Root and Cert:\CurrentUser\Root) are exported to .\source\win_ca_bundle.pem and injected via REQUESTS_CA_BUNDLE and SSL_CERT_FILE so HTTPS requests succeed behind enterprise TLS inspection proxies without disabling SSL verification.
- Active Transcription Lockout: Once 'Start transcription' is clicked, both 'Start transcription' and '⚙ Installed Models & Updates' are locked (state=DISABLED) until the batch completes or is stopped, preventing cache contention while WhisperX is running.
- Timestamped Build Backups & Automatic Rollback: Before rebuilding .\SlideScript.exe, .\source\build.ps1 archives the previous executable to .\backups\WhisperX_Transcription_YYYYMMDD_HHMMSS.exe (retaining the 5 most recent builds) and automatically restores it if PyInstaller fails.