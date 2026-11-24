# Architecture & Engineering Best Practices (SlideScript v1.0.0)

This blueprint defines the architectural, security, UI/UX, build, and release standards applied to **SlideScript v1.0.0**, adapted from the **SyncContactsOG** engineering guidelines.

---

## 1. Standard Repository Structure
- **Repository Root**: Uppercase GitHub meta files (`README.md`, `CHANGELOG.md`, `LICENSE`, `requirements.txt`, `.gitignore`), portable executable (`SlideScript.exe`), and user configuration (`slidescript_config.json`, gitignored).
- **`source/`**: Pure runtime application code (`SlideScript.py`, `run_engine_fast.py`, `SlideScript.ico`, `win_ca_bundle.pem`, `build.ps1`).
- **`tools/`**: Standalone developer & documentation automation scripts (`generate_specs_from_code.py`, `convert_md_2_html.py`), excluded from PyInstaller runtime bundles.
- **`docs/`**: Lowercase kebab-case Markdown sources (`specifications-en.md`, `specifications-fr.md`, `best-practices.md`) rendered into standalone offline HTML files (`specifications-en.html`, `specifications-fr.html`, `best-practices.html`, `changelog.html`).
- **`runtime/`, `ffmpeg/`, `hf_cache/`**: Gitignored portable Python 3.12 runtime, FFmpeg binaries, and offline AI model weights cache.
- **`releases/vX.Y.Z/`**: Gitignored release staging folder containing `-win64.zip`, `-source.zip`, `SHA256SUMS.txt`, and `RELEASE_NOTES.md`.

## 2. Portable Application & Security Standards
- **DPAPI-Protected Token Storage**: The optional Hugging Face Read token (`HF_TOKEN`) is encrypted with Windows DPAPI (`CryptProtectData` with entropy `SlideScript::HF_TOKEN::DPAPI::v1`) and stored as `REG_BINARY` under `HKCU\Software\SlideScript\HuggingFaceTokenDPAPI` without requiring administrator privileges.
- **Strict Enterprise TLS Validation**: Public root CA certificates from `Cert:\LocalMachine\Root` and `Cert:\CurrentUser\Root` are exported to `source/win_ca_bundle.pem` and injected via `REQUESTS_CA_BUNDLE` and `SSL_CERT_FILE` without disabling SSL verification.

## 3. Out-of-Tree Build & Release Pipeline (`source/build.ps1`)
- **Static Syntax Gate**: Runs `py_compile` before invoking PyInstaller.
- **Out-of-Tree Compilation**: Builds inside `C:\Temp\Build_SlideScript_v<Version>\` to prevent file-locking issues on synced drives and keep the working directory clean without needing a `backups/` folder.

## Windows Executable Metadata (VSVersionInfo)
- Always embed native Windows version resources (FileVersion, ProductVersion, ProductName, FileDescription) into SlideScript.exe via PyInstaller --version-file, derived dynamically from __version__ in source/SlideScript.py.
