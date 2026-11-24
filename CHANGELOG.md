# Changelog

## [1.1.0] - 2026-11-24

### Added
- Reorganize the Installed Models & Cache Manager into 4 distinct sections: **Transcription Models (Whisper)**, **Forced Alignment Models (optional — improve timestamps)**, **Speaker Identification Models**, and **Active & Recent Operations**.
- Display real local disk usage (deduplicated from `blobs/`) for installed models, dynamically check Hugging Face Hub for newer revisions and remote download sizes, and highlight available updates with an active blue **Update** button and online size indicator.

### Changed
- Switch the French forced-alignment model to the fast, lightweight `bofenghuang/asr-wav2vec2-ctc-french` (~380 MB) and filter out unused Flax/TensorFlow weight files (`*.h5`, `*.msgpack`, `*.ot`) during downloads.
- Make French forced alignment optional (falling back seamlessly to `--no_align` when not cached) so offline transcription and speaker identification work immediately without extra downloads.
## [1.0.2] - 2026-11-24

### Added
- Embed native Windows version metadata (`FileVersion`, `ProductVersion`, `ProductName`, `FileDescription`) into `SlideScript.exe` via PyInstaller `VSVersionInfo` to ease version identification and reduce antivirus/SmartScreen false positives.

### Changed- Set `tiny` as the default Whisper model so first-time users can transcribe immediately without extra downloads.
- Pre-bundle the lightweight `tiny` Whisper model (~75 MB) and `pyannote` speaker diarization models (~35 MB) inside `.\hf_cache\` in the `win64` portable package for instant out-of-the-box offline execution.

### Fixed
- Monkey-patch `os.symlink` to fall back to `shutil.copy2` on Windows in both the main process and model download workers, completely eliminating `WinError 1314` privilege failures.
- Disable Hugging Face Hub symlinks (`HF_HUB_DISABLE_SYMLINKS=1`) to prevent `WinError 1314` privilege errors on standard Windows accounts and OneDrive folders.
## [1.0.1] - 2026-11-23

### Fixed
- Monkey-patch `os.symlink` to fall back to `shutil.copy2` on Windows in both the main process and model download workers, completely eliminating `WinError 1314` privilege failures.
- Bundle `runtime/` and `ffmpeg/` inside the portable `win64` release package so `SlideScript.exe` runs out of the box without missing environment errors.
- Force creation and usage of the local `.\hf_cache\` directory at startup so downloaded Whisper, WhisperX, and pyannote models persist reliably.
- Prevent duplicate main GUI window instantiation when opening the Installed Models & Updates manager.
## [1.0.0] - 2026-11-23
- Initial release of SlideScript v1.0.0.
