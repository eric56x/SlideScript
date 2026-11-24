import re, sys

def sync_specs(py_path: str, spec_fr_path: str, spec_en_path: str, version: str, today_iso: str):
    with open(py_path, "r", encoding="utf-8") as f:
        py_src = f.read()

    m_eng = re.search(r'ENGINE_BACKEND\s*=\s*["\']([^"\']+)["\']', py_src)
    engine_name = m_eng.group(1) if m_eng else "whisperx"

    model_ids = re.findall(r'\{"id":\s*"([^"]+)"', py_src)
    models_str = ", ".join(model_ids) if model_ids else "tiny, small, medium, large-v3"

    m_ext = re.search(r'EXTENSIONS_MEDIA\s*=\s*\{([^}]+)\}', py_src)
    if m_ext:
        exts = re.findall(r'["\'](\.[a-z0-9]+)["\']', m_ext.group(1))
        exts_str = ", ".join(exts) if exts else ".mp4, .mkv, .mov, .avi, .mp3, .wav, .m4a, .flac, .ogg"
    else:
        exts_str = ".mp4, .mkv, .mov, .avi, .mp3, .wav, .m4a, .flac, .ogg"

    with open(spec_fr_path, "r", encoding="utf-8") as f:
        fr = f.read()
    fr = re.sub(r'(?m)^(- \*\*Version courante\*\* : )`[^`]+`', rf'\1`{version}`', fr)
    fr = re.sub(r'(?m)^(- \*\*Date de mise à jour\*\* : )`[^`]+`', rf'\1`{today_iso}`', fr)
    fr = re.sub(r"(?m)^(- \*\*Noyau d'inférence actif\*\* : )`[^`]+`", rf'\1`{engine_name}`', fr)
    fr = re.sub(r'(?m)^(- \*\*Formats supportés\*\* : )`[^`]+`', rf'\1`{exts_str}`', fr)
    fr = re.sub(r'(?m)^(- \*\*Catalogue de modèles\*\* : )`[^`]+`', rf'\1`{models_str}`', fr)
    with open(spec_fr_path, "w", encoding="utf-8") as f:
        f.write(fr)

    with open(spec_en_path, "r", encoding="utf-8") as f:
        en = f.read()
    en = re.sub(r'(?m)^(- \*\*Current Version\*\* : )`[^`]+`', rf'\1`{version}`', en)
    en = re.sub(r'(?m)^(- \*\*Last Updated\*\* : )`[^`]+`', rf'\1`{today_iso}`', en)
    en = re.sub(r'(?m)^(- \*\*Active Inference Engine\*\* : )`[^`]+`', rf'\1`{engine_name}`', en)
    en = re.sub(r'(?m)^(- \*\*Supported Formats\*\* : )`[^`]+`', rf'\1`{exts_str}`', en)
    en = re.sub(r'(?m)^(- \*\*Model Catalog\*\* : )`[^`]+`', rf'\1`{models_str}`', en)
    with open(spec_en_path, "w", encoding="utf-8") as f:
        f.write(en)

    print(f"SPECS_SYNCED_FROM_CODE_OK (v{version}, engine={engine_name}, models={len(model_ids)})")

if __name__ == "__main__":
    sync_specs(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
