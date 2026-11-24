import argparse
import os
import sys

os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")
os.environ.setdefault("OMP_STACKSIZE", "8M")

import torch
import whisperx
from whisperx.diarize import DiarizationPipeline

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("media")
    parser.add_argument("--model", default="small")
    parser.add_argument("--language", default=None)
    parser.add_argument("--device", default="cpu")
    parser.add_argument("--compute_type", default="int8")
    parser.add_argument("--batch_size", type=int, default=8)
    parser.add_argument("--threads", type=int, default=0)
    parser.add_argument("--diarize", action="store_true")
    parser.add_argument("--hf_token", default=None)
    parser.add_argument("--output_format", default="json")
    parser.add_argument("--output_dir", required=True)
    args = parser.parse_args()

    logical = os.cpu_count() or 8
    optimal_threads = args.threads if args.threads > 0 else max(1, logical // 2)

    # CTranslate2 (Whisper) est optimal et stable à 4 threads + batch_size=8 sur Windows CPU
    ct2_threads = min(4, optimal_threads) if args.device == "cpu" else optimal_threads
    effective_batch = min(8, args.batch_size) if args.device == "cpu" else args.batch_size

    print("Performing voice activity detection using Pyannote...", flush=True)
    model = whisperx.load_model(
        args.model,
        device=args.device,
        compute_type=args.compute_type,
        language=args.language,
        vad_method="pyannote",
        threads=ct2_threads
    )

    if args.device == "cpu":
        torch.set_num_threads(optimal_threads)
        print(f"[PERF] Whisper threads={ct2_threads}, batch={effective_batch} | PyTorch/Diarisation threads={optimal_threads}/{logical}", flush=True)

    audio = whisperx.load_audio(args.media)
    print("Performing transcription...", flush=True)
    result = model.transcribe(audio, batch_size=effective_batch, print_progress=True)
    detected_lang = result.get("language", args.language or "en")
    print(f"Detected language: {detected_lang}", flush=True)

    del model
    if args.device == "cuda":
        torch.cuda.empty_cache()

    if args.diarize:
        print("Performing alignment...", flush=True)
        model_a, metadata = whisperx.load_align_model(language_code=detected_lang, device=args.device)
        result = whisperx.align(result["segments"], model_a, metadata, audio, args.device, return_char_alignments=False)
        del model_a

        print("Performing diarization...", flush=True)
        try:
            diarize_model = DiarizationPipeline(token=args.hf_token, device=args.device)
        except TypeError:
            diarize_model = DiarizationPipeline(use_auth_token=args.hf_token, device=args.device)

        try:
            diarize_model.model.segmentation.step = 0.25
            diarize_model.model.segmentation_batch_size = 32
            diarize_model.model.embedding_batch_size = 32
            print("[PERF] Diarisation accélérée activée (step=0.25, batch=32).", flush=True)
        except Exception:
            pass

        if args.device == "cpu":
            try:
                diarize_model.model = torch.quantization.quantize_dynamic(
                    diarize_model.model, {torch.nn.Linear}, dtype=torch.qint8
                )
                print("[PERF] Quantification INT8 activée sur pyannote CPU.", flush=True)
            except Exception:
                pass

        diarize_segments = diarize_model(audio)
        result = whisperx.assign_word_speakers(diarize_segments, result)

    from whisperx.utils import get_writer
    writer = get_writer("json", args.output_dir)
    result["language"] = detected_lang
    writer(result, args.media, {"max_line_width": None, "max_line_count": None, "highlight_words": False})

if __name__ == "__main__":
    main()