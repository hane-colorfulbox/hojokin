# -*- coding: utf-8 -*-
"""セカンドオピニオン加点の社内MTG録画を文字起こし（faster-whisper large-v3-turbo / CPU・int8）。
ローカル完結・Anthropic API は呼ばない。出力: SRT + タイムスタンプ付きtxt。
docs/案件メモ/申請MTG_リファレンス/_scripts/transcribe.py を流用。"""
import sys, subprocess, time, pathlib
sys.stdout.reconfigure(encoding="utf-8")

OUT = pathlib.Path(__file__).resolve().parent.parent / "_transcripts"
OUT.mkdir(parents=True, exist_ok=True)

JOBS = [
    ("社内MTG_2026-06-23", pathlib.Path(r"D:/user/Videos/レコーディング 2026-06-23 112947.mp4")),
]


def fmt_ts(sec: float) -> str:
    h = int(sec // 3600); m = int((sec % 3600) // 60); s = sec % 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def fmt_srt(sec: float) -> str:
    h = int(sec // 3600); m = int((sec % 3600) // 60); s = int(sec % 60); ms = int((sec - int(sec)) * 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def extract_wav(mp4: pathlib.Path, wav: pathlib.Path):
    if wav.exists():
        print(f"  wav既存: {wav.name}", flush=True)
        return
    print(f"  音声抽出中: {mp4.name}", flush=True)
    subprocess.run([
        "ffmpeg", "-y", "-i", str(mp4), "-vn", "-ac", "1", "-ar", "16000",
        "-c:a", "pcm_s16le", str(wav)
    ], check=True, capture_output=True)


def main():
    from faster_whisper import WhisperModel
    print("モデル読込中: large-v3-turbo (cpu/int8/16threads)...", flush=True)
    t0 = time.time()
    model = WhisperModel(
        "mobiuslabsgmbh/faster-whisper-large-v3-turbo",
        device="cpu", compute_type="int8", cpu_threads=16, local_files_only=True,
    )
    print(f"モデル読込完了 ({time.time()-t0:.1f}s)", flush=True)

    for name, mp4 in JOBS:
        if not mp4.exists():
            print(f"!! ファイル無し: {mp4}", flush=True); continue
        print(f"\n===== {name} =====", flush=True)
        wav = OUT / f"{name}.wav"
        extract_wav(mp4, wav)

        t1 = time.time()
        segments, info = model.transcribe(
            str(wav), language="ja", task="transcribe",
            beam_size=5, vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=500),
            condition_on_previous_text=True,
        )
        print(f"  検出言語={info.language} 音声長={info.duration:.0f}s 文字起こし開始", flush=True)

        srt_lines, txt_lines = [], []
        n = 0
        for seg in segments:
            n += 1
            text = seg.text.strip()
            srt_lines.append(f"{n}\n{fmt_srt(seg.start)} --> {fmt_srt(seg.end)}\n{text}\n")
            txt_lines.append(f"[{fmt_ts(seg.start)}] {text}")
            if n % 25 == 0:
                print(f"    ...{n}セグメント (再生位置 {fmt_ts(seg.end)}, 経過 {time.time()-t1:.0f}s)", flush=True)

        (OUT / f"{name}.srt").write_text("\n".join(srt_lines), encoding="utf-8")
        (OUT / f"{name}.txt").write_text("\n".join(txt_lines), encoding="utf-8")
        try:
            wav.unlink()
        except OSError:
            pass
        print(f"  完了: {n}セグメント / {time.time()-t1:.0f}s -> {name}.txt / .srt", flush=True)

    print("\n全て完了", flush=True)


if __name__ == "__main__":
    main()
