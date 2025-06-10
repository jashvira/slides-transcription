"""Command line interface for transcribing slide audio."""

import argparse
import os
import sys
import zipfile
from gooey import Gooey, GooeyParser
from transcriber.transcribe_slides import transcribe_slides, check_pptx_integrity


def build_parser():
    """Build command line argument parser."""
    parser = argparse.ArgumentParser(description="Transcribe audio from PowerPoint slides")
    parser.add_argument("pptx", nargs="+", help="Path to PowerPoint file(s) or directory containing PPTX files")
    parser.add_argument("output", help="Output directory for transcriptions")
    parser.add_argument("--model", default="whisper-1", help="OpenAI model to use")
    parser.add_argument("--language", default="en", help="Language code")
    parser.add_argument("--task", default="transcribe", choices=["transcribe", "translate"], help="Task to perform")
    parser.add_argument("--prefix", help="Prefix for output files (defaults to PPTX filename)")
    parser.add_argument("--api-key", help="OpenAI API key")
    parser.add_argument("--recursive", "-r", action="store_true", help="Process directories recursively")
    return parser


def find_pptx_files(paths: list, recursive: bool = False) -> list:
    """Find all PPTX files in the given paths."""
    pptx_files = []
    for path in paths:
        if os.path.isfile(path):
            if path.lower().endswith('.pptx'):
                pptx_files.append(path)
        elif os.path.isdir(path):
            if recursive:
                for root, _, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith('.pptx'):
                            pptx_files.append(os.path.join(root, file))
            else:
                for file in os.listdir(path):
                    if file.lower().endswith('.pptx'):
                        pptx_files.append(os.path.join(path, file))
    return sorted(pptx_files)


def count_valid_audio_slides(pptx_path):
    """Count slides with valid, non-corrupted audio."""
    print(f"Scanning for valid audio slides in: {os.path.basename(pptx_path)}", flush=True)
    is_valid, corrupted_files = check_pptx_integrity(pptx_path)
    valid_slide_indices = set()
    try:
        with zipfile.ZipFile(pptx_path) as z:
            for idx in range(1, 1000):
                rels_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
                if rels_name not in z.namelist():
                    break
                with z.open(rels_name) as rels_file:
                    import xml.etree.ElementTree as ET
                    tree = ET.parse(rels_file)
                    audio_tracks = [rel for rel in tree.getroot() if "audio" in rel.attrib.get("Type", "").lower()]
                    for rel in audio_tracks:
                        target = rel.attrib.get("Target", "")
                        media_path = target.replace("../", "ppt/")
                        if media_path not in corrupted_files:
                            valid_slide_indices.add(idx)
    except Exception:
        pass
    print(f"Found {len(valid_slide_indices)} valid slides in: {os.path.basename(pptx_path)}", flush=True)
    return len(valid_slide_indices)


@Gooey(
    program_name="Slides Transcriber",
    program_description="Extract audio from PowerPoint slides and transcribe them using OpenAI's Whisper model",
    default_size=(800, 600),
    tabbed_groups=True,
    navigation='TABBED'
)
def main():
    """Main entry point."""
    parser = build_parser()
    args = parser.parse_args()

    print("Scanning for PPTX files...", flush=True)
    # Find all PPTX files
    pptx_files = find_pptx_files(args.pptx, args.recursive)
    if not pptx_files:
        print("No PPTX files found.")
        return

    print(f"Found {len(pptx_files)} PPTX file(s):")
    for i, file in enumerate(pptx_files, 1):
        print(f"{i}. {file}")

    print("Counting valid slides with audio in each file...", flush=True)
    # Count valid slides with audio across all files
    total_slides = 0
    slides_per_file = []
    for pptx_file in pptx_files:
        count = count_valid_audio_slides(pptx_file)
        slides_per_file.append(count)
        total_slides += count
    if total_slides == 0:
        print("No slides with audio found in any file.")
        return

    slides_done = 0
    try:
        for i, (pptx_file, slide_count) in enumerate(zip(pptx_files, slides_per_file), 1):
            print(f"\nStarting transcription for file {i}/{len(pptx_files)}: {pptx_file}", flush=True)
            try:
                def progress_callback(current_slide=None, total_slides_in_file=None):
                    nonlocal slides_done, total_slides
                    slides_done += 1
                    percent = int(slides_done / total_slides * 100)
                    print(f"[PROGRESS] {percent}")
                    if current_slide is not None and total_slides_in_file is not None:
                        print(f"Processing file {i}/{len(pptx_files)}: {os.path.basename(pptx_file)} (Slide {current_slide}/{total_slides_in_file})", flush=True)
                transcribe_slides(
                    pptx_file,
                    args.output,
                    model=args.model,
                    language=args.language,
                    task=args.task,
                    prefix=args.prefix,
                    api_key=args.api_key,
                    progress_callback=progress_callback,
                    file_num=i,
                    total_files=len(pptx_files),
                    total_slides_in_file=slide_count
                )
                print(f"Finished transcription for file {i}/{len(pptx_files)}: {pptx_file}", flush=True)
            except Exception as e:
                print(f"Error processing {pptx_file}: {str(e)}")
                continue
        print("[PROGRESS] 100")
        print("All files processed!", flush=True)
    except KeyboardInterrupt:
        print("\nTranscription interrupted by user. Temporary files cleaned up.")
        return


if __name__ == "__main__":
    main()
