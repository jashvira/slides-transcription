import argparse


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Transcribe slide audio using Whisper")
    parser.add_argument("pptx", help="Path to PPTX file")
    parser.add_argument("output", help="Output directory for transcripts")
    parser.add_argument(
        "--model",
        default="base",
        help="Whisper model name or path (default: base)",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch a GUI instead of the CLI",
    )
    return parser


def parse_args(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    return args
