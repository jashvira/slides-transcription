import argparse


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Transcribe slide audio using the OpenAI Speech to Text API"
    )
    parser.add_argument("pptx", help="Path to PPTX file")
    parser.add_argument("output", help="Output directory for transcripts")
    parser.add_argument(
        "--model",
        default="whisper-1",
        help="OpenAI STT model name (default: whisper-1)",
    )
    parser.add_argument(
        "--language",
        default="en",
        help="Language spoken in the audio (default: en)",
    )
    parser.add_argument(
        "--task",
        default="transcribe",
        choices=["transcribe", "translate"],
        help="Task to perform (default: transcribe)",
    )
    parser.add_argument(
        "--prefix",
        default="slide",
        help="Filename prefix for transcripts (default: slide)",
    )
    parser.add_argument(
        "--api-key",
        help="OpenAI API key (defaults to OPENAI_API_KEY env var)",
        widget="PasswordField",
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
