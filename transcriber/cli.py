import argparse

try:
    from gooey import GooeyParser  # type: ignore
except Exception:
    GooeyParser = None


def build_parser(for_gui: bool = False) -> argparse.ArgumentParser:
    """Return an ArgumentParser or GooeyParser depending on ``for_gui``."""

    Parser = GooeyParser if for_gui and GooeyParser else argparse.ArgumentParser

    parser = Parser(
        description="Transcribe slide audio using the OpenAI Speech to Text API",
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
    if for_gui and GooeyParser:
        parser.add_argument(
            "--api-key",
            help="OpenAI API key (defaults to OPENAI_API_KEY env var)",
            widget="PasswordField",
        )
    else:
        parser.add_argument(
            "--api-key",
            help="OpenAI API key (defaults to OPENAI_API_KEY env var)",
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
