import os
import sys
import tempfile

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from transcriber.transcribe_slides import extract_slide_audio

TEST_PPTX = os.path.join(os.path.dirname(__file__), 'assets', 'multimedia.pptx')

def test_extract_slide_audio_no_tracks():
    assert os.path.isfile(TEST_PPTX), "sample PPTX missing"
    with tempfile.TemporaryDirectory() as tmpdir:
        tracks = list(extract_slide_audio(TEST_PPTX, tmpdir))
        assert tracks == []
        assert os.listdir(tmpdir) == []
