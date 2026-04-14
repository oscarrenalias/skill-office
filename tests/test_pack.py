"""Tests for pypptx.ops.pack — unpack and pack operations."""
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation

from pypptx.ops.pack import CONTENT_TYPES_ENTRY, pack, unpack


class TestUnpack:
    def test_creates_destination_directory(self, minimal_pptx, tmp_path):
        out_dir = tmp_path / "output"
        result = unpack(minimal_pptx, out_dir)
        assert out_dir.is_dir()
        assert result == out_dir

    def test_content_types_xml_present(self, minimal_pptx, tmp_path):
        out_dir = tmp_path / "output"
        unpack(minimal_pptx, out_dir)
        assert (out_dir / "[Content_Types].xml").is_file()

    def test_pptx_structure_extracted(self, minimal_pptx, tmp_path):
        out_dir = tmp_path / "output"
        unpack(minimal_pptx, out_dir)
        assert (out_dir / "ppt" / "presentation.xml").is_file()
        assert (out_dir / "ppt" / "_rels" / "presentation.xml.rels").is_file()

    def test_returns_dest_path(self, minimal_pptx, tmp_path):
        out_dir = tmp_path / "output"
        result = unpack(minimal_pptx, out_dir)
        assert result == out_dir

    def test_raises_on_missing_src(self, tmp_path):
        with pytest.raises(ValueError, match="does not exist"):
            unpack(tmp_path / "nonexistent.pptx", tmp_path / "out")

    def test_raises_on_non_file_src(self, tmp_path):
        with pytest.raises(ValueError, match="not a file"):
            unpack(tmp_path, tmp_path / "out")

    def test_raises_on_non_zip_src(self, tmp_path):
        bad = tmp_path / "bad.pptx"
        bad.write_bytes(b"not a zip file at all")
        with pytest.raises(ValueError, match="not a valid ZIP"):
            unpack(bad, tmp_path / "out")

    def test_first_entry_ordering_preserved(self, minimal_pptx, tmp_path):
        """Original ZIP entries are all extracted."""
        out_dir = tmp_path / "output"
        unpack(minimal_pptx, out_dir)
        slides_dir = out_dir / "ppt" / "slides"
        assert slides_dir.is_dir()
        assert len(list(slides_dir.glob("slide*.xml"))) == 3


class TestPack:
    def test_round_trip_reopens_with_python_pptx(self, minimal_pptx, tmp_path):
        """Unpack → pack → re-open with python-pptx succeeds."""
        unpacked = tmp_path / "unpacked"
        unpack(minimal_pptx, unpacked)
        output = tmp_path / "repacked.pptx"
        pack(unpacked, output)
        prs = Presentation(str(output))
        assert len(prs.slides) == 3

    def test_content_types_is_first_entry(self, minimal_pptx, tmp_path):
        """[Content_Types].xml must be the first ZIP entry per the OPC spec."""
        unpacked = tmp_path / "unpacked"
        unpack(minimal_pptx, unpacked)
        output = tmp_path / "repacked.pptx"
        pack(unpacked, output)
        with zipfile.ZipFile(output, "r") as zf:
            assert zf.namelist()[0] == CONTENT_TYPES_ENTRY

    def test_returns_dest_path(self, minimal_pptx, tmp_path):
        unpacked = tmp_path / "unpacked"
        unpack(minimal_pptx, unpacked)
        output = tmp_path / "repacked.pptx"
        result = pack(unpacked, output)
        assert result == output

    def test_atomic_write_no_tmp_file_left(self, minimal_pptx, tmp_path):
        """No .tmp file remains after a successful pack."""
        unpacked = tmp_path / "unpacked"
        unpack(minimal_pptx, unpacked)
        output = tmp_path / "repacked.pptx"
        pack(unpacked, output)
        assert not any(tmp_path.glob("*.tmp"))

    def test_raises_on_missing_src_dir(self, tmp_path):
        with pytest.raises(ValueError, match="does not exist"):
            pack(tmp_path / "nonexistent_dir", tmp_path / "out.pptx")

    def test_raises_when_src_is_not_directory(self, minimal_pptx, tmp_path):
        with pytest.raises(ValueError, match="not a directory"):
            pack(minimal_pptx, tmp_path / "out.pptx")

    def test_raises_on_missing_content_types(self, tmp_path):
        empty_dir = tmp_path / "empty"
        empty_dir.mkdir()
        with pytest.raises(ValueError, match=r"\[Content_Types\]\.xml"):
            pack(empty_dir, tmp_path / "out.pptx")

    def test_raises_on_missing_dest_parent_dir(self, minimal_pptx, tmp_path):
        """pack raises when the output parent directory does not exist."""
        unpacked = tmp_path / "unpacked"
        unpack(minimal_pptx, unpacked)
        output = tmp_path / "nonexistent_dir" / "out.pptx"
        with pytest.raises(Exception):
            pack(unpacked, output)
