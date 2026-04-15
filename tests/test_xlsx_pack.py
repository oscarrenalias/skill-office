"""Tests for pyxlsx.ops.pack — unpack and pack operations."""
import pathlib

import openpyxl
import pytest

from pyxlsx.ops.pack import pack, unpack


# ── unpack() ──────────────────────────────────────────────────────────────────


class TestUnpack:
    def test_extracts_to_dest_dir(self, minimal_xlsx, tmp_path):
        dest = tmp_path / "unpacked"
        result = unpack(str(minimal_xlsx), str(dest))
        assert result["unpacked_dir"] == str(dest)
        assert dest.is_dir()

    def test_xlsx_structure_present(self, minimal_xlsx, tmp_path):
        dest = tmp_path / "unpacked"
        unpack(str(minimal_xlsx), str(dest))
        assert (dest / "xl" / "workbook.xml").exists()

    def test_default_output_dir_is_stem(self, minimal_xlsx):
        result = unpack(str(minimal_xlsx))
        expected = pathlib.Path(minimal_xlsx).parent / pathlib.Path(minimal_xlsx).stem
        assert result["unpacked_dir"] == str(expected)
        assert expected.is_dir()

    def test_missing_file_exits_1(self, tmp_path):
        with pytest.raises(SystemExit) as exc_info:
            unpack(str(tmp_path / "missing.xlsx"))
        assert exc_info.value.code == 1

    def test_bad_zip_exits_1(self, tmp_path):
        bad = tmp_path / "bad.xlsx"
        bad.write_bytes(b"not a zip file at all")
        with pytest.raises(SystemExit) as exc_info:
            unpack(str(bad))
        assert exc_info.value.code == 1


# ── pack() ────────────────────────────────────────────────────────────────────


class TestPack:
    def test_round_trip_openable_by_openpyxl(self, minimal_xlsx, tmp_path):
        unpacked = tmp_path / "unpacked"
        unpack(str(minimal_xlsx), str(unpacked))
        repacked = tmp_path / "repacked.xlsx"
        result = pack(str(unpacked), str(repacked))
        assert result["output_file"] == str(repacked)

        wb = openpyxl.load_workbook(str(repacked))
        assert "Data" in wb.sheetnames
        wb.close()

    def test_atomic_write_no_tmp_files_left(self, minimal_xlsx, tmp_path):
        unpacked = tmp_path / "unpacked"
        unpack(str(minimal_xlsx), str(unpacked))
        out_dir = tmp_path / "output"
        out_dir.mkdir()
        repacked = out_dir / "result.xlsx"

        pack(str(unpacked), str(repacked))

        tmp_files = list(out_dir.glob("*.tmp"))
        assert tmp_files == []

    def test_nonexistent_source_dir_exits_1(self, tmp_path):
        with pytest.raises(SystemExit) as exc_info:
            pack(str(tmp_path / "no_such_dir"), str(tmp_path / "out.xlsx"))
        assert exc_info.value.code == 1
