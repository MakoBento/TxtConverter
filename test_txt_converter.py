"""
TxtConverter のユニットテスト
"""
import os
import re
import sys
import tempfile
from pathlib import Path

import pytest

# テスト対象モジュールのインポート
import txt_converter as tc


# ---------------------------------------------------------------------------
# テキスト系ファイル変換のテスト
# ---------------------------------------------------------------------------

class TestConvertTextFile:
    """テキスト系ファイル（Excel以外）の変換テスト"""

    def test_cs_file_converted(self, tmp_path):
        """C# ファイルがtxtに変換される"""
        src = tmp_path / "sample.cs"
        src.write_text("public class Hello {}", encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        result = tc.convert_file(str(src), str(dst), lambda msg: None)

        assert result is True
        txt_files = list(dst.glob("*.txt"))
        assert len(txt_files) == 1
        assert txt_files[0].read_text(encoding="utf-8") == "public class Hello {}"

    def test_py_file_converted(self, tmp_path):
        """Python ファイルがtxtに変換される"""
        src = tmp_path / "script.py"
        src.write_text("print('hello')", encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        result = tc.convert_file(str(src), str(dst), lambda msg: None)

        assert result is True
        txt_files = list(dst.glob("*.txt"))
        assert len(txt_files) == 1
        assert txt_files[0].read_text(encoding="utf-8") == "print('hello')"

    def test_output_filename_is_original_name_plus_txt(self, tmp_path):
        """出力ファイル名は元のファイル名（拡張子含む）に.txtを付加した形式 (例: mycode.ts.txt)"""
        src = tmp_path / "mycode.ts"
        src.write_text("const x = 1;", encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        tc.convert_file(str(src), str(dst), lambda msg: None)

        txt_files = list(dst.glob("*.txt"))
        assert len(txt_files) == 1
        assert txt_files[0].name == "mycode.ts.txt"

    def test_all_text_extensions_converted(self, tmp_path):
        """テキスト系全拡張子が変換対象になる"""
        dst = tmp_path / "output"
        dst.mkdir()

        text_extensions = [
            ".cs", ".md", ".txt", ".sql", ".py", ".js",
            ".html", ".htm", ".ts", ".tsx", ".css", ".vue", ".json", ".xml"
        ]

        for ext in text_extensions:
            src = tmp_path / f"file{ext}"
            src.write_text(f"content for {ext}", encoding="utf-8")
            result = tc.convert_file(str(src), str(dst), lambda msg: None)
            assert result is True, f"{ext} ファイルの変換が失敗した"

    def test_unsupported_extension_skipped(self, tmp_path):
        """非対応拡張子はスキップされてFalseを返す"""
        src = tmp_path / "file.docx"
        src.write_text("doc content", encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        result = tc.convert_file(str(src), str(dst), lambda msg: None)

        assert result is False
        assert list(dst.glob("*.txt")) == []

    def test_log_message_on_success(self, tmp_path):
        """変換成功時にログメッセージが記録される"""
        src = tmp_path / "sample.md"
        src.write_text("# Title", encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        logs = []
        tc.convert_file(str(src), str(dst), logs.append)

        assert any("[完了]" in msg for msg in logs)

    def test_log_message_on_skip(self, tmp_path):
        """スキップ時にログメッセージが記録される"""
        src = tmp_path / "file.docx"
        src.write_text("content", encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        logs = []
        tc.convert_file(str(src), str(dst), logs.append)

        assert any("[スキップ]" in msg for msg in logs)

    def test_json_file_content_preserved(self, tmp_path):
        """JSON ファイルの内容がそのまま保持される"""
        content = '{"key": "value", "num": 42}'
        src = tmp_path / "data.json"
        src.write_text(content, encoding="utf-8")
        dst = tmp_path / "output"
        dst.mkdir()

        tc.convert_file(str(src), str(dst), lambda msg: None)

        txt_files = list(dst.glob("*.txt"))
        assert txt_files[0].read_text(encoding="utf-8") == content


# ---------------------------------------------------------------------------
# タイムスタンプサブフォルダ作成のテスト
# ---------------------------------------------------------------------------

class TestCreateOutputSubdir:
    """出力先タイムスタンプサブフォルダ作成のテスト"""

    def test_subdir_is_created(self, tmp_path):
        """TxtConvert_で始まるサブフォルダが作成される"""
        subdir = tc.create_output_subdir(str(tmp_path))
        assert os.path.isdir(subdir)

    def test_subdir_name_format(self, tmp_path):
        """サブフォルダ名が TxtConvert_yyyyMMdd_HHmmss 形式である"""
        subdir = tc.create_output_subdir(str(tmp_path))
        name = os.path.basename(subdir)
        assert re.match(r"^TxtConvert_\d{8}_\d{6}$", name), f"フォルダ名形式が不正: {name}"

    def test_subdir_is_inside_base_dir(self, tmp_path):
        """サブフォルダが指定ベースディレクトリの直下に作成される"""
        subdir = tc.create_output_subdir(str(tmp_path))
        assert Path(subdir).parent == tmp_path


# ---------------------------------------------------------------------------
# フォルダ対象ファイル収集のテスト
# ---------------------------------------------------------------------------

class TestCollectTargetFiles:
    """フォルダ内の対象ファイル収集のテスト"""

    def test_collect_excel_files(self, tmp_path):
        """Excel ファイルが収集される"""
        (tmp_path / "book.xlsx").write_bytes(b"dummy")
        (tmp_path / "book2.xlsm").write_bytes(b"dummy")

        files = tc.collect_target_files(str(tmp_path))

        names = [os.path.basename(f) for f in files]
        assert "book.xlsx" in names
        assert "book2.xlsm" in names

    def test_collect_text_files(self, tmp_path):
        """テキスト系ファイルが収集される"""
        (tmp_path / "code.py").write_text("x=1", encoding="utf-8")
        (tmp_path / "readme.md").write_text("# hi", encoding="utf-8")
        (tmp_path / "style.css").write_text("body{}", encoding="utf-8")

        files = tc.collect_target_files(str(tmp_path))

        names = [os.path.basename(f) for f in files]
        assert "code.py" in names
        assert "readme.md" in names
        assert "style.css" in names

    def test_collect_excludes_unsupported(self, tmp_path):
        """非対応拡張子は収集されない"""
        (tmp_path / "doc.docx").write_bytes(b"dummy")
        (tmp_path / "img.png").write_bytes(b"dummy")
        (tmp_path / "code.py").write_text("x=1", encoding="utf-8")

        files = tc.collect_target_files(str(tmp_path))

        names = [os.path.basename(f) for f in files]
        assert "doc.docx" not in names
        assert "img.png" not in names
        assert "code.py" in names

    def test_collect_returns_empty_for_empty_folder(self, tmp_path):
        """空フォルダでは空リストを返す"""
        files = tc.collect_target_files(str(tmp_path))
        assert files == []

    def test_collect_includes_subfolders(self, tmp_path):
        """サブフォルダ内のファイルも再帰的に収集される"""
        sub = tmp_path / "sub"
        sub.mkdir()
        (sub / "deep.py").write_text("x=1", encoding="utf-8")
        (tmp_path / "root.js").write_text("var x=1", encoding="utf-8")

        files = tc.collect_target_files(str(tmp_path))

        names = [os.path.basename(f) for f in files]
        assert "deep.py" in names
        assert "root.js" in names

    def test_collect_includes_nested_subfolders(self, tmp_path):
        """ネストされたサブフォルダ内のファイルも収集される"""
        nested = tmp_path / "a" / "b" / "c"
        nested.mkdir(parents=True)
        (nested / "deep.ts").write_text("const x=1", encoding="utf-8")

        files = tc.collect_target_files(str(tmp_path))

        names = [os.path.basename(f) for f in files]
        assert "deep.ts" in names


# ---------------------------------------------------------------------------
# Excel変換の既存テスト（回帰テスト）
# ---------------------------------------------------------------------------

class TestExcelConversionRegression:
    """Excel変換機能が引き続き動作することを確認する回帰テスト"""

    def test_excel_extension_check(self):
        """Excel拡張子が EXCEL_EXTENSIONS に含まれている"""
        assert ".xlsx" in tc.EXCEL_EXTENSIONS
        assert ".xlsm" in tc.EXCEL_EXTENSIONS
        assert ".xls" in tc.EXCEL_EXTENSIONS

    def test_text_extension_check(self):
        """テキスト系拡張子が TEXT_EXTENSIONS に含まれている"""
        expected = [
            ".cs", ".md", ".txt", ".sql", ".py", ".js",
            ".html", ".htm", ".ts", ".tsx", ".css", ".vue", ".json", ".xml"
        ]
        for ext in expected:
            assert ext in tc.TEXT_EXTENSIONS, f"{ext} が TEXT_EXTENSIONS に含まれていない"
