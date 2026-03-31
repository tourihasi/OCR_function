#!/usr/bin/env python3
"""
スキャンPDF → OCR済みPDF + Word 一括変換スクリプト
【処理フロー】
  STEP 1: ocrmypdf で OCR（向き補正・歪み補正）→ 検索可能PDF保存
  STEP 2: Microsoft Word COM で OCR済みPDF → .docx 保存

必要なインストール:
  pip install ocrmypdf pypdf
  Tesseract + jpn.traineddata が必要
  Microsoft Word がインストールされている必要があります
"""

import sys
import time
import subprocess
from pathlib import Path

# =============================================
# ★ 設定
# =============================================
PDF_FOLDER     = r"C:\Users\j7214\Desktop\建築設備"
OCR_FOLDER     = r"C:\Users\j7214\Desktop\OCR_output"   # OCR済みPDF出力先
WORD_FOLDER    = r"C:\Users\j7214\Desktop\Word_output"  # Word出力先
PARALLEL_JOBS  = 1    # 並列数（メモリ不足なら2に下げる）
TEST_PAGES     = 1    # テストするページ数
# =============================================


def find_pdfs(folder: Path) -> list[Path]:
    """一時ファイルを除外してPDF一覧取得"""
    pdfs = sorted(
        p for p in folder.glob("*.pdf")
        if not p.name.startswith("~$") and not p.name.startswith("_")
    )
    if not pdfs:
        print(f"❌ PDFファイルが見つかりません: {folder}")
        sys.exit(1)
    print(f"📂 対象PDFファイル ({len(pdfs)}件):")
    for p in pdfs:
        size_mb = p.stat().st_size / (1024 ** 2)
        print(f"   - {p.name}  ({size_mb:.1f} MB)")
    return pdfs


def extract_pages(input_pdf: Path, output_pdf: Path, last_page: int):
    """先頭Nページを抽出"""
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(str(input_pdf))
    writer = PdfWriter()
    for i in range(min(last_page, len(reader.pages))):
        writer.add_page(reader.pages[i])
    with open(output_pdf, "wb") as f:
        writer.write(f)


def run_ocr(input_pdf: Path, output_pdf: Path):
    """ocrmypdf で OCR処理"""
    cmd = [
        sys.executable, "-m", "ocrmypdf",
        "--language",    "jpn",
        "--output-type", "pdf",
        "--optimize",    "1",
        "--jobs",        str(PARALLEL_JOBS),
        "--force-ocr",
        str(input_pdf),
        str(output_pdf),
    ]
    subprocess.run(cmd, check=True)


def convert_to_word(ocr_pdf: Path, output_docx: Path, word_app) -> bool:
    """Word COM で OCR済みPDF → .docx 変換"""
    try:
        doc = word_app.Documents.Open(
            str(ocr_pdf),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        doc.SaveAs2(str(output_docx), FileFormat=12)
        doc.Close(SaveChanges=False)
        return True
    except Exception as e:
        print(f"  ❌ Word変換エラー: {e}")
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        return False


def process_file(pdf_file: Path, ocr_folder: Path, word_folder: Path,
                 tmp_dir: Path, word_app, test_pages: int = None) -> bool:
    """1ファイルの処理（OCR → PDF保存 → Word変換）"""
    prefix    = "_test_" if test_pages else ""
    tmp_clip  = tmp_dir / f"_clip_{pdf_file.stem}.pdf" if test_pages else None
    ocr_pdf   = ocr_folder / f"{prefix}{pdf_file.name}"
    out_docx  = word_folder / f"{prefix}{pdf_file.stem}.docx"

    try:
        input_path = pdf_file

        # テスト時はページ抽出（複数ページのみ）
        if test_pages:
            from pypdf import PdfReader
            total = len(PdfReader(str(pdf_file)).pages)
            if total > test_pages:
                print(f"  📄 先頭{test_pages}ページを抽出中（全{total}ページ）...")
                extract_pages(pdf_file, tmp_clip, test_pages)
                input_path = tmp_clip
            else:
                print(f"  📄 全{total}ページをそのまま処理（{test_pages}ページ以下）")
                tmp_clip = None  # 抽出不要

        # STEP 1: OCR → 検索可能PDF
        print(f"  ① OCR処理中 → {ocr_pdf.name}")
        t = time.time()
        run_ocr(input_path, ocr_pdf)
        print(f"     完了 ({time.time()-t:.1f}秒)")

        # STEP 2: OCR済みPDF → Word
        print(f"  ② Word変換中 → {out_docx.name}")
        t = time.time()
        ok = convert_to_word(ocr_pdf, out_docx, word_app)
        if not ok:
            return False
        elapsed = time.time() - t
        size_mb = out_docx.stat().st_size / (1024 ** 2)
        print(f"     完了 ({elapsed:.1f}秒 / {size_mb:.1f} MB)")
        return True

    except Exception as e:
        print(f"  ❌ エラー: {e}")
        return False

    finally:
        if tmp_clip:
            tmp_clip.unlink(missing_ok=True)


def main():
    print("=" * 55)
    print("  スキャンPDF → OCR済みPDF + Word 一括変換")
    print("=" * 55)

    # pywin32 確認
    try:
        import win32com.client
    except ImportError:
        print("❌ pywin32 が未インストールです")
        print("   実行: python -m pip install pywin32")
        sys.exit(1)

    pdf_folder  = Path(PDF_FOLDER)
    ocr_folder  = Path(OCR_FOLDER)
    word_folder = Path(WORD_FOLDER)
    tmp_dir     = pdf_folder / "_tmp"

    ocr_folder.mkdir(exist_ok=True)
    word_folder.mkdir(exist_ok=True)
    tmp_dir.mkdir(exist_ok=True)

    pdfs = find_pdfs(pdf_folder)

    # Word起動
    print("\n⏳ Microsoft Word を起動中...")
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        print("✅ Word 起動完了")
    except Exception as e:
        print(f"❌ Word の起動に失敗: {e}")
        sys.exit(1)

    try:
        # ----------------------------
        # STEP 1: テスト処理
        # ----------------------------
        print(f"\n{'─'*55}")
        print(f"  STEP 1: テスト処理（{pdfs[0].name} 先頭{TEST_PAGES}ページ）")
        print(f"{'─'*55}\n")

        ok = process_file(pdfs[0], ocr_folder, word_folder,
                          tmp_dir, word, test_pages=TEST_PAGES)
        if not ok:
            print("❌ テスト失敗。設定を確認してください。")
            sys.exit(1)

        print(f"\n{'='*55}")
        print(f"  ✅ テスト完了！")
        print(f"  👉 以下のファイルを開いて確認してください:")
        print(f"     PDF : {ocr_folder}\\_test_{pdfs[0].name}")
        print(f"     Word: {word_folder}\\_test_{pdfs[0].stem}.docx")
        print(f"\n  確認ポイント:")
        print(f"    - PDFでCtrl+F テキスト検索できるか")
        print(f"    - Wordで文字が正しく認識されているか")
        print(f"    - ページの向きが正しいか")
        print(f"{'='*55}")

        # ----------------------------
        # 本番処理への確認
        # ----------------------------
        ans = input("\n本番処理（全PDFファイル）を開始しますか？ [y/N]: ").strip().lower()
        if ans != "y":
            print("\n⏸  処理を中断しました。")
            sys.exit(0)

        # ----------------------------
        # STEP 2: 本番処理
        # ----------------------------
        print(f"\n{'─'*55}")
        print(f"  STEP 2: 本番処理（全{len(pdfs)}ファイル）")
        print(f"{'─'*55}")

        total_start = time.time()
        success, failed = 0, []

        for pdf_file in pdfs:
            print(f"\n📄 {pdf_file.name}")
            ok = process_file(pdf_file, ocr_folder, word_folder, tmp_dir, word)
            if ok:
                success += 1
            else:
                failed.append(pdf_file.name)

        elapsed = (time.time() - total_start) / 60
        print(f"\n{'='*55}")
        print(f"  🎉 全処理完了！")
        print(f"  処理時間:      {elapsed:.1f} 分")
        print(f"  成功:          {success} ファイル")
        if failed:
            print(f"  失敗:          {len(failed)} ファイル")
            for f in failed:
                print(f"    - {f}")
        print(f"  OCR出力:       {ocr_folder}")
        print(f"  Word出力:      {word_folder}")
        print(f"{'='*55}")

    finally:
        word.Quit()
        try:
            tmp_dir.rmdir()
        except Exception:
            pass


if __name__ == "__main__":
    main()