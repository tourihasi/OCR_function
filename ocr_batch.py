#!/usr/bin/env python3
"""
スキャンPDF → 検索可能PDF 一括変換スクリプト
・フォルダ内の全PDFにOCRをかけて検索可能PDFとして保存
・2段階処理（テスト → 本番）

必要なインストール:
  pip install ocrmypdf
  Tesseract + jpn.traineddata が必要
"""

import sys
import time
from pathlib import Path

# =============================================
# ★ 設定
# =============================================
PDF_FOLDER    = r"/app/input"
OUTPUT_FOLDER = r"/app/OCR_output"
PARALLEL_JOBS = 4   # 並列数（メモリ不足なら2に下げる）
TEST_PAGES    = 5   # テストするページ数
# =============================================


def find_pdfs(folder: Path) -> list[Path]:
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


def extract_pages(input_pdf: Path, output_pdf: Path, start: int, end: int):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(str(input_pdf))
    writer = PdfWriter()
    total = len(reader.pages)
    end = min(end, total)
    for i in range(start - 1, end):
        writer.add_page(reader.pages[i])
    with open(output_pdf, "wb") as f:
        writer.write(f)
    return total


def run_ocr(input_path: Path, output_path: Path):
    import subprocess
    cmd = [
        sys.executable, "-m", "ocrmypdf",
        "--language",    "jpn",
        "--output-type", "pdf",
        "--optimize",    "1",
        "--jobs",        str(PARALLEL_JOBS),
        "--rotate-pages",
        "--deskew",
        "--skip-text",
        str(input_path),
        str(output_path),
    ]
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"  ❌ OCRエラー (終了コード: {e.returncode})")
        raise


def process_file(pdf_file: Path, output_folder: Path, tmp_dir: Path,
                 first_page: int = None, last_page: int = None) -> bool:
    prefix   = "_test_" if last_page else ""
    out_pdf  = output_folder / f"{prefix}{pdf_file.name}"
    tmp_extract = tmp_dir / f"_extract_{pdf_file.stem}.pdf" if last_page else None

    try:
        input_path = pdf_file

        # テスト時はページ抽出
        if last_page:
            print(f"  📄 先頭{last_page}ページを抽出中...")
            extract_pages(pdf_file, tmp_extract, first_page or 1, last_page)
            input_path = tmp_extract

        print(f"  ⏳ OCR処理中...")
        t = time.time()
        run_ocr(input_path, out_pdf)
        elapsed = time.time() - t
        size_mb = out_pdf.stat().st_size / (1024 ** 2)
        print(f"  ✅ 完了 — {elapsed:.1f}秒 / {size_mb:.1f} MB → {out_pdf.name}")
        return True

    except Exception as e:
        print(f"  ❌ エラー: {e}")
        return False

    finally:
        if tmp_extract:
            tmp_extract.unlink(missing_ok=True)


def main():
    print("=" * 55)
    print("  スキャンPDF → 検索可能PDF 一括変換スクリプト")
    print("=" * 55)

    pdf_folder    = Path(PDF_FOLDER)
    output_folder = Path(OUTPUT_FOLDER)
    tmp_dir       = pdf_folder / "_tmp"
    output_folder.mkdir(exist_ok=True)
    tmp_dir.mkdir(exist_ok=True)

    pdfs = find_pdfs(pdf_folder)

    # ----------------------------
    # STEP 1: テスト処理
    # ----------------------------
    print(f"\n{'─'*55}")
    print(f"  STEP 1: テスト処理（{pdfs[0].name} 先頭{TEST_PAGES}ページ）")
    print(f"{'─'*55}\n")

    ok = process_file(pdfs[0], output_folder, tmp_dir,
                      first_page=1, last_page=TEST_PAGES)
    if not ok:
        print("❌ テスト失敗。設定を確認してください。")
        sys.exit(1)

    test_out = output_folder / f"_test_{pdfs[0].name}"
    print(f"\n{'='*55}")
    print(f"  ✅ テスト完了！")
    print(f"  👉 以下のファイルを開いて確認してください:")
    print(f"     {test_out}")
    print(f"\n  確認ポイント:")
    print(f"    - Ctrl+F でテキスト検索できるか")
    print(f"    - ページの向きが正しく補正されているか")
    print(f"    - 文字認識が許容範囲内か")
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
        ok = process_file(pdf_file, output_folder, tmp_dir)
        if ok:
            success += 1
        else:
            failed.append(pdf_file.name)

    try:
        tmp_dir.rmdir()
    except Exception:
        pass

    elapsed = (time.time() - total_start) / 60
    print(f"\n{'='*55}")
    print(f"  🎉 全処理完了！")
    print(f"  処理時間:     {elapsed:.1f} 分")
    print(f"  成功:         {success} ファイル")
    if failed:
        print(f"  失敗:         {len(failed)} ファイル")
        for f in failed:
            print(f"    - {f}")
    print(f"  出力フォルダ: {output_folder}")
    print(f"{'='*55}")


if __name__ == "__main__":
    main()
