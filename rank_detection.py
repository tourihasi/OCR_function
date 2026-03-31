#!/usr/bin/env python3
"""
PDFランク画像（A/B/C）検出 → Word挿入スクリプト

【処理フロー】
  STEP 1: 指定したPDFでテスト → ランク検出確認
  STEP 2: 確認OK → 全ファイル一括処理

必要なインストール:
  pip install opencv-python-headless pymupdf python-docx pywin32
"""

import sys
import subprocess
from pathlib import Path

# =============================================
# ★★★ ここを必ず設定してください ★★★
# =============================================

# 各フォルダのパス
PDF_FOLDER      = r"C:\Users\j7214\Desktop\建築設備"
OCR_FOLDER      = r"C:\Users\j7214\Desktop\OCR_output"
WORD_FOLDER     = r"C:\Users\j7214\Desktop\Word_output"
OUTPUT_FOLDER   = r"C:\Users\j7214\Desktop\Word_ranked"
TEMPLATE_FOLDER = r"C:\Users\j7214\Desktop\建築設備\templates"

# テスト対象ファイルを直接指定
# ★ ここにテストしたいPDFのファイル名を入力してください
TEST_FILENAME = "建築設備士問題集-016.pdf"   # ← ここを変更

# 照合設定
MATCH_THRESHOLD = 0.50  # 照合スコアの閾値（0.0〜1.0）
DPI             = 150   # PDF→画像変換のDPI

# =============================================


def load_templates(folder: Path) -> dict:
    """A/B/Cのテンプレート画像を読み込む"""
    import cv2
    import numpy as np
    from PIL import Image

    templates = {}
    for rank in ["A", "B", "C"]:
        path = folder / f"rank_{rank}.png"
        if not path.exists():
            print(f"  ⚠️  テンプレートなし: {path}")
            continue
        img = Image.open(path).convert("L")
        templates[rank] = np.array(img)
        print(f"  ✅ rank_{rank}.png 読み込み完了 {img.size}")
    return templates


def pdf_to_images(pdf_path: Path, dpi: int) -> list:
    """PyMuPDFでPDFを画像リストに変換（poppler不要）"""
    import fitz
    from PIL import Image
    import io

    doc = fitz.open(str(pdf_path))
    images = []
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        images.append(img)
    doc.close()
    return images


def detect_ranks(pdf_path: Path, templates: dict) -> dict:
    """PDFの各ページのランクを検出 → {ページ番号: ランク}"""
    import cv2
    import numpy as np

    print(f"  📷 PDF→画像変換中 (DPI={DPI})...")
    images = pdf_to_images(pdf_path, DPI)

    rank_map = {}
    scales = [0.5, 0.75, 1.0, 1.25, 1.5, 2.0]

    for i, img in enumerate(images):
        page_no  = i + 1
        img_gray = np.array(img.convert("L"))
        best_rank, best_score = None, 0.0

        for rank, tmpl in templates.items():
            for scale in scales:
                h = int(tmpl.shape[0] * scale)
                w = int(tmpl.shape[1] * scale)
                if h < 10 or w < 10:
                    continue
                if h > img_gray.shape[0] or w > img_gray.shape[1]:
                    continue
                resized = cv2.resize(tmpl, (w, h))
                result  = cv2.matchTemplate(img_gray, resized, cv2.TM_CCOEFF_NORMED)
                score   = result.max()
                if score > best_score:
                    best_score = score
                    best_rank  = rank

        if best_score >= MATCH_THRESHOLD:
            rank_map[page_no] = best_rank
            print(f"  🎯 ページ{page_no}: ランク{best_rank} (スコア:{best_score:.3f})")
        else:
            print(f"  　 ページ{page_no}: ランクなし (最高スコア:{best_score:.3f})")

    return rank_map


def run_ocr(src_pdf: Path, ocr_pdf: Path):
    """ocrmypdfでOCR処理"""
    print(f"  ⏳ OCR処理中...")
    subprocess.run([
        sys.executable, "-m", "ocrmypdf",
        "--language", "jpn",
        "--force-ocr",
        "--jobs", "1",
        str(src_pdf),
        str(ocr_pdf),
    ], check=True)
    print(f"  ✅ OCR完了")


def convert_to_word(ocr_pdf: Path, out_docx: Path, word_app) -> bool:
    """WordのCOMでOCR済みPDF→Word変換"""
    try:
        doc = word_app.Documents.Open(
            str(ocr_pdf),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        doc.SaveAs2(str(out_docx), FileFormat=12)
        doc.Close(SaveChanges=False)
        return True
    except Exception as e:
        print(f"  ❌ Word変換エラー: {e}")
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        return False


def insert_ranks_to_word(docx_path: Path, rank_map: dict, output_path: Path, word_app) -> bool:
    """Wordファイルにランク情報を挿入"""
    if not rank_map:
        print(f"  ⚠️  検出ランクなし。そのままコピーします。")
        import shutil
        shutil.copy(docx_path, output_path)
        return True

    try:
        doc = word_app.Documents.Open(
            str(docx_path),
            AddToRecentFiles=False,
        )
        total_pages = doc.ComputeStatistics(2)

        for page_no, rank in sorted(rank_map.items()):
            if page_no > total_pages:
                continue
            rng = doc.GoTo(1, 1, page_no)
            rng.Collapse(1)
            rng.InsertBefore(f"【ランク：{rank}】\n")

        doc.SaveAs2(str(output_path), FileFormat=12)
        doc.Close(SaveChanges=False)
        print(f"  ✅ 保存完了: {output_path.name}")
        return True

    except Exception as e:
        print(f"  ❌ Word操作エラー: {e}")
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        return False


def process_one(filename: str, word_app, templates: dict) -> bool:
    """1ファイルの処理（OCR → Word変換 → ランク挿入）"""
    stem     = Path(filename).stem
    src_pdf  = Path(PDF_FOLDER)  / filename
    ocr_pdf  = Path(OCR_FOLDER)  / filename
    docx     = Path(WORD_FOLDER) / f"{stem}.docx"
    out_docx = Path(OUTPUT_FOLDER) / f"{stem}.docx"

    if not src_pdf.exists():
        print(f"  ❌ PDFが見つかりません: {src_pdf}")
        return False

    # OCR処理（未処理の場合のみ）
    if not ocr_pdf.exists():
        run_ocr(src_pdf, ocr_pdf)
    else:
        print(f"  ✅ OCR済みPDFあり: {ocr_pdf.name}")

    # Word変換（未処理の場合のみ）
    if not docx.exists():
        print(f"  ⏳ Word変換中...")
        ok = convert_to_word(ocr_pdf, docx, word_app)
        if not ok:
            return False
    else:
        print(f"  ✅ Wordファイルあり: {docx.name}")

    # ランク検出
    rank_map = detect_ranks(ocr_pdf, templates)

    # ランク挿入
    return insert_ranks_to_word(docx, rank_map, out_docx, word_app)


def main():
    print("=" * 55)
    print("  PDFランク検出 → Word挿入スクリプト")
    print("=" * 55)

    # フォルダ作成
    for folder in [OCR_FOLDER, WORD_FOLDER, OUTPUT_FOLDER]:
        Path(folder).mkdir(exist_ok=True)

    # テンプレート読み込み
    print("\n📋 テンプレート読み込み中...")
    templates = load_templates(Path(TEMPLATE_FOLDER))
    if not templates:
        print(f"❌ テンプレートが見つかりません: {TEMPLATE_FOLDER}")
        print("   rank_A.png / rank_B.png / rank_C.png を配置してください")
        sys.exit(1)

    # Word起動
    print("\n⏳ Microsoft Word を起動中...")
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    print("✅ Word 起動完了")

    try:
        # ----------------------------
        # STEP 1: テスト処理
        # ----------------------------
        print(f"\n{'─'*55}")
        print(f"  STEP 1: テスト処理（{TEST_FILENAME}）")
        print(f"{'─'*55}\n")

        ok = process_one(TEST_FILENAME, word, templates)

        if not ok:
            print("❌ テスト失敗。")
            sys.exit(1)

        print(f"\n{'='*55}")
        print(f"  ✅ テスト完了！")
        print(f"  👉 以下のファイルを開いて確認してください:")
        print(f"     {OUTPUT_FOLDER}\\{Path(TEST_FILENAME).stem}.docx")
        print(f"\n  確認ポイント:")
        print(f"    - ページ先頭に【ランク：X】が挿入されているか")
        print(f"    - ランクの種類（A/B/C）が正しいか")
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
        pdfs = sorted(
            p for p in Path(PDF_FOLDER).glob("*.pdf")
            if not p.name.startswith("~$") and not p.name.startswith("_")
        )

        print(f"\n{'─'*55}")
        print(f"  STEP 2: 本番処理（全{len(pdfs)}ファイル）")
        print(f"{'─'*55}")

        success, failed = 0, []
        for pdf in pdfs:
            print(f"\n📄 {pdf.name}")
            ok = process_one(pdf.name, word, templates)
            if ok:
                success += 1
            else:
                failed.append(pdf.name)

        print(f"\n{'='*55}")
        print(f"  🎉 全処理完了！")
        print(f"  成功: {success} / {len(pdfs)} ファイル")
        if failed:
            for f in failed:
                print(f"    ❌ {f}")
        print(f"  出力先: {OUTPUT_FOLDER}")
        print(f"{'='*55}")

    finally:
        word.Quit()


if __name__ == "__main__":
    main()