#!/usr/bin/env python3
"""
OCR済みPDF 軽量化スクリプト（OCR機能付き）
【処理フロー】
  1. PyMuPDFでページを300DPI JPEG画像に変換（サイズ削減）
  2. ocrmypdfで再OCRをかけて検索可能PDFを生成
  3. メタデータ削除・PDF構造最適化

必要なインストール:
  pip install pymupdf ocrmypdf
  Tesseract + jpn.traineddata が必要
"""

import sys
import time
from pathlib import Path

# =============================================
# ★ 設定
# =============================================
INPUT_FOLDER  = r"C:\Users\j7214\Desktop\建築設備"       # 処理対象PDFフォルダ
OUTPUT_FOLDER = r"C:\Users\j7214\Desktop\OCR_compressed" # 圧縮後の出力先

TARGET_DPI    = 400   # 解像度（元600DPI→400DPIで約55%削減）
JPEG_QUALITY  = 60    # JPEG品質（0〜100）
TEST_FILENAME = ""    # テスト対象ファイル名（空の場合は1ファイル目）
# =============================================


def compress_and_ocr(input_pdf: Path, output_pdf: Path) -> tuple:
    """
    画像を300DPIに圧縮 → ocrmypdfで再OCRをかけて検索可能PDFを生成
    戻り値: (元サイズMB, 圧縮後サイズMB)
    """
    import fitz
    import subprocess

    doc     = fitz.open(str(input_pdf))
    new_doc = fitz.open()
    zoom    = TARGET_DPI / 72
    mat     = fitz.Matrix(zoom, zoom)

    # STEP 1: ページを300DPI JPEG画像に変換
    for i, page in enumerate(doc):
        print(f"    ページ {i+1}/{len(doc)} 画像変換中...", end="\r")
        pix       = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        img_bytes = pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY)
        new_page  = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=img_bytes)

    print()

    # 一時ファイルに保存
    new_doc.set_metadata({})
    tmp_path = output_pdf.parent / f"_tmp_{output_pdf.name}"
    new_doc.save(str(tmp_path), garbage=4, deflate=True, clean=True)
    new_doc.close()
    doc.close()

    # STEP 2: ocrmypdfで再OCR → 検索可能PDFに変換
    print(f"    OCR処理中...", end="\r")

    # 出力ファイルが開かれている場合は削除
    if output_pdf.exists():
        try:
            output_pdf.unlink()
        except PermissionError:
            print(f"\n❌ 出力ファイルが開かれています。閉じてから再実行してください:")
            print(f"   {output_pdf}")
            tmp_path.unlink(missing_ok=True)
            raise

    try:
        subprocess.run([
            sys.executable, "-m", "ocrmypdf",
            "--language",    "jpn",
            "--output-type", "pdf",
            "--jobs",        "1",
            "--force-ocr",
            str(tmp_path),
            str(output_pdf),
        ], check=True, capture_output=True)
        print(f"    OCR完了     ")
    except subprocess.CalledProcessError as e:
        err = (e.stderr or b"").decode("utf-8", errors="ignore")
        print(f"\n    ❌ OCRエラー: {err[:300]}")
        tmp_path.unlink(missing_ok=True)
        raise
    except Exception as e:
        print(f"\n    ❌ 予期せぬエラー: {e}")
        tmp_path.unlink(missing_ok=True)
        raise
    finally:
        tmp_path.unlink(missing_ok=True)

    src_mb = input_pdf.stat().st_size  / (1024 ** 2)
    dst_mb = output_pdf.stat().st_size / (1024 ** 2)
    return src_mb, dst_mb


def find_pdfs(folder: Path) -> list[Path]:
    pdfs = sorted(
        p for p in folder.glob("*.pdf")
        if not p.name.startswith("~$") and not p.name.startswith("_")
    )
    if not pdfs:
        print(f"❌ PDFが見つかりません: {folder}")
        sys.exit(1)
    print(f"📂 対象PDFファイル ({len(pdfs)}件):")
    for p in pdfs:
        size_mb = p.stat().st_size / (1024 ** 2)
        print(f"   - {p.name}  ({size_mb:.1f} MB)")
    return pdfs


def main():
    print("=" * 55)
    print("  OCR済みPDF 軽量化スクリプト（OCR機能付き）")
    print("=" * 55)
    print(f"\n  設定:")
    print(f"    解像度:      {TARGET_DPI} DPI（元600DPI → 約55%削減）")
    print(f"    JPEG品質:    {JPEG_QUALITY}%")
    print(f"    OCR再処理:   あり（検索可能PDF維持）")
    print(f"    メタデータ:  削除")

    input_folder  = Path(INPUT_FOLDER)
    output_folder = Path(OUTPUT_FOLDER)
    output_folder.mkdir(exist_ok=True)

    pdfs = find_pdfs(input_folder)

    # テスト対象の決定
    if TEST_FILENAME:
        test_pdf = input_folder / TEST_FILENAME
        if not test_pdf.exists():
            print(f"❌ 指定ファイルが見つかりません: {test_pdf}")
            sys.exit(1)
    else:
        test_pdf = pdfs[0]

    # ----------------------------
    # STEP 1: テスト処理（1ファイル）
    # ----------------------------
    print(f"\n{'─'*55}")
    print(f"  STEP 1: テスト処理（{test_pdf.name}）")
    print(f"{'─'*55}\n")

    test_out = output_folder / f"_test_{test_pdf.name}"
    print(f"  ⏳ 圧縮＋OCR処理中...")
    t = time.time()

    try:
        src_mb, dst_mb = compress_and_ocr(test_pdf, test_out)
    except Exception as e:
        print(f"❌ エラー: {e}")
        sys.exit(1)

    elapsed = time.time() - t
    ratio   = (1 - dst_mb / src_mb) * 100

    print(f"\n{'='*55}")
    print(f"  ✅ テスト完了！")
    print(f"  元サイズ:   {src_mb:.2f} MB")
    print(f"  圧縮後:     {dst_mb:.2f} MB")
    print(f"  削減率:     {ratio:.1f}%")
    print(f"  処理時間:   {elapsed:.1f} 秒")
    print(f"  出力先:     {test_out}")
    print(f"\n  確認ポイント:")
    print(f"    - Ctrl+F でテキスト検索できるか ← OCR確認")
    print(f"    - 画質が許容範囲内か")
    print(f"    - ファイルサイズが十分小さいか")
    print(f"{'='*55}")

    print(f"\n  💡 画質調整が必要な場合:")
    print(f"    画質を上げたい → JPEG_QUALITY を {JPEG_QUALITY} より大きくする")
    print(f"    さらに小さくしたい → JPEG_QUALITY を {JPEG_QUALITY} より小さくする")

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
    total_src   = 0.0
    total_dst   = 0.0
    success, failed = 0, []

    for pdf in pdfs:
        out = output_folder / pdf.name
        print(f"\n📄 {pdf.name}")
        print(f"  ⏳ 圧縮＋OCR処理中...")
        t = time.time()
        try:
            src_mb, dst_mb = compress_and_ocr(pdf, out)
            ratio   = (1 - dst_mb / src_mb) * 100
            elapsed = time.time() - t
            print(f"  ✅ {src_mb:.2f}MB → {dst_mb:.2f}MB ({ratio:.1f}%削減) {elapsed:.1f}秒")
            total_src += src_mb
            total_dst += dst_mb
            success += 1
        except Exception as e:
            print(f"  ❌ エラー: {e}")
            failed.append(pdf.name)

    total_elapsed = (time.time() - total_start) / 60
    total_ratio   = (1 - total_dst / total_src) * 100 if total_src > 0 else 0

    print(f"\n{'='*55}")
    print(f"  🎉 全処理完了！")
    print(f"  処理時間:   {total_elapsed:.1f} 分")
    print(f"  成功:       {success} / {len(pdfs)} ファイル")
    print(f"  合計削減:   {total_src:.2f}MB → {total_dst:.2f}MB ({total_ratio:.1f}%削減)")
    if failed:
        for f in failed:
            print(f"    ❌ {f}")
    print(f"  出力先:     {output_folder}")
    print(f"{'='*55}")


if __name__ == "__main__":
    main()