#!/usr/bin/env python3
"""
スキャンPDF → OCR → SQLite 格納スクリプト
【2段階処理】
  STEP 1: 最初の5ページでテスト → テキスト抽出確認
  STEP 2: 確認OKなら全ページ処理 → SQLiteに格納

必要なインストール:
  pip install ocrmypdf pdfplumber pypdf
  # macOS:  brew install tesseract tesseract-lang ghostscript
  # Ubuntu: sudo apt install tesseract-ocr tesseract-ocr-jpn ghostscript
  # Windows: https://ocrmypdf.readthedocs.io/en/latest/installation.html#windows
"""

import subprocess
import sqlite3
import sys
import time
from pathlib import Path

# =============================================
# ★ 設定：環境に合わせて変更してください
# =============================================
PDF_FOLDER  = r"C:\Users\j7214\Desktop\建築設備"  # PDFが格納されているフォルダ
DB_PATH     = r"C:\Users\j7214\Desktop\建築設備\建築設備.db"  # 出力SQLiteのパス
TEST_PAGES  = 5    # テストするページ数
PARALLEL_JOBS = 4  # 並列処理数（メモリ不足なら2に下げる）
# =============================================


def find_pdfs(folder: Path) -> list[Path]:
    """フォルダ内のPDFファイルを一覧取得"""
    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        print(f"❌ PDFファイルが見つかりません: {folder}")
        sys.exit(1)
    print(f"📂 対象PDFファイル ({len(pdfs)}件):")
    for p in pdfs:
        size_mb = p.stat().st_size / (1024 ** 2)
        print(f"   - {p.name}  ({size_mb:.1f} MB)")
    return pdfs


def extract_pages(input_pdf: Path, output_pdf: Path, start: int, end: int):
    """指定ページ範囲を抽出"""
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
    """OCR実行（画像PDF → 検索可能PDF）"""
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
        print(f"❌ OCRエラー (終了コード: {e.returncode})")
        sys.exit(1)


def extract_text_to_db(pdf_path: Path, conn: sqlite3.Connection,
                        source_name: str, page_offset: int = 0):
    """検索可能PDFからテキストを抽出してDBに格納"""
    import pdfplumber

    cursor = conn.cursor()
    count = 0

    with pdfplumber.open(str(pdf_path)) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            page_no = page_offset + i + 1
            text = page.extract_text() or ""
            text = text.strip()

            cursor.execute("""
                INSERT OR REPLACE INTO pages
                    (source_file, page_no, content, char_count)
                VALUES (?, ?, ?, ?)
            """, (source_name, page_no, text, len(text)))

            if (i + 1) % 10 == 0 or (i + 1) == total:
                print(f"  📄 {i+1}/{total} ページ処理済み", end="\r")

        conn.commit()
        count = total

    print(f"\n  ✅ {count} ページをDBに格納しました")
    return count


def setup_db(db_path: Path) -> sqlite3.Connection:
    """SQLiteのテーブル初期化"""
    conn = sqlite3.connect(str(db_path))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pages (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT    NOT NULL,
            page_no     INTEGER NOT NULL,
            content     TEXT,
            char_count  INTEGER,
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (source_file, page_no)
        )
    """)
    conn.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS pages_fts
        USING fts5(content, source_file, page_no UNINDEXED)
    """)
    conn.commit()
    print(f"✅ DB初期化完了: {db_path}")
    return conn


def rebuild_fts(conn: sqlite3.Connection):
    """全文検索インデックスを再構築"""
    conn.execute("DELETE FROM pages_fts")
    conn.execute("""
        INSERT INTO pages_fts (content, source_file, page_no)
        SELECT content, source_file, page_no FROM pages
        WHERE content != ''
    """)
    conn.commit()
    print("✅ 全文検索インデックスを再構築しました")


def show_sample(conn: sqlite3.Connection):
    """格納結果のサンプル表示"""
    rows = conn.execute("""
        SELECT source_file, page_no, substr(content, 1, 100) as preview
        FROM pages
        WHERE content != ''
        ORDER BY page_no
        LIMIT 5
    """).fetchall()

    print("\n📋 格納データのサンプル（先頭5件）:")
    print("─" * 60)
    for row in rows:
        print(f"  ファイル: {row[0]}")
        print(f"  ページ:   {row[1]}")
        print(f"  内容:     {row[2]}...")
        print("─" * 60)


def main():
    print("=" * 60)
    print("  スキャンPDF → OCR → SQLite 格納スクリプト")
    print("=" * 60)

    folder   = Path(PDF_FOLDER)
    db_path  = Path(DB_PATH)
    tmp_dir  = folder / "_tmp_ocr"
    tmp_dir.mkdir(exist_ok=True)

    # PDFファイル一覧取得
    pdfs = find_pdfs(folder)

    # DB初期化
    conn = setup_db(db_path)

    # ----------------------------
    # STEP 1: テスト処理（先頭5ページ）
    # ----------------------------
    print(f"\n{'─'*60}")
    print(f"  STEP 1: テスト処理（最初の {TEST_PAGES} ページ）")
    print(f"{'─'*60}")

    test_pdf    = pdfs[0]
    tmp_extract = tmp_dir / "_test_extract.pdf"
    tmp_ocr     = tmp_dir / "_test_ocr.pdf"

    print(f"\n対象: {test_pdf.name}")
    total_pages = extract_pages(test_pdf, tmp_extract, 1, TEST_PAGES)
    print(f"全ページ数: {total_pages}")

    print(f"\n⏳ OCR処理中（{TEST_PAGES}ページ）...")
    run_ocr(tmp_extract, tmp_ocr)

    print(f"\n⏳ テキスト抽出 → DB格納中...")
    extract_text_to_db(tmp_ocr, conn, source_name=test_pdf.name, page_offset=0)

    # クリーンアップ
    tmp_extract.unlink(missing_ok=True)
    tmp_ocr.unlink(missing_ok=True)

    show_sample(conn)

    print(f"\n{'='*60}")
    print(f"  ✅ テスト完了！")
    print(f"  DBファイル: {db_path}")
    print(f"\n  確認ポイント:")
    print(f"    - 上のサンプルに日本語テキストが表示されているか")
    print(f"    - 文字化けや空白が多すぎないか")
    print(f"{'='*60}")

    # ----------------------------
    # STEP 2: 本番処理への確認
    # ----------------------------
    ans = input("\n本番処理（全ページ・全PDFファイル）を開始しますか？ [y/N]: ").strip().lower()

    if ans != "y":
        print("\n⏸  処理を中断しました。サンプルを確認後、再度実行してください。")
        conn.close()
        sys.exit(0)

    # ----------------------------
    # STEP 2: 本番処理（全ページ）
    # ----------------------------
    print(f"\n{'─'*60}")
    print(f"  STEP 2: 本番処理（全PDFファイル・全ページ）")
    print(f"{'─'*60}")

    total_stored = 0
    start_time = time.time()

    for pdf_file in pdfs:
        print(f"\n📄 処理中: {pdf_file.name}")
        tmp_ocr_full = tmp_dir / f"_ocr_{pdf_file.stem}.pdf"

        print(f"  ⏳ OCR処理中...")
        run_ocr(pdf_file, tmp_ocr_full)

        print(f"  ⏳ テキスト抽出 → DB格納中...")
        count = extract_text_to_db(
            tmp_ocr_full, conn,
            source_name=pdf_file.name,
            page_offset=0
        )
        total_stored += count
        tmp_ocr_full.unlink(missing_ok=True)

    # 全文検索インデックス再構築
    rebuild_fts(conn)
    conn.close()

    # 一時フォルダ削除
    try:
        tmp_dir.rmdir()
    except Exception:
        pass

    elapsed = (time.time() - start_time) / 60
    print(f"\n{'='*60}")
    print(f"  🎉 全処理完了！")
    print(f"  処理時間:       {elapsed:.1f} 分")
    print(f"  格納ページ数:   {total_stored} ページ")
    print(f"  DBファイル:     {db_path}")
    print(f"{'='*60}")

    print(f"""
【DBの使い方サンプル】
import sqlite3
conn = sqlite3.connect(r"{db_path}")

# ページ検索
rows = conn.execute(
    "SELECT source_file, page_no, content FROM pages WHERE content LIKE ?",
    ('%空調%',)
).fetchall()

# 全文検索（高速）
rows = conn.execute(
    "SELECT source_file, page_no FROM pages_fts WHERE pages_fts MATCH ?",
    ('空調',)
).fetchall()
""")


if __name__ == "__main__":
    main()
