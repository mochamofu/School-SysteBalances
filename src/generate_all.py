# -*- coding: utf-8 -*-
"""
学次会計 業務改善ツール ─ 一括生成スクリプト
実行方法: python src/generate_all.py

4つのExcelファイルを output/ フォルダに生成します。
"""

import os, sys, time

# src/ ディレクトリを path に追加（どこから実行しても動くように）
sys.path.insert(0, os.path.dirname(__file__))

# output フォルダを作成
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

print("=" * 60)
print("  学次会計 業務改善ツール ─ Excel ファイル生成")
print("=" * 60)

# ── ツール① 番号紐付けテンプレート ──────────────────────────────────────────
print("\n[1/4] 番号紐付けテンプレート.xlsx を生成中...")
from tool1_class_linking import create_workbook as gen1
gen1(os.path.join(OUTPUT_DIR, "番号紐付けテンプレート.xlsx"))

# ── ツール② CoPilot プロンプトガイド ─────────────────────────────────────────
print("[2/4] CoPilotプロンプトガイド_学次会計用.xlsx を生成中...")
from tool2_copilot_guide import create_workbook as gen2
gen2(os.path.join(OUTPUT_DIR, "CoPilotプロンプトガイド_学次会計用.xlsx"))

# ── ツール③ マスターブック ───────────────────────────────────────────────────
print("[3/4] master_マスターブック.xlsx を生成中...")
from tool3_master_book import create_workbook as gen3
gen3(os.path.join(OUTPUT_DIR, "master_マスターブック.xlsx"))

# ── ツール④ 口座データバリデーション ─────────────────────────────────────────
print("[4/4] 口座データバリデーション.xlsx を生成中...")
from tool4_account_validation import create_workbook as gen4
gen4(os.path.join(OUTPUT_DIR, "口座データバリデーション.xlsx"))

print("\n" + "=" * 60)
print("  ✅ 全ファイル生成完了！")
print(f"  保存先: {os.path.abspath(OUTPUT_DIR)}")
print("=" * 60)
print("""
次のステップ:
  1. output/ フォルダ内の4つの .xlsx ファイルを確認する
  2. master_マスターブック.xlsx にVBAマクロを追加する
     → Excel で開く → Alt+F11 → ファイル→インポート
     → vba/MasterBook_Macros.bas を選択
  3. ファイルを「マクロ有効ブック (.xlsm)」として保存し直す
  4. 学校のパソコンに転送する（フォルダごとUSB等でコピー）
""")
