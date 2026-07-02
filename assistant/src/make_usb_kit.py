# -*- coding: utf-8 -*-
"""デモ用USBキットを組み立てるスクリプト

リポジトリ内の成果物を「USBメモリにそのままコピーできる1フォルダ」に集約し、
ZIP も作る（GoogleドライブへのアップロードはこのZIP1つを上げるだけでよい）。

使い方:
    python assistant/src/make_usb_kit.py
生成物:
    assistant/output/デモキット/                 … USBへコピーする実体
    assistant/output/積立金入力アシスタント_デモキット.zip
"""
import os
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "assistant", "output")
KIT = os.path.join(OUT, "デモキット")

README_ROOT = """================================================================
 積立金会計 入力アシスタント ― デモキット
================================================================

このフォルダをUSBメモリにそのままコピーして持ち運びます。
2台のPCの役割:
  PC1 = 事務担当のPC役（ライブデモの主役）
  PC2 = 保管用PC役（プレゼン表示 + 最後に成果物を受け取る）

■ フォルダの使い方
  01_デモ環境_PC1へコピー/
      → デモ当日、PC1のデスクトップにフォルダごとコピーして使う。
        中の「セットアップ手順.txt」を必ず読むこと（マクロの
        ブロック解除と、初回のみVBAインポートが必要）。
  02_プレゼン資料_PC2で開く/
      → PC2で開く資料一式（製品LP・手順書PDF・操作アニメ）。
        すべてオフラインで動きます。
  03_保管用_デモ中は空のまま/
      → デモの締めで、PC1で完成したマスターと精算書PDFを
        ここにコピーし、PC2に渡して「保管用PCへの格納」を再現する。

■ 鉄則
  ・このキットに本物の生徒データを入れないこと（練習用の架空データのみ）
  ・デモ前日にPC1で通しリハーサルを1回行うこと
詳しい進行は「02_プレゼン資料_PC2で開く/デモ実施手順書.pdf」を参照。
================================================================
"""

README_SETUP = """================================================================
 01_デモ環境 セットアップ手順（デモ当日にPC1で・約5分）
================================================================

1. このフォルダごと、PC1のデスクトップにコピーする
   （USBから直接開かない。動作が遅く、マクロもブロックされるため）

2. コピーした「積立金入力アシスタント.xlsm」を右クリック
   → プロパティ → 下部の「許可する（ブロックの解除）」にチェック → OK
   ※インターネット/USB由来のマクロをExcelが既定でブロックするため必須

3. 積立金入力アシスタント.xlsm を開き、「コンテンツの有効化」を押す

4. 「設定」シートのC3を、このPCでの練習用マスターのフルパスに直す
   例: C:\\Users\\(ユーザー名)\\Desktop\\01_デモ環境_PC1へコピー\\練習用_令和X年度生積立金.xlsx

5. メニューシートのボタンが動くか1つ試す（⑧整合性チェックが安全）

--- 初回のみ（自宅で事前に）------------------------------------
「積立金入力アシスタント.xlsm」がまだ無い場合（.xlsxしか無い場合）:
 a. 積立金入力アシスタント.xlsx を開く → Alt+F11
 b. ファイル → ファイルのインポート で VBAモジュール/A00〜A07 を全部入れる
 c. 「Excelマクロ有効ブック(*.xlsm)」として保存
 d. Alt+F8 → 「初期設定」を実行（メニューにボタンが並ぶ）
================================================================
"""

README_STORE = """このフォルダはデモ中は空のままにしておく。

デモの締めの演出で使う:
 1. PC1で一括入力を終えたマスター（練習用_令和X年度生積立金.xlsx）と
    「精算書PDF」フォルダを、USBのこのフォルダへコピー
 2. USBをPC2に挿し、このフォルダをPC2のデスクトップへコピーして開く
 3. 「いつもの保管用PCへの格納が、そのまま最終工程になります」と締める
"""


def copy(src_rel, dst_dir, dst_name=None, required=True):
    src = os.path.join(REPO, src_rel)
    if not os.path.exists(src):
        if required:
            raise FileNotFoundError(f"必要なファイルがありません: {src_rel}\n"
                                    f"先に build_assistant.py / make_practice_files.py / "
                                    f"docs/build_*_manual.py を実行してください。")
        print(f"  (スキップ: {src_rel})")
        return
    dst = os.path.join(dst_dir, dst_name or os.path.basename(src))
    if os.path.isdir(src):
        shutil.copytree(src, dst, dirs_exist_ok=True)
    else:
        os.makedirs(dst_dir, exist_ok=True)
        shutil.copy2(src, dst)
    print(f"  + {os.path.relpath(dst, KIT)}")


def main():
    if os.path.exists(KIT):
        shutil.rmtree(KIT)
    os.makedirs(KIT)

    # ---- ルートREADME ----
    with open(os.path.join(KIT, "README_はじめにお読みください.txt"), "w", encoding="utf-8-sig") as f:
        f.write(README_ROOT)

    # ---- 01 デモ環境 ----
    d1 = os.path.join(KIT, "01_デモ環境_PC1へコピー")
    os.makedirs(d1)
    with open(os.path.join(d1, "セットアップ手順.txt"), "w", encoding="utf-8-sig") as f:
        f.write(README_SETUP)
    copy("assistant/output/積立金入力アシスタント.xlsx", d1)
    copy("assistant/output/練習用_令和X年度生積立金.xlsx", d1)
    copy("assistant/output/練習用_掲示用名簿.xlsx", d1)
    copy("assistant/output/練習用_空のマスター.xlsx", d1)
    vba_dst = os.path.join(d1, "VBAモジュール")
    os.makedirs(vba_dst)
    vba_src = os.path.join(REPO, "assistant", "vba")
    for name in sorted(os.listdir(vba_src)):
        if name.endswith(".bas"):
            shutil.copy2(os.path.join(vba_src, name), vba_dst)
            print(f"  + 01_デモ環境_PC1へコピー/VBAモジュール/{name}")

    # ---- 02 プレゼン資料 ----
    d2 = os.path.join(KIT, "02_プレゼン資料_PC2で開く")
    os.makedirs(d2)
    copy("site/index.html", d2, "製品LP.html")
    copy("docs/入力アシスタント_手順書.pdf", d2)
    copy("docs/デモ実施手順書.pdf", d2, required=False)
    copy("docs/動作確認_2台PCシミュレーション手順書.pdf", d2, required=False)
    copy("docs/animations", d2, "操作アニメーション")

    # ---- 03 保管用（空） ----
    d3 = os.path.join(KIT, "03_保管用_デモ中は空のまま")
    os.makedirs(d3)
    with open(os.path.join(d3, "このフォルダの使い方.txt"), "w", encoding="utf-8-sig") as f:
        f.write(README_STORE)

    # ---- ZIP化（Googleドライブへはこの1ファイルを上げるだけ） ----
    zip_path = os.path.join(OUT, "積立金入力アシスタント_デモキット.zip")
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(KIT):
            for name in files:
                p = os.path.join(root, name)
                z.write(p, os.path.relpath(p, OUT))
    size_mb = os.path.getsize(zip_path) / 1024 / 1024
    print()
    print(f"完成: {KIT}")
    print(f"完成: {zip_path} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
