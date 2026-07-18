# -*- coding: utf-8 -*-
"""学校へ郵送する導入用USB（2本構成）の中身を生成する。

全国展開（オンライン導入）用。デモキット（make_usb_kit.py＝自分が持ち歩く営業・
リハーサル用）とは別物で、こちらは学校に「送りっぱなし」にする配布物。

  USB①_記録側PC用 … マスターのある保管用PCに差す。アシスタント一式＋練習データ＋資料
  USB②_送信側PC用 … 銀行データ・名簿を受け取るPCに差す。受け渡しフォルダ雛形＋練習データ

生成先: assistant/output/導入用USBキット/ と ZIP 2本
実行前提: build_assistant.py / make_practice_files.py / 各PDFが生成済みであること。

★出荷前に必ずやること（Windows作業・1回だけ）
  1. 積立金入力アシスタント.xlsx に A00〜A11.bas をインポートし、
     「積立金入力アシスタント.xlsm」として保存（＝出荷用マスター）
  2. USB①の「積立金入力アシスタント」フォルダ内の .xlsx を、その .xlsm に差し替える
  3. 差し替え後、練習用マスターで ①④⑤⑪ を1回ずつ動かして出荷検査
  ※これで学校側のVBAインポート作業がゼロになる（Zoomで一番つまずく工程の排除）
"""
import os
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "assistant", "output")
KIT = os.path.join(OUT, "導入用USBキット")


def write_text(path, text):
    with open(path, "w", encoding="utf-8-sig", newline="\r\n") as f:  # メモ帳で開ける形式
        f.write(text)


def copy(rel, dst, newname=None, required=True):
    src = os.path.join(REPO, rel)
    if not os.path.exists(src):
        if required:
            raise FileNotFoundError(src)
        print(f"  - skip (not found): {rel}")
        return
    name = newname or os.path.basename(src)
    if os.path.isdir(src):
        shutil.copytree(src, os.path.join(dst, name))
    else:
        shutil.copy2(src, os.path.join(dst, name))
    print(f"  + {os.path.relpath(os.path.join(dst, name), OUT)}")


README_USB1 = """＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
　USB①　記録側PC用（積立金マスターがある保管用パソコンへ）
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

最初に「00_スタートガイド_まずこの紙.pdf」を開いてください。
図だけで手順が分かります（同封の紙と同じものです）。

　1.「1_積立金入力アシスタント」フォルダを丸ごと
　　 デスクトップにコピーする（USBから直接開かない）
　2. 手順書の「手順2」のとおり、マクロを許可する設定をする
　3. Zoomレクチャーの日を待つ（うまくいかなくても当日一緒にやります）

・生徒の実データはこのUSBに入っていません（練習用の架空データのみ）。
・いまお使いの積立金マスターには何も手を加えません。
"""

README_USB2 = """＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
　USB②　送信側PC用（銀行データ・名簿を受け取るパソコンへ）
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

やることは1つだけです。

　1.「1_積立金データ受け渡し」フォルダを丸ごと
　　 デスクトップにコピーする

以後、銀行から届いた振替結果や掲示用名簿はこのフォルダに
入れて整理し、USBで記録側（保管用）パソコンへ運びます。
くわしい流れはZoomレクチャーでご説明します。
"""

README_HANDOFF = """このフォルダの使い方（デスクトップにコピーしてから）

01_銀行から届いたデータ … ゆうちょの振替結果などを、届いた日ごとに入れる
02_名簿など学校の書類　 … 掲示用名簿など、マスターに反映したいものを入れる
03_記録側PCへ運ぶ箱　　 … 記録側へ渡すファイルをここにまとめてUSBへコピー
99_渡し終わったもの　　 … 渡し終わったファイルの置き場（消さずに残す）
"""


def main():
    if os.path.exists(KIT):
        shutil.rmtree(KIT)
    os.makedirs(KIT)

    # ============ USB① 記録側PC用 ============
    u1 = os.path.join(KIT, "USB1_記録側PC用")
    os.makedirs(u1)
    copy("docs/04_オンライン導入/スタートガイド_図解版.pdf", u1, "00_スタートガイド_まずこの紙.pdf")
    write_text(os.path.join(u1, "0_はじめにお読みください.txt"), README_USB1)

    a = os.path.join(u1, "1_積立金入力アシスタント")
    os.makedirs(a)
    copy("assistant/output/積立金入力アシスタント.xlsx", a)
    write_text(os.path.join(a, "★出荷前にxlsmへ差し替え.txt"),
               "【事業者側の作業メモ・学校に渡る前に削除】\r\n"
               "この .xlsx に VBAモジュール（A00〜A11）をインポートして\r\n"
               "「積立金入力アシスタント.xlsm」として保存したものに差し替え、\r\n"
               "このメモと VBAモジュール予備フォルダを消してから出荷する。\r\n")
    vba_dst = os.path.join(a, "VBAモジュール予備")
    os.makedirs(vba_dst)
    vba_src = os.path.join(REPO, "assistant", "vba")
    for name in sorted(os.listdir(vba_src)):
        if name.endswith(".bas"):
            shutil.copy2(os.path.join(vba_src, name), vba_dst)
    print("  + USB1_記録側PC用/1_積立金入力アシスタント/VBAモジュール予備/(12本)")

    p = os.path.join(u1, "2_練習用データ")
    os.makedirs(p)
    for f in ["練習用_令和X年度生積立金.xlsx", "練習用_口座マスター.xlsx",
              "練習用_振替結果.xlsx", "練習用_掲示用名簿.xlsx", "練習用_空のマスター.xlsx"]:
        copy(f"assistant/output/{f}", p)
    write_text(os.path.join(p, "練習データについて.txt"),
               "架空の80名（4クラス×20名）の練習用データです。実在の生徒・口座は\r\n"
               "一切含まれません。Zoomレクチャーではこのデータで一巡の練習をします。\r\n"
               "・振替結果の照合(⑪)の想定結果： 読取80件／振替済78件／未納2名（精算番号7・44）\r\n")

    d = os.path.join(u1, "3_資料")
    os.makedirs(d)
    copy("docs/04_オンライン導入/オンライン導入手順書.pdf", d)
    copy("docs/01_学校向けマニュアル/使い方ガイド_図解版.pdf", d, required=False)
    copy("docs/01_学校向けマニュアル/入力アシスタント_手順書.pdf", d)
    copy("docs/02_営業・商談資料/価格とサービスのご説明.pdf", d, required=False)

    # ============ USB② 送信側PC用 ============
    u2 = os.path.join(KIT, "USB2_送信側PC用")
    os.makedirs(u2)
    copy("docs/04_オンライン導入/スタートガイド_図解版.pdf", u2, "00_スタートガイド_まずこの紙.pdf")
    write_text(os.path.join(u2, "0_はじめにお読みください.txt"), README_USB2)

    h = os.path.join(u2, "1_積立金データ受け渡し")
    os.makedirs(h)
    write_text(os.path.join(h, "このフォルダの使い方.txt"), README_HANDOFF)
    for sub in ["01_銀行から届いたデータ", "02_名簿など学校の書類",
                "03_記録側PCへ運ぶ箱", "99_渡し終わったもの"]:
        os.makedirs(os.path.join(h, sub))
        # 空フォルダはZIPで消えるので中に説明を1枚置く
        write_text(os.path.join(h, sub, "（このフォルダに入れます）.txt"), sub + "\r\n")

    p2 = os.path.join(u2, "2_練習用データ")
    os.makedirs(p2)
    copy("assistant/output/練習用_振替結果.xlsx", p2)
    copy("assistant/output/練習用_掲示用名簿.xlsx", p2)

    d2 = os.path.join(u2, "3_資料")
    os.makedirs(d2)
    copy("docs/04_オンライン導入/オンライン導入手順書.pdf", d2)

    # ============ 手元用: ラベル印刷シート ============
    copy("docs/04_オンライン導入/USBラベル印刷シート.pdf", KIT, required=False)

    # ============ 出荷前チェックリスト（USBには入れない・手元用） ============
    write_text(os.path.join(KIT, "出荷前チェックリスト_手元用.txt"),
               "【出荷前チェックリスト】（このファイルはUSBにコピーしない）\r\n"
               "\r\n"
               "□ USB1の .xlsx をマクロ組込済みの .xlsm に差し替えた\r\n"
               "□ 「★出荷前にxlsmへ差し替え.txt」と「VBAモジュール予備」を削除した\r\n"
               "□ 差し替えた .xlsm を練習用マスターで動作検査した（①④⑤⑪を各1回）\r\n"
               "□ USBはFAT32またはexFATでフォーマットした新品を使用\r\n"
               "□ ラベル「①記録側PC用」「②送信側PC用」を貼った\r\n"
               "□ 実在の生徒データ・口座データが入っていないことを最終確認した\r\n"
               "□ スタートガイド_図解版 と オンライン導入手順書 を印刷して同封した\r\n□ USBラベル印刷シートのラベルを切ってUSB2本と封筒に貼った\r\n□ スタートガイドの「サポート連絡先」記入欄に電話番号等を手書きした\r\n□ 発送メール（Zoom日程・事前チェックリスト・手順書PDF添付）を送った\r\n")

    # ============ ZIP化（USBごと） ============
    for name in ["USB1_記録側PC用", "USB2_送信側PC用"]:
        zip_path = os.path.join(OUT, f"導入用{name}.zip")
        if os.path.exists(zip_path):
            os.remove(zip_path)
        base = os.path.join(KIT, name)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(base):
                for fn in files:
                    fp = os.path.join(root, fn)
                    z.write(fp, os.path.relpath(fp, KIT))
        print(f"\n完成: {zip_path} ({os.path.getsize(zip_path)/1024/1024:.1f} MB)")


if __name__ == "__main__":
    main()
