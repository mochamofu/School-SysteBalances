# -*- coding: utf-8 -*-
"""2台PCシミュレーションの通しアニメーション

送信側PC（データ受け取り役）と記録側PC（入力・記録役）を左右に並べ、
「銀行の振替結果が届く → USBで渡す → ⑪で照合 → ⑤で一括入力 →
 マスター記録 → 精算書PDF → USBで保管用へ」の業務の流れを1本で見せる。
"""
import math
import _engine as E

# 横長キャンバスに切り替え（エンジンのグローバルを差し替える）
E.W, E.H = 1560, 780

PC1_X, PC2_X = 20, 800          # 各ウィンドウの左端
PC_W, PC_H = 740, 560           # ウィンドウサイズ
PC_Y = 100

MINOU_ROWS = (3,)               # 表示中の4行のうち未納の行（見本）


def role_tags(d, active=None):
    """各PCの上に役割プレートを描く。active='send'/'rec'で光らせる"""
    for key, x, label, sub in [
        ("send", PC1_X, "送信側PC", "データ受け取り役（窓口）"),
        ("rec", PC2_X, "記録側PC", "入力・記録役（会計担当）"),
    ]:
        hot = (active == key)
        d.rounded_rectangle([x, 30, x + 340, 86], radius=10,
                            fill=(0xF5, 0xB7, 0x31) if hot else E.NAVY,
                            outline=(0xC9, 0x8F, 0x00) if hot else E.NAVY, width=3)
        d.text((x + 16, 36), label, font=E.F_TITLE,
               fill=(0x12, 0x20, 0x3D) if hot else E.WHITE)
        d.text((x + 16, 62), sub, font=E.F_SMALL,
               fill=(0x12, 0x20, 0x3D) if hot else (0xCF, 0xDA, 0xEC))


def table(d, x, y, heads, rows, colw, red_rows=(), green_rows=(), yellow=False):
    cx = x
    for w, h in zip(colw, heads):
        E.fill_cell(d, (cx, y, cx + w, y + 26), E.NAVY, h, E.WHITE, E.F_SMALL)
        cx += w
    for ri, row in enumerate(rows):
        cx = x
        ry = y + 26 + ri * 26
        for w, v in zip(colw, row):
            if ri in red_rows:
                fill, tc = E.LRED, E.RED
            elif ri in green_rows:
                fill, tc = E.LGREEN, E.GREEN
            elif yellow:
                fill, tc = E.YELLOW, (0x33, 0x33, 0x33)
            else:
                fill, tc = E.WHITE, (0x33, 0x33, 0x33)
            E.fill_cell(d, (cx, ry, cx + w, ry + 26), fill, v, tc, E.F_SMALL)
            cx += w


BANK_ROWS = [("10000", "80000000", "76,000", "0"),
             ("10001", "80000111", "76,000", "0"),
             ("10005", "80000555", "76,000", "0"),
             ("10006", "80000666", "0", "1"),
             ("10007", "80000777", "76,000", "0")]


def pc1_bank(d, selected=False):
    E.app_window(d, PC1_X, PC_Y, PC_W, PC_H, "練習用_振替結果.xlsx ― 銀行から届いた結果票",
                 icon_color=(0x8A, 0x6D, 0x3B))
    d.text((PC1_X + 30, PC_Y + 60), "自動払込み 振替結果（80件）", font=E.F_BODY, fill=E.NAVY)
    table(d, PC1_X + 30, PC_Y + 95, ["口座記号", "口座番号", "金額", "振替結果"],
          BANK_ROWS, [120, 150, 110, 100], red_rows=(3,))
    d.text((PC1_X + 30, PC_Y + 95 + 26 * 6 + 6), "…（80行つづく／うち2件が資金不足）…",
           font=E.F_SMALL, fill=E.GRAY)
    if selected:
        d.rectangle([PC1_X + 28, PC_Y + 119, PC1_X + 30 + 480, PC_Y + 95 + 26 * 6],
                    outline=(0x1A, 0x73, 0xE8), width=4)
        d.text((PC1_X + 30, PC_Y + 320), "4列を選択してコピー（Ctrl+C）", font=E.F_BODY,
               fill=(0x1A, 0x73, 0xE8))


def pc1_idle(d, note="（記録側の作業を待っている）"):
    E.app_window(d, PC1_X, PC_Y, PC_W, PC_H, "送信側PC ― デスクトップ", icon_color=(0x8A, 0x6D, 0x3B))
    d.text((PC1_X + 30, PC_Y + 70), "[Excel] 練習用_振替結果.xlsx", font=E.F_BODY, fill=(0x33, 0x33, 0x33))
    d.text((PC1_X + 30, PC_Y + 104), "[Excel] 練習用_掲示用名簿.xlsx", font=E.F_BODY, fill=(0x33, 0x33, 0x33))
    d.text((PC1_X + 30, PC_Y + 170), note, font=E.F_SMALL, fill=E.GRAY)


def pc1_store(d):
    E.app_window(d, PC1_X, PC_Y, PC_W, PC_H, "送信側PC（保管用PC役） ― 04_保管用フォルダ",
                 icon_color=(0x1E, 0x6E, 0x3E))
    d.text((PC1_X + 30, PC_Y + 66), "USBから届いた今月の成果物:", font=E.F_BODY, fill=E.NAVY)
    for i, name in enumerate(["[Excel]　 練習用_令和X年度生積立金.xlsx（記録済み）",
                              "[フォルダ] 精算書PDF（78ファイル・組-番号_氏名.pdf）",
                              "[メモ]　　 チェック結果（問題なし・未納2名）"]):
        d.text((PC1_X + 44, PC_Y + 104 + i * 36), name, font=E.F_BODY, fill=(0x33, 0x33, 0x33))
    d.rounded_rectangle([PC1_X + 30, PC_Y + 250, PC1_X + 660, PC_Y + 330], radius=10,
                        fill=E.LGREEN, outline=E.GREEN, width=3)
    d.text((PC1_X + 48, PC_Y + 264), "いつもの「保管用PCへの格納」がそのまま最終工程。", font=E.F_BODY, fill=E.GREEN)
    d.text((PC1_X + 48, PC_Y + 294), "ここまでの所要時間: 約3分（従来: 半日仕事）", font=E.F_BODY, fill=E.GREEN)


def pc2_menu(d, hot=None):
    E.app_window(d, PC2_X, PC_Y, PC_W, PC_H, "積立金入力アシスタント.xlsm ― メニュー")
    items = [("⑪ 振替結果を照合", 0), ("⑤ 収入をマスターへ一括入力", 1),
             ("⑧ マスターの整合性をチェック", 2), ("⑨ 精算書を一括PDF保存", 3)]
    d.text((PC2_X + 30, PC_Y + 60), "◆ きょう使うボタン", font=E.F_BODY, fill=E.GRAY)
    for label, i in items:
        y = PC_Y + 95 + i * 60
        is_hot = (hot == i)
        d.rounded_rectangle([PC2_X + 30, y, PC2_X + 380, y + 46], radius=9,
                            fill=(0xF5, 0xB7, 0x31) if is_hot else E.WHITE,
                            outline=(0xC9, 0x8F, 0x00) if is_hot else (0xD5, 0xDE, 0xED), width=2)
        d.text((PC2_X + 46, y + 12), label, font=E.F_BODY,
               fill=(0x12, 0x20, 0x3D) if is_hot else E.NAVY)


def pc2_paste(d, pasted=True, matched=False, dialog=False):
    E.app_window(d, PC2_X, PC_Y, PC_W, PC_H, "アシスタント ― 振替結果取込シート")
    d.text((PC2_X + 24, PC_Y + 54), "貼り付け（12行目〜）", font=E.F_SMALL, fill=E.NAVY)
    if pasted:
        table(d, PC2_X + 24, PC_Y + 80, ["口座記号", "口座番号", "金額", "結果"],
              [r for r in BANK_ROWS[:4]], [86, 110, 80, 60], yellow=True)
    if matched:
        d.text((PC2_X + 390, PC_Y + 54), "照合結果（自動）", font=E.F_SMALL, fill=E.NAVY)
        table(d, PC2_X + 390, PC_Y + 80, ["精算", "氏名", "判定"],
              [("1", "あおき あおい", "振替済"), ("2", "いしだ いつき", "振替済"),
               ("6", "かねこ かえで", "振替済"), ("7", "きたむら くるみ", "未納")],
              [56, 150, 76], red_rows=(3,), green_rows=(0, 1, 2))
        E.draw_arrow(d, (PC2_X + 350, PC_Y + 150), (PC2_X + 385, PC_Y + 150), color=E.BLUE, width=4)
    if dialog:
        d.rounded_rectangle([PC2_X + 120, PC_Y + 300, PC2_X + 620, PC_Y + 430], radius=10,
                            fill=E.WHITE, outline=E.NAVY, width=3)
        d.rectangle([PC2_X + 120, PC_Y + 300, PC2_X + 620, PC_Y + 334], fill=E.NAVY)
        d.text((PC2_X + 136, PC_Y + 306), "振替結果の照合完了", font=E.F_BODY, fill=E.WHITE)
        d.text((PC2_X + 140, PC_Y + 348), "読取80件 ／ 振替済78名 ／ 未納2名", font=E.F_BODY, fill=E.NAVY)
        d.text((PC2_X + 140, PC_Y + 382), "未納2名は収入入力シートの未納者表へ転記済み", font=E.F_SMALL, fill=E.GRAY)


def pc2_income(d, dialog=False):
    E.app_window(d, PC2_X, PC_Y, PC_W, PC_H, "アシスタント ― 収入入力シート")
    rows = [("収入枠No", "2"), ("件名", "マスター口座振替 10月分"), ("一人あたり金額", "76,000")]
    y = PC_Y + 60
    for label, val in rows:
        d.rectangle([PC2_X + 24, y, PC2_X + 190, y + 32], fill=E.LBLUE, outline=E.BORDER)
        d.text((PC2_X + 34, y + 6), label, font=E.F_SMALL, fill=E.NAVY)
        d.rectangle([PC2_X + 190, y, PC2_X + 430, y + 32], fill=E.YELLOW, outline=(0xE0, 0xC3, 0x41), width=2)
        d.text((PC2_X + 200, y + 6), val, font=E.F_SMALL, fill=(0x33, 0x33, 0x33))
        y += 38
    d.text((PC2_X + 24, y + 8), "―未納者表―（⑪が自動で記入済み）", font=E.F_SMALL, fill=E.NAVY)
    table(d, PC2_X + 24, y + 34, ["精算番号", "氏名（メモ）"],
          [("7", "きたむら くるみ"), ("44", "はやし ひなた")], [90, 180], green_rows=(0, 1))
    if dialog:
        d.rounded_rectangle([PC2_X + 300, PC_Y + 330, PC2_X + 690, PC_Y + 440], radius=10,
                            fill=E.WHITE, outline=E.NAVY, width=3)
        d.rectangle([PC2_X + 300, PC_Y + 330, PC2_X + 690, PC_Y + 362], fill=E.NAVY)
        d.text((PC2_X + 314, PC_Y + 336), "一括入力の完了", font=E.F_BODY, fill=E.WHITE)
        d.text((PC2_X + 318, PC_Y + 374), "入金あり 78名 ／ 未納 2名", font=E.F_BODY, fill=E.NAVY)
        d.text((PC2_X + 318, PC_Y + 404), "バックアップ作成済み・マスター保存済み", font=E.F_SMALL, fill=E.GRAY)


def pc2_master(d):
    E.app_window(d, PC2_X, PC_Y, PC_W, PC_H, "練習用_令和X年度生積立金.xlsx ― データシート",
                 icon_color=(0x1E, 0x6E, 0x3E))
    heads = ["精算", "氏名", "未納", "枠2(振替)", "収入合計"]
    rows = [("1", "あおき あおい", "", "76,000", "152,300"),
            ("2", "いしだ いつき", "", "76,000", "152,300"),
            ("7", "きたむら くるみ", "未納", "", "76,300"),
            ("8", "くどう けんと", "", "76,000", "152,300"),
            ("44", "はやし ひなた", "未納", "", "76,300")]
    table(d, PC2_X + 30, PC_Y + 70, heads, rows, [60, 170, 70, 110, 110], red_rows=(2, 4))
    d.text((PC2_X + 30, PC_Y + 70 + 26 * 6 + 8), "…（80行）… 金額・未納の印・合計まで自動で反映",
           font=E.F_SMALL, fill=E.GRAY)
    d.rounded_rectangle([PC2_X + 30, PC_Y + 330, PC2_X + 640, PC_Y + 386], radius=8,
                        fill=E.LGREEN, outline=E.GREEN, width=2)
    d.text((PC2_X + 46, PC_Y + 344), "マスターは今までのファイルのまま。中身だけが一瞬で完成",
           font=E.F_BODY, fill=E.GREEN)


def pc2_pdf(d):
    E.app_window(d, PC2_X, PC_Y, PC_W, PC_H, "精算書PDF フォルダ ― ⑨の出力",
                 icon_color=(0x1E, 0x6E, 0x3E))
    for i, name in enumerate(["1-1_あおきあおい.pdf", "1-2_いしだいつき.pdf",
                              "1-3_うえのうた.pdf", "1-4_えんどうえま.pdf", "…（78ファイル）"]):
        d.text((PC2_X + 40, PC_Y + 70 + i * 34), ("[PDF] " if i < 4 else "　　  ") + name,
               font=E.F_BODY, fill=(0x33, 0x33, 0x33))
    d.text((PC2_X + 40, PC_Y + 260), "「組-番号_氏名.pdf」で自動命名 → USBで保管用へ",
           font=E.F_SMALL, fill=E.GRAY)


def usb(d, t, reverse=False):
    """USBメモリがPC間を弧を描いて移動する（t: 0→1）"""
    if reverse:
        x0, x1 = PC2_X + 200, PC1_X + 400
    else:
        x0, x1 = PC1_X + 400, PC2_X + 200
    x = x0 + (x1 - x0) * t
    y = PC_Y + 250 - math.sin(t * math.pi) * 190
    d.rounded_rectangle([x - 42, y - 18, x + 42, y + 18], radius=8,
                        fill=(0x33, 0x33, 0x33), outline=(0xF5, 0xB7, 0x31), width=3)
    d.rectangle([x + 42, y - 10, x + 58, y + 10], fill=(0xB0, 0xB0, 0xB0))
    d.text((x - 30, y - 9), "USB", font=E.F_BODY, fill=E.WHITE)


# ==================================================================
# フレーム組み立て
# ==================================================================
frames = []


def scene(bg_fn, n, cursor_from=None, cursor_to=None, click=False,
          caption=None, extra_fn=None):
    """1場面分のフレームを積む。extra_fn(t) で追加アニメ（USBなど）"""
    for i in range(n):
        t = i / max(n - 1, 1)
        img, d = E.new_frame()
        bg_fn(d)
        if extra_fn:
            extra_fn(d, t)
        if cursor_from and cursor_to:
            pos = E.lerp_pt(cursor_from, cursor_to, min(t * 1.4, 1.0))
            E.draw_cursor(d, pos, clicked=(click and t > 0.75))
        elif cursor_to:
            E.draw_cursor(d, cursor_to, clicked=click)
        if caption:
            E.callout(d, caption)
        frames.append(img)


C1 = (PC1_X + 300, PC_Y + 200)   # PC1のテーブルあたり
B11 = (PC2_X + 200, PC_Y + 118)  # ⑪ボタン
B5 = (PC2_X + 200, PC_Y + 178)   # ⑤ボタン
B9 = (PC2_X + 200, PC_Y + 298)   # ⑨ボタン

# 場面1: 銀行から届く
scene(lambda d: (role_tags(d, "send"), pc1_bank(d), pc2_menu(d)), 22,
      cursor_to=C1, caption="場面1｜送信側PCに、銀行の振替結果（80件）が届いた")
# 場面2: コピー
scene(lambda d: (role_tags(d, "send"), pc1_bank(d, selected=True), pc2_menu(d)), 18,
      cursor_to=C1, click=True, caption="場面2｜4列を選択してコピー（またはファイルごとUSBへ）")
# 場面3: USBで渡す
scene(lambda d: (role_tags(d), pc1_bank(d, selected=True), pc2_menu(d)), 26,
      caption="場面3｜USBで記録側PCへ渡す", extra_fn=lambda d, t: usb(d, t))
# 場面4: 貼り付け
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_paste(d, pasted=True)), 18,
      cursor_to=(PC2_X + 200, PC_Y + 150), caption="場面4｜「振替結果取込」シートの12行目に貼り付け")
# 場面5: ⑪をクリック
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_menu(d, hot=0)), 16,
      cursor_from=(PC2_X + 500, PC_Y + 300), cursor_to=B11, click=True,
      caption="場面5｜メニューの「⑪振替結果を照合」を押す")
# 場面6: 照合結果
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_paste(d, pasted=True, matched=True, dialog=True)), 34,
      caption="場面6｜数秒で判定完了 ― 未納2名が自動で見つかる（探す作業ゼロ）")
# 場面7: 収入入力→⑤
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_income(d)), 18,
      cursor_to=(PC2_X + 300, PC_Y + 100), caption="場面7｜未納者表はもう埋まっている。枠No・件名・金額だけ入力")
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_menu(d, hot=1)), 14,
      cursor_from=(PC2_X + 500, PC_Y + 300), cursor_to=B5, click=True,
      caption="場面8｜「⑤収入をマスターへ一括入力」を押す")
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_income(d, dialog=True)), 24,
      caption="場面8｜78名分の入金が一括記録（自動バックアップ済み）")
# 場面9: マスター確認
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_master(d)), 30,
      caption="場面9｜マスターを開くと、金額・未納の印・合計まで反映済み")
# 場面10: 精算書PDF
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_menu(d, hot=3)), 14,
      cursor_from=(PC2_X + 500, PC_Y + 360), cursor_to=B9, click=True,
      caption="場面10｜「⑨精算書を一括PDF保存」")
scene(lambda d: (role_tags(d, "rec"), pc1_idle(d), pc2_pdf(d)), 22,
      caption="場面10｜生徒別の精算書PDFが一括生成される")
# 場面11: USBで保管用へ
scene(lambda d: (role_tags(d), pc1_idle(d, note="（成果物の到着を待つ）"), pc2_pdf(d)), 26,
      caption="場面11｜マスターと精算書PDFをUSBで保管用PCへ",
      extra_fn=lambda d, t: usb(d, t, reverse=True))
# 場面12: 保管完了
scene(lambda d: (role_tags(d, "send"), pc1_store(d), pc2_menu(d)), 40,
      caption="完了｜いつもの保管業務で締め。ここまで約3分（従来は半日仕事）")

E.save_gif(frames, "2台PCシミュレーション.gif", ms_per_frame=90)
