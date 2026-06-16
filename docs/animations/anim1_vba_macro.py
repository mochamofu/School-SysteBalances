# -*- coding: utf-8 -*-
"""①VBAマクロの導入・実行アニメーション(master_マスターブック.xlsm)"""
from _engine import *

BODY = (20, 20, 980, 600)


def excel_base(d, active_tab_idx=6, badge=None, badge_hot=False, result=None):
    body = app_window(d, *BODY[:2], BODY[2]-BODY[0], BODY[3]-BODY[1],
                       "master_マスターブック.xlsm - Excel")
    ry = ribbon(d, body, ["ホーム", "挿入", "ページレイアウト", "数式", "データ", "表示", "開発"], active_tab_idx)
    gx0, gy0 = body[0] + 10, ry + 10
    gx1, gy1 = body[2] - 10, body[3] - 50
    sheet_grid(d, gx0, gy0, gx1, gy1, col_w=120, row_h=28)
    headers = ["生徒番号", "氏名", "クラス", "確定金額"]
    for i, h in enumerate(headers):
        fill_cell(d, cell_rect((gx0, gy0), i, 0, 120, 28), LBLUE, h, NAVY, F_SMALL)
    sheet_tabs(d, gx0, body[3] - 34, ["生徒マスター", "支出記録", "個人別管理表(出力)", "精算書(出力)"], 3)
    if badge:
        key_badge(d, 905, 130, badge, hot=badge_hot)
        d.text((845, 95), "キーボードで", font=F_SMALL, fill=GRAY)
    if result:
        d.rectangle([gx0 + 40, gy0 + 60, gx1 - 40, gy0 + 130], fill=LGREEN, outline=GREEN, width=2)
        d.text((gx0 + 60, gy0 + 85), result, font=F_H, fill=GREEN)
    return body


def vbe_base(d, menu_open=False, dd_highlight=None, module_added=False, file_menu_hot=False):
    body = app_window(d, *BODY[:2], BODY[2]-BODY[0], BODY[3]-BODY[1],
                       "Microsoft Visual Basic for Applications", icon_color=NAVY)
    items = ["ファイル", "編集", "表示", "挿入", "デバッグ", "実行", "ツール", "ウィンドウ", "ヘルプ"]
    mbar_bottom, positions = menu_bar(d, body[0], body[1], body[2], items, active_idx=0 if file_menu_hot else None)
    # 左:プロジェクトツリー
    tree_x0, tree_y0, tree_x1, tree_y1 = body[0] + 10, mbar_bottom + 10, body[0] + 230, body[3] - 10
    d.rectangle([tree_x0, tree_y0, tree_x1, tree_y1], fill=(0xfa, 0xfa, 0xfa), outline=BORDER, width=1)
    d.text((tree_x0 + 8, tree_y0 + 8), "プロジェクト - VBAProject", font=F_SMALL, fill=GRAY)
    d.text((tree_x0 + 16, tree_y0 + 32), "▾ VBAProject (master_マスターブック.xlsm)", font=F_SMALL, fill=GRAY)
    d.text((tree_x0 + 32, tree_y0 + 54), "▸ Microsoft Excel Objects", font=F_SMALL, fill=GRAY)
    if module_added:
        d.rectangle([tree_x0 + 28, tree_y0 + 74, tree_x1 - 8, tree_y0 + 96], fill=LGREEN)
        d.text((tree_x0 + 32, tree_y0 + 78), "▸ 標準モジュール / Module1", font=F_SMALL, fill=GREEN)
    # 右:コードウィンドウ(空)
    code_x0 = tree_x1 + 10
    d.rectangle([code_x0, tree_y0, body[2] - 10, tree_y1], fill=WHITE, outline=BORDER, width=1)
    d.text((code_x0 + 12, tree_y0 + 10), "' コードウィンドウ", font=F_SMALL, fill=(0xaa, 0xaa, 0xaa))
    if module_added:
        sample = ["Sub 精算書_行非表示()", "    ' 金額が0円の行を非表示にする", "    ...", "End Sub"]
        for i, line in enumerate(sample):
            d.text((code_x0 + 12, tree_y0 + 40 + i * 22), line, font=F_SMALL, fill=GRAY)
    if menu_open:
        dd_items = ["ファイルのインポート(I)...", "ファイルのエクスポート(E)...", "閉じてMicrosoft Excelへ戻る(C)"]
        dropdown(d, positions[0][0], positions[0][3], dd_items, highlight_idx=dd_highlight)
    return body, positions


def file_dialog(d, base_draw_fn, hot=False):
    base_draw_fn(d)
    dx, dy, dw, dh = 280, 160, 440, 300
    body = dialog_box(d, dx, dy, dw, dh, "ファイルを開く")
    d.text((dx + 14, dy + 44), "📁 vba/", font=F_BODY, fill=GRAY)
    items = ["MasterBook_Macros.bas"]
    for i, it in enumerate(items):
        iy = dy + 80 + i * 30
        d.rectangle([dx + 10, iy, dx + dw - 10, iy + 28], fill=LBLUE, outline=BLUE, width=2)
        d.text((dx + 18, iy + 5), "📄 " + it, font=F_BODY, fill=NAVY)
    btn = button(d, dx + dw - 110, dy + dh - 46, 90, 34, "開く", hot=hot)
    return btn


def macro_dialog(d, base_draw_fn, selected=0, hot=False):
    base_draw_fn(d)
    dx, dy, dw, dh = 260, 140, 480, 340
    body = dialog_box(d, dx, dy, dw, dh, "マクロ")
    macros = ["精算書_行非表示", "精算書_行表示", "精算書_一括印刷", "精算書_一括PDF出力", "個人別管理表_クラス更新"]
    for i, m in enumerate(macros):
        iy = dy + 46 + i * 32
        if i == selected:
            d.rectangle([dx + 12, iy, dx + dw - 12, iy + 28], fill=LBLUE, outline=BLUE, width=2)
        d.text((dx + 20, iy + 5), m, font=F_BODY, fill=NAVY if i == selected else GRAY)
    btn = button(d, dx + dw - 110, dy + dh - 46, 90, 34, "実行", hot=hot)
    return btn


def build():
    shots = []

    # 1. Excelでファイルを開いた状態 -> Alt+F11
    def s1(d):
        excel_base(d, badge="Alt + F11", badge_hot=True)
    shots.append(dict(draw=s1, to=(905, 130), text="Step1: master_マスターブック.xlsm を開き、Alt+F11 を押す",
                       click=True, hold=20))

    # 2. VBE表示、ファイルメニューへ
    def s2(d):
        vbe_base(d, file_menu_hot=True)
    shots.append(dict(draw=s2, to=(60, 48), text="Step2: VBEditor が開く。「ファイル」メニューをクリック",
                       click=True, hold=14))

    # 3. ドロップダウン表示、ファイルのインポートへ
    def s3(d):
        vbe_base(d, menu_open=True, dd_highlight=0)
    shots.append(dict(draw=s3, to=(80, 110), text="Step3: 「ファイルのインポート」をクリック",
                       click=True, hold=14))

    # 4. ファイル選択ダイアログ、MasterBook_Macros.bas を選んで開く
    def s4(d):
        file_dialog(d, lambda dd: vbe_base(dd, menu_open=False), hot=True)
    shots.append(dict(draw=s4, to=(280 + 440 - 65, 160 + 300 - 29),
                       text="Step4: vba/MasterBook_Macros.bas を選んで「開く」",
                       click=True, hold=16))

    # 5. モジュールが追加された状態、再びファイルメニューへ
    def s5(d):
        vbe_base(d, module_added=True, file_menu_hot=True)
    shots.append(dict(draw=s5, to=(60, 48), text="Step5: Module1 が追加された。「ファイル」メニューをクリック",
                       click=True, hold=14))

    # 6. ドロップダウンで「閉じてExcelへ戻る」
    def s6(d):
        vbe_base(d, menu_open=True, dd_highlight=2, module_added=True)
    shots.append(dict(draw=s6, to=(80, 166), text="Step6: 「閉じてMicrosoft Excelへ戻る」をクリック",
                       click=True, hold=14))

    # 7. Excelに戻り、Alt+F8へ
    def s7(d):
        excel_base(d, badge="Alt + F8", badge_hot=True)
    shots.append(dict(draw=s7, to=(905, 130), text="Step7: Excelに戻った。Alt+F8 でマクロ一覧を開く",
                       click=True, hold=18))

    # 8. マクロ選択ダイアログ、実行ボタン
    def s8(d):
        macro_dialog(d, lambda dd: excel_base(dd), selected=0, hot=True)
    shots.append(dict(draw=s8, to=(260 + 480 - 65, 140 + 340 - 29),
                       text="Step8: マクロを選択して「実行」をクリック",
                       click=True, hold=18))

    # 9. 完了
    def s9(d):
        excel_base(d, result="✓ マクロが正常に実行されました（精算書の0円行が非表示）")
    shots.append(dict(draw=s9, to=(500, 320), text="完了！同じ手順で他のマクロも実行できます", hold=26))

    play_shots(shots, "vba_macro.gif", start_pos=(905, 130))


if __name__ == "__main__":
    build()
