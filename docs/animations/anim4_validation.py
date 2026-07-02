# -*- coding: utf-8 -*-
"""④口座データバリデーション の使い方アニメーション"""
from _engine import *

BODY = (20, 20, 980, 600)
GX0 = BODY[0] + 10
COLW, ROWH = 150, 30


def base(d, active_tab=0, filled=0, fixed=False, point_arrow=False):
    body = app_window(d, BODY[0], BODY[1], BODY[2] - BODY[0], BODY[3] - BODY[1],
                       "④口座データバリデーション.xlsx - Excel")
    sheets = ["口座データ入力", "バリデーション結果"]
    tabs_y = body[1] + 10
    sheet_tabs(d, GX0, tabs_y, sheets, active_tab)
    gy0 = body[1] + 50
    gx1, gy1 = body[2] - 10, body[3] - 40
    sheet_grid(d, GX0, gy0, gx1, gy1, col_w=COLW, row_h=ROWH)

    if active_tab == 0:
        headers = ["申込者氏名", "フリガナ", "口座番号", "支店名"]
        for i, h in enumerate(headers):
            fill_cell(d, cell_rect((GX0, gy0), i, 0, COLW, ROWH), LBLUE, h, NAVY, F_SMALL)
        vals = ["佐藤 花子", "ｓﾅﾄｳ ﾊﾟﾅｺ" if not fixed else "サトウ ハナコ", "1234567", "本店"]
        for i in range(filled):
            color = LGREEN if fixed else WHITE
            fill_cell(d, cell_rect((GX0, gy0), i, 1, COLW, ROWH), color, vals[i],
                      GREEN if fixed else GRAY, F_SMALL)

    elif active_tab == 1:
        headers = ["氏名", "判定", "不一致理由"]
        for i, h in enumerate(headers):
            fill_cell(d, cell_rect((GX0, gy0), i, 0, COLW, ROWH), LBLUE, h, NAVY, F_SMALL)
        fill_cell(d, cell_rect((GX0, gy0), 0, 1, COLW, ROWH), WHITE, "山田 太郎")
        fill_cell(d, cell_rect((GX0, gy0), 1, 1, COLW, ROWH), LGREEN, "OK", GREEN, F_SMALL)
        fill_cell(d, cell_rect((GX0, gy0), 2, 1, COLW, ROWH), WHITE, "")

        ng_color = LGREEN if fixed else LRED
        ng_text = "OK" if fixed else "NG"
        ng_text_color = GREEN if fixed else RED
        fill_cell(d, cell_rect((GX0, gy0), 0, 2, COLW, ROWH), WHITE, "佐藤 花子")
        fill_cell(d, cell_rect((GX0, gy0), 1, 2, COLW, ROWH), ng_color, ng_text, ng_text_color, F_SMALL)
        reason = "" if fixed else "フリガナが半角文字"
        fill_cell(d, cell_rect((GX0, gy0), 2, 2, COLW, ROWH), ng_color if fixed else (0xff, 0xee, 0xee),
                  reason, ng_text_color, F_SMALL)

        if point_arrow:
            reason_rect = cell_rect((GX0, gy0), 2, 2, COLW, ROWH)
            rx, ry = (reason_rect[0] + reason_rect[2]) / 2, reason_rect[1] - 6
            draw_arrow(d, (rx, ry - 60), (rx, ry - 8), color=RED, width=4)
            d.text((rx - 70, ry - 90), "理由がここに出る", font=F_SMALL, fill=RED)
    return body


def build():
    shots = []
    A_cell = cell_rect((GX0, 70), 0, 1, COLW, ROWH)
    acx, acy = (A_cell[0] + A_cell[2]) / 2, (A_cell[1] + A_cell[3]) / 2

    shots.append(dict(draw=lambda d: base(d, 0), to=(acx, acy),
                       text="Step1: 口座データ入力シートの空白行をクリック", click=True, hold=14))
    shots.append(dict(draw=lambda d: base(d, 0, filled=2), to=(acx + COLW, acy), hold=10,
                       text="Step1: 氏名・フリガナなど申込者情報を入力"))
    shots.append(dict(draw=lambda d: base(d, 0, filled=4), to=(acx + COLW * 3, acy), hold=16,
                       text="入力したら自動でバリデーション結果に反映されます"))

    tmp_img, tmp_d = new_frame()
    sheets = ["口座データ入力", "バリデーション結果"]
    cx = GX0
    tab_centers = []
    for t in sheets:
        tw = text_w(tmp_d, t, F_SMALL) + 20
        tab_centers.append((cx + tw / 2, BODY[1] + 23))
        cx += tw + 4

    shots.append(dict(draw=lambda d: base(d, 0, filled=4), to=tab_centers[1],
                       text="Step2: 「バリデーション結果」タブをクリック", click=True, hold=14))

    NG_cell = cell_rect((GX0, 70), 1, 2, COLW, ROWH)
    ngcx, ngcy = (NG_cell[0] + NG_cell[2]) / 2, (NG_cell[1] + NG_cell[3]) / 2
    shots.append(dict(draw=lambda d: base(d, 1), to=(ngcx, ngcy),
                       text="Step3: 赤色のNG行を見つける", hold=18))

    R_cell = cell_rect((GX0, 70), 2, 2, COLW, ROWH)
    rcx, rcy = (R_cell[0] + R_cell[2]) / 2, (R_cell[1] + R_cell[3]) / 2
    shots.append(dict(draw=lambda d: base(d, 1, point_arrow=True), to=(rcx, rcy),
                       text="Step4: 右側の「不一致理由」列を確認する", hold=22))

    shots.append(dict(draw=lambda d: base(d, 1, point_arrow=True), to=tab_centers[0],
                       text="Step5: 「口座データ入力」タブに戻って修正する", click=True, hold=14))

    shots.append(dict(draw=lambda d: base(d, 0, filled=4, fixed=True), to=(acx + COLW, acy),
                       text="フリガナを全角カタカナに修正", hold=16))

    shots.append(dict(draw=lambda d: base(d, 0, filled=4, fixed=True), to=tab_centers[1],
                       text="Step6: 「バリデーション結果」タブで確認", click=True, hold=14))

    shots.append(dict(draw=lambda d: base(d, 1, fixed=True), to=(ngcx, ngcy),
                       text="✓ 修正すると自動的にOKに変わります", hold=28))

    play_shots(shots, "validation.gif", start_pos=(acx, acy))


if __name__ == "__main__":
    build()
