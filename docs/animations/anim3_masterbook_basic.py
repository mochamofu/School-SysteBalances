# -*- coding: utf-8 -*-
"""③master_マスターブック 基本操作（生徒登録・支出記録入力）アニメーション"""
from _engine import *

BODY = (20, 20, 980, 600)
GX0, GY0 = BODY[0] + 10, BODY[1] + 60
COLW, ROWH = 130, 30


def base(d, active_tab=0, student_row=None, expense_row=None, flash_cell=None):
    body = app_window(d, BODY[0], BODY[1], BODY[2] - BODY[0], BODY[3] - BODY[1],
                       "master_マスターブック.xlsm - Excel")
    sheets = ["生徒マスター", "支出記録", "個人別管理表(出力)", "精算書(出力)"]
    sheet_tabs(d, GX0, body[1] + 10, sheets, active_tab)
    gy0 = body[1] + 50
    gx1, gy1 = body[2] - 10, body[3] - 40
    sheet_grid(d, GX0, gy0, gx1, gy1, col_w=COLW, row_h=ROWH)

    if active_tab == 0:
        headers = ["生徒番号", "氏名", "クラス", "保護者連絡先"]
        sample = [["04001", "山田 太郎", "1年A組", "090-xxxx-0001"]]
        for i, h in enumerate(headers):
            fill_cell(d, cell_rect((GX0, gy0), i, 0, COLW, ROWH), LBLUE, h, NAVY, F_SMALL)
        for i, v in enumerate(sample[0]):
            fill_cell(d, cell_rect((GX0, gy0), i, 1, COLW, ROWH), WHITE, v, GRAY, F_SMALL)
        if student_row is not None:
            vals = ["04002", "佐藤 花子", "1年A組", ""][:student_row]
            for i in range(student_row):
                fill_cell(d, cell_rect((GX0, gy0), i, 2, COLW, ROWH), LGREEN, vals[i], GREEN, F_SMALL)

    elif active_tab == 1:
        headers = ["支出項目番号", "生徒番号", "金額", "備考"]
        for i, h in enumerate(headers):
            fill_cell(d, cell_rect((GX0, gy0), i, 0, COLW, ROWH), LBLUE, h, NAVY, F_SMALL)
        if expense_row is not None:
            vals = ["3", "04002", "1,500", ""][:expense_row]
            for i in range(expense_row):
                fill_cell(d, cell_rect((GX0, gy0), i, 1, COLW, ROWH), LGREEN, vals[i], GREEN, F_SMALL)

    elif active_tab == 2:
        headers = ["生徒番号", "氏名", "①教材費", "②校外学習費", "確定金額合計"]
        for i, h in enumerate(headers):
            fill_cell(d, cell_rect((GX0, gy0), i, 0, COLW, ROWH), LBLUE, h, NAVY, F_SMALL)
        rows = [["04001", "山田 太郎", "3,000", "2,000", "5,000"],
                ["04002", "佐藤 花子", "3,000", "1,500", "4,500"]]
        for r, row in enumerate(rows):
            for i, v in enumerate(row):
                col = LGREEN if (flash_cell == r and i == len(row) - 1) else WHITE
                fill_cell(d, cell_rect((GX0, gy0), i, r + 1, COLW, ROWH), col, v,
                          GREEN if col == LGREEN else GRAY, F_SMALL)
    return body


def build():
    shots = []
    A_cell = cell_rect((GX0, GY0), 0, 2, COLW, ROWH)
    acx, acy = (A_cell[0] + A_cell[2]) / 2, (A_cell[1] + A_cell[3]) / 2
    D_cell = cell_rect((GX0, GY0), 3, 2, COLW, ROWH)
    dcx, dcy = (D_cell[0] + D_cell[2]) / 2, (D_cell[1] + D_cell[3]) / 2

    shots.append(dict(draw=lambda d: base(d, 0), to=(acx, acy),
                       text="Step1: 生徒マスターシートの空白行を選ぶ", click=True, hold=14))
    shots.append(dict(draw=lambda d: base(d, 0, student_row=1), to=(acx + COLW, acy), hold=8,
                       text="Step2: 生徒番号・氏名・クラスを入力"))
    shots.append(dict(draw=lambda d: base(d, 0, student_row=2), to=(acx + COLW * 2, acy), hold=8,
                       text="Step2: 生徒番号・氏名・クラスを入力"))
    shots.append(dict(draw=lambda d: base(d, 0, student_row=3), to=(dcx, dcy), hold=16,
                       text="入力するだけで登録完了です"))

    # シート切り替え（支出記録タブ）のタブ中心座標を計算
    tabs_y = BODY[1] + 10
    tmp_img, tmp_d = new_frame()
    sheets = ["生徒マスター", "支出記録", "個人別管理表(出力)", "精算書(出力)"]
    cx = GX0
    tab_centers = []
    for t in sheets:
        tw = text_w(tmp_d, t, F_SMALL) + 20
        tab_centers.append((cx + tw / 2, tabs_y + 13))
        cx += tw + 4

    shots.append(dict(draw=lambda d: base(d, 0, student_row=3), to=tab_centers[1],
                       text="Step3: 「支出記録」シートのタブをクリック", click=True, hold=14))

    E_cell = cell_rect((GX0, GY0), 0, 2, COLW, ROWH)
    ecx, ecy = (E_cell[0] + E_cell[2]) / 2, (E_cell[1] + E_cell[3]) / 2
    shots.append(dict(draw=lambda d: base(d, 1), to=(ecx, ecy),
                       text="Step4: 空白行に支出項目番号を入力", click=True, hold=12))
    shots.append(dict(draw=lambda d: base(d, 1, expense_row=2), to=(ecx + COLW, ecy), hold=8,
                       text="Step4: 続けて生徒番号を入力"))
    shots.append(dict(draw=lambda d: base(d, 1, expense_row=3), to=(ecx + COLW * 2, ecy), hold=14,
                       text="Step4: 最後に金額を入力すれば記録完了"))

    shots.append(dict(draw=lambda d: base(d, 1, expense_row=3), to=tab_centers[2],
                       text="Step5: 「個人別管理表(出力)」タブをクリック", click=True, hold=14))

    F_cell = cell_rect((GX0, GY0), 4, 2, COLW, ROWH)
    fcx, fcy = (F_cell[0] + F_cell[2]) / 2, (F_cell[1] + F_cell[3]) / 2
    shots.append(dict(draw=lambda d: base(d, 2, flash_cell=1), to=(fcx, fcy),
                       text="✓ 入力すると確定金額が自動的に反映されます", hold=28))

    play_shots(shots, "masterbook_basic.gif", start_pos=(acx, acy))


if __name__ == "__main__":
    build()
