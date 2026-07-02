# -*- coding: utf-8 -*-
"""②CoPilotプロンプトガイドの使い方アニメーション"""
from _engine import *

BODY = (20, 20, 980, 600)
PROMPT = "支出記録シートの今月分を集計して..."


def base(d, copilot_open=False, pasted=False, replied=False, badge=None, badge_hot=False):
    body = app_window(d, BODY[0], BODY[1], BODY[2] - BODY[0], BODY[3] - BODY[1],
                       "②CoPilotプロンプトガイド_学次会計用.xlsx - Excel")
    excel_w = (body[2] - 290) if copilot_open else body[2]
    gx0, gy0 = body[0] + 10, body[1] + 10
    gx1, gy1 = excel_w - 10, body[3] - 50
    sheet_grid(d, gx0, gy0, gx1, gy1, col_w=160, row_h=30)
    headers = ["カテゴリ", "プロンプト文", "備考"]
    for i, h in enumerate(headers):
        fill_cell(d, cell_rect((gx0, gy0), i, 0, 160, 30), LBLUE, h, NAVY, F_SMALL)
    fill_cell(d, cell_rect((gx0, gy0), 0, 1, 160, 30), WHITE, "月次集計")
    fill_cell(d, cell_rect((gx0, gy0), 1, 1, 160, 30), YELLOW, PROMPT, NAVY, F_SMALL, border=(0xc9, 0xa8, 0x00))
    sheet_tabs(d, gx0, body[3] - 34, ["プロンプトテンプレート", "Outlookルール設定", "運用カレンダー"], 0)

    if badge:
        key_badge(d, gx1 - 70, gy0 + 70, badge, hot=badge_hot)

    if copilot_open:
        px0 = excel_w + 6
        d.rectangle([px0, body[1], body[2] - 6, body[3] - 6], fill=(0xf5, 0xf7, 0xfb), outline=BORDER, width=2)
        d.rectangle([px0, body[1], body[2] - 6, body[1] + 34], fill=NAVY)
        d.text((px0 + 12, body[1] + 7), "Copilot", font=F_H, fill=WHITE)
        chat_y0 = body[1] + 44
        if replied:
            d.rounded_rectangle([px0 + 10, chat_y0, body[2] - 16, chat_y0 + 80], radius=8,
                                 fill=WHITE, outline=BORDER, width=1)
            d.text((px0 + 20, chat_y0 + 10), "✓ 今月分の支出を集計しました。", font=F_SMALL, fill=GREEN)
            d.text((px0 + 20, chat_y0 + 32), "  内訳は下表のとおりです。", font=F_SMALL, fill=GREEN)
        # 入力欄
        box_y0 = body[3] - 110
        d.rectangle([px0 + 10, box_y0, body[2] - 16, body[3] - 56], fill=WHITE, outline=BLUE, width=2)
        if pasted:
            d.text((px0 + 18, box_y0 + 10), PROMPT, font=F_SMALL, fill=GRAY)
        else:
            d.text((px0 + 18, box_y0 + 10), "メッセージを入力...", font=F_SMALL, fill=(0xaa, 0xaa, 0xaa))
        send_btn = button(d, body[2] - 90, body[3] - 46, 70, 32, "送信", hot=False)
        return body, (px0 + 10, box_y0, body[2] - 16, body[3] - 56), send_btn
    return body, None, None


def build():
    shots = []
    yellow_cell = cell_rect((30, 30), 1, 1, 160, 30)
    ycx, ycy = (yellow_cell[0] + yellow_cell[2]) / 2, (yellow_cell[1] + yellow_cell[3]) / 2

    def s1(d):
        base(d)
    shots.append(dict(draw=s1, to=(ycx, ycy), text="Step1: 黄色いセルのプロンプト文をクリックして選択",
                       click=True, hold=16))

    def s2(d):
        base(d, badge="Ctrl + C", badge_hot=True)
    shots.append(dict(draw=s2, to=(910, 100), text="Step2: コピーする（Ctrl + C）", click=True, hold=16))

    # 画面右側のCopilotパネルのチャット欄座標を計算
    excel_w = BODY[2] - 290
    chat_box_rect = (excel_w + 6 + 10, BODY[3] - 110, BODY[2] - 6 - 16, BODY[3] - 56)
    chat_cx = (chat_box_rect[0] + chat_box_rect[2]) / 2
    chat_cy = (chat_box_rect[1] + chat_box_rect[3]) / 2

    def s3draw(d):
        base(d, copilot_open=True)
    shots.append(dict(draw=s3draw, to=(chat_cx, chat_cy),
                       text="Step3: 画面右側の Copilot チャット欄をクリック", click=True, hold=16))

    def s4(d):
        base(d, copilot_open=True, pasted=True)
    shots.append(dict(draw=s4, to=(chat_cx, chat_cy), text="Step4: 貼り付けて送信（Ctrl + V）",
                       click=False, hold=16))

    send_x, send_y = BODY[2] - 90 + 35, BODY[3] - 46 + 16

    def s5(d):
        base(d, copilot_open=True, pasted=True)
    shots.append(dict(draw=s5, to=(send_x, send_y), text="Step5: 「送信」ボタンをクリック",
                       click=True, hold=16))

    def s6(d):
        base(d, copilot_open=True, pasted=True, replied=True)
    shots.append(dict(draw=s6, to=(send_x, send_y), text="完了！Copilotが回答を生成しました", hold=26))

    play_shots(shots, "copilot_prompt.gif", start_pos=(ycx, ycy))


if __name__ == "__main__":
    build()
