# -*- coding: utf-8 -*-
"""矢印・マウスカーソル付き操作アニメーション(GIF)を作るための共通エンジン。
実際のExcel/VBE/Copilotの画面そのものではなく、操作の流れを示す簡略化した模式図。
"""
from PIL import Image, ImageDraw, ImageFont

W, H = 1000, 640
FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"

NAVY = (0x20, 0x38, 0x64)
BLUE = (0x2f, 0x5f, 0xa8)
LBLUE = (0xdc, 0xe8, 0xfa)
GREEN = (0x1e, 0x8e, 0x3e)
LGREEN = (0xe2, 0xf6, 0xe6)
RED = (0xc4, 0x2b, 0x1e)
LRED = (0xfb, 0xe3, 0xe1)
YELLOW = (0xff, 0xf3, 0xb0)
GRAY = (0x55, 0x55, 0x55)
LGRAY = (0xe8, 0xe8, 0xe8)
WHITE = (255, 255, 255)
BORDER = (0xb0, 0xb0, 0xb0)

def font(size):
    return ImageFont.truetype(FONT_PATH, size)

F_TITLE = font(22)
F_H = font(16)
F_BODY = font(15)
F_SMALL = font(13)
F_CALLOUT = font(19)
F_BADGE = font(15)


def new_frame():
    img = Image.new("RGB", (W, H), (0xf2, 0xf2, 0xf2))
    return img, ImageDraw.Draw(img)


def text_w(draw, text, fnt):
    bbox = draw.textbbox((0, 0), text, font=fnt)
    return bbox[2] - bbox[0]


def app_window(draw, x, y, w, h, title, icon_color=BLUE):
    # ウィンドウ全体の枠と上部タイトルバー
    draw.rectangle([x, y, x + w, y + h], fill=WHITE, outline=BORDER, width=2)
    draw.rectangle([x, y, x + w, y + 34], fill=icon_color)
    draw.text((x + 12, y + 7), title, font=F_H, fill=WHITE)
    # ウィンドウ操作ボタン(模式)
    for i, c in enumerate([(0xff, 0x9f, 0x43), (0x4c, 0xd1, 0x64), (0xff, 0x5f, 0x57)]):
        cx = x + w - 18 - i * 22
        draw.ellipse([cx - 6, y + 11, cx + 6, y + 23], fill=c)
    return (x, y + 34, x + w, y + h)


def ribbon(draw, body, tabs, active_idx):
    x0, y0, x1, y1 = body
    draw.rectangle([x0, y0, x1, y0 + 44], fill=(0xf6, 0xf7, 0xf9), outline=BORDER, width=1)
    cx = x0 + 14
    for i, t in enumerate(tabs):
        tw = text_w(draw, t, F_BODY) + 18
        if i == active_idx:
            draw.rectangle([cx, y0 + 4, cx + tw, y0 + 40], fill=WHITE, outline=BLUE, width=2)
            draw.text((cx + 9, y0 + 13), t, font=F_BODY, fill=BLUE)
        else:
            draw.text((cx + 9, y0 + 13), t, font=F_BODY, fill=GRAY)
        cx += tw + 6
    return y0 + 44


def sheet_grid(draw, x0, y0, x1, y1, col_w=90, row_h=30, n_header_rows=1):
    draw.rectangle([x0, y0, x1, y1], fill=WHITE, outline=BORDER, width=1)
    x = x0
    while x < x1:
        draw.line([(x, y0), (x, y1)], fill=LGRAY, width=1)
        x += col_w
    y = y0
    while y < y1:
        draw.line([(x0, y), (x1, y)], fill=LGRAY, width=1)
        y += row_h


def cell_rect(grid_origin, col, row, col_w=90, row_h=30):
    x0, y0 = grid_origin
    return (x0 + col * col_w, y0 + row * row_h, x0 + (col + 1) * col_w, y0 + (row + 1) * row_h)


def fill_cell(draw, rect, color, text=None, text_color=GRAY, fnt=None, border=None):
    draw.rectangle(rect, fill=color, outline=border or LGRAY, width=1)
    if text is not None:
        fnt = fnt or F_SMALL
        draw.text((rect[0] + 6, rect[1] + 7), text, font=fnt, fill=text_color)


def sheet_tabs(draw, x0, y, tabs, active_idx):
    cx = x0
    for i, t in enumerate(tabs):
        tw = text_w(draw, t, F_SMALL) + 20
        if i == active_idx:
            draw.rectangle([cx, y, cx + tw, y + 26], fill=WHITE, outline=BLUE, width=2)
            draw.text((cx + 10, y + 5), t, font=F_SMALL, fill=BLUE)
        else:
            draw.rectangle([cx, y, cx + tw, y + 26], fill=(0xe4, 0xe4, 0xe4), outline=BORDER, width=1)
            draw.text((cx + 10, y + 5), t, font=F_SMALL, fill=GRAY)
        cx += tw + 4


def menu_bar(draw, x0, y0, x1, items, active_idx=None):
    draw.rectangle([x0, y0, x1, y0 + 28], fill=(0xec, 0xec, 0xec), outline=BORDER, width=1)
    cx = x0 + 10
    positions = []
    for i, t in enumerate(items):
        tw = text_w(draw, t, F_SMALL) + 16
        if active_idx == i:
            draw.rectangle([cx, y0 + 2, cx + tw, y0 + 26], fill=BLUE)
            draw.text((cx + 8, y0 + 6), t, font=F_SMALL, fill=WHITE)
        else:
            draw.text((cx + 8, y0 + 6), t, font=F_SMALL, fill=GRAY)
        positions.append((cx, y0 + 26, cx + tw, y0 + 28))
        cx += tw + 4
    return y0 + 28, positions


def dropdown(draw, x, y, items, highlight_idx=None):
    w = max(text_w(draw, t, F_SMALL) for t in items) + 40
    h = 28 * len(items)
    draw.rectangle([x, y, x + w, y + h], fill=WHITE, outline=BORDER, width=2)
    for i, t in enumerate(items):
        iy = y + i * 28
        if i == highlight_idx:
            draw.rectangle([x + 2, iy + 2, x + w - 2, iy + 26], fill=LBLUE)
        draw.text((x + 14, iy + 6), t, font=F_SMALL, fill=GRAY if i != highlight_idx else NAVY)
    return (x, y, x + w, y + h)


def dialog_box(draw, x, y, w, h, title):
    draw.rectangle([x + 6, y + 6, x + w + 6, y + h + 6], fill=(0, 0, 0, 0))  # placeholder shadow skip
    draw.rectangle([x, y, x + w, y + h], fill=WHITE, outline=NAVY, width=2)
    draw.rectangle([x, y, x + w, y + 30], fill=NAVY)
    draw.text((x + 10, y + 6), title, font=F_BODY, fill=WHITE)
    return (x, y + 30, x + w, y + h)


def button(draw, x, y, w, h, text, primary=True, hot=False):
    fill = BLUE if primary else WHITE
    txt_color = WHITE if primary else GRAY
    if hot:
        draw.rectangle([x - 2, y - 2, x + w + 2, y + h + 2], outline=(0xff, 0xc1, 0x07), width=3)
    draw.rectangle([x, y, x + w, y + h], fill=fill, outline=BLUE, width=2)
    tw = text_w(draw, text, F_BODY)
    draw.text((x + (w - tw) / 2, y + (h - 22) / 2), text, font=F_BODY, fill=txt_color)
    return (x, y, x + w, y + h)


def key_badge(draw, cx, cy, text, hot=False):
    w = text_w(draw, text, F_BADGE) + 28
    h = 34
    x, y = cx - w / 2, cy - h / 2
    outline = (0xff, 0xc1, 0x07) if hot else GRAY
    draw.rounded_rectangle([x, y, x + w, y + h], radius=8, fill=(0x33, 0x33, 0x33), outline=outline, width=3)
    tw = text_w(draw, text, F_BADGE)
    draw.text((x + (w - tw) / 2, y + 7), text, font=F_BADGE, fill=WHITE)
    return (x, y, x + w, y + h)


def draw_cursor(draw, pos, clicked=False):
    x, y = pos
    pts = [(x, y), (x, y + 22), (x + 6, y + 17), (x + 10, y + 26),
           (x + 14, y + 24), (x + 10, y + 15), (x + 18, y + 15)]
    draw.polygon(pts, fill=(0, 0, 0), outline=WHITE)
    if clicked:
        for r in (10, 18, 26):
            draw.ellipse([x - r, y - r, x + r, y + r], outline=(0xff, 0xc1, 0x07), width=2)


def draw_arrow(draw, p1, p2, color=RED, width=4):
    import math
    draw.line([p1, p2], fill=color, width=width)
    ang = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
    head = 14
    pA = (p2[0] + head * math.cos(ang + 2.6), p2[1] + head * math.sin(ang + 2.6))
    pB = (p2[0] + head * math.cos(ang - 2.6), p2[1] + head * math.sin(ang - 2.6))
    draw.polygon([p2, pA, pB], fill=color)


def callout(draw, text, hot=True):
    pad = 16
    tw = text_w(draw, text, F_CALLOUT)
    bw = tw + pad * 2
    bx = (W - bw) / 2
    by = H - 56
    draw.rounded_rectangle([bx, by, bx + bw, by + 40], radius=10,
                            fill=(0xff, 0xf3, 0xb0) if hot else LGRAY, outline=NAVY, width=2)
    draw.text((bx + pad, by + 9), text, font=F_CALLOUT, fill=NAVY)


def step_badge(draw, n):
    draw.ellipse([24, 24, 70, 70], fill=NAVY)
    s = str(n)
    tw = text_w(draw, s, F_TITLE)
    draw.text((24 + (46 - tw) / 2, 32), s, font=F_TITLE, fill=WHITE)


def lerp(a, b, t):
    return a + (b - a) * t


def lerp_pt(p1, p2, t):
    return (lerp(p1[0], p2[0], t), lerp(p1[1], p2[1], t))


def save_gif(frames, path, ms_per_frame=45):
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=ms_per_frame,
                    loop=0, optimize=False)
    print(f"saved {path} ({len(frames)} frames)")


def play_shots(shots, out_path, move_frames=16, hold_frames=16, start_pos=(500, 320)):
    """shots: list of dict(draw=fn(draw)->None, to=(x,y), text=str, click=bool, hold=int)"""
    frames = []
    cur_pos = start_pos
    for shot in shots:
        target = shot["to"]
        bg_fn = shot["draw"]
        text = shot.get("text")
        click = shot.get("click", False)
        hold = shot.get("hold", hold_frames)
        mf = shot.get("move", move_frames)
        for i in range(mf):
            t = i / (mf - 1) if mf > 1 else 1
            pos = lerp_pt(cur_pos, target, t)
            img, d = new_frame()
            bg_fn(d)
            draw_cursor(d, pos, clicked=(click and t > 0.96))
            if text:
                callout(d, text)
            frames.append(img)
        for _ in range(hold):
            img, d = new_frame()
            bg_fn(d)
            draw_cursor(d, target, clicked=click)
            if text:
                callout(d, text)
            frames.append(img)
        cur_pos = target
    save_gif(frames, out_path)
