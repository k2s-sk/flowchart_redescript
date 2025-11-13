import cv2
import numpy as np
import pytesseract
import xlwings as xw

# ====== 設定 ======
IMG_PATH = "flowchart.png"
SCALE = 0.2        # Excel上のスケール倍率
ARROW_MIN_LEN = 50 # 矢印線の最小長さ
CONNECT_THRESHOLD = 80  # 矢印端点と矩形中心の距離閾値

# ====== ① 前処理 ======
img = cv2.imread(IMG_PATH)
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
blur = cv2.GaussianBlur(gray, (5, 5), 0)
edges = cv2.Canny(blur, 50, 150)

# ====== ② 矢印（線分）検出 ======
lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=80,
                        minLineLength=ARROW_MIN_LEN, maxLineGap=10)
arrows = []
if lines is not None:
    for l in lines:
        x1, y1, x2, y2 = l[0]
        arrows.append((x1, y1, x2, y2))

# ====== ③ 矩形検出 ======
_, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

blocks = []
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    if w < 30 or h < 20:
        continue
    roi = gray[y:y+h, x:x+w]
    text = pytesseract.image_to_string(roi, lang="jpn", config="--psm 6").strip()
    cx, cy = x + w/2, y + h/2
    blocks.append({"x": x, "y": y, "w": w, "h": h, "cx": cx, "cy": cy, "text": text})

# ====== ④ 矢印補正ロジック ======
def nearest_block(point):
    """座標(point)に最も近い矩形中心を返す"""
    px, py = point
    min_d = float('inf')
    nearest = None
    for blk in blocks:
        d = np.hypot(px - blk["cx"], py - blk["cy"])
        if d < min_d:
            min_d = d
            nearest = blk
    return nearest if min_d < CONNECT_THRESHOLD else None

connections = []
for (x1, y1, x2, y2) in arrows:
    start_blk = nearest_block((x1, y1))
    end_blk   = nearest_block((x2, y2))
    if start_blk and end_blk and start_blk != end_blk:
        connections.append((start_blk, end_blk))

# ====== ⑤ Excel再構成 ======
app = xw.App(visible=True)
wb = app.books.add()
sheet = wb.sheets[0]

# 列幅を1に設定
for col in range(1, 101):
    sheet.range((1, col)).column_width = 1

# --- 矩形をExcelに描画 ---
shape_map = {}
for blk in blocks:
    left = blk["x"] * SCALE
    top = blk["y"] * SCALE
    width = blk["w"] * SCALE
    height = blk["h"] * SCALE

    rect = sheet.shapes.add_shape(1, left, top, width, height)
    rect.textframe.characters().text = blk["text"]
    rect.line.fore_color.rgb = (0, 0, 0)
    rect.fill.fore_color.rgb = (255, 255, 255)

    # 後で参照できるようマップ
    shape_map[(blk["cx"], blk["cy"])] = rect

# --- 矢印描画（補正後）---
for (blk1, blk2) in connections:
    sheet.shapes.add_connector(
        Type=1,  # msoConnectorStraight
        BeginX=blk1["cx"] * SCALE,
        BeginY=blk1["cy"] * SCALE,
        EndX=blk2["cx"] * SCALE,
        EndY=blk2["cy"] * SCALE,
    )

print("Excelにフローチャートを再構成しました。")
