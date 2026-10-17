"""把 dsh-pet 的 webm 转成本项目用的 PNG 帧(24fps)+ 表情包。

用法:
    python build-pet.py <webm 目录> <输出目录> [memes 目录]

依赖: pip install av   (PyAV,自带 FFmpeg)

帧处理流程:
  1. 解码 webm(yuv420p,纯黑底)
  2. 从四边泛洪,把连通的近黑区域抠成透明(角色内部的黑保留)
  3. 边缘羽化前先做"颜色外扩":把透明边缘的颜色换成最近的不透明像素颜色,
     这样合成到浅色背景上不会再有一圈黑边
  4. 按内容并集裁剪 → 缩放到 200px 高 → 24fps 逐帧存 PNG
  5. 写 pet.json

想加动作:在 JOBS 里加一行(动作名, 显示名, 文件名, 起始帧, 帧数),重跑即可。
"""
import av, os, sys, json, glob
from PIL import Image, ImageFilter
import numpy as np
from collections import deque

if len(sys.argv) < 3:
    print(__doc__)
    sys.exit(1)

SRC = sys.argv[1]
OUT = sys.argv[2]
MEME_SRC = sys.argv[3] if len(sys.argv) > 3 else None

FPS_OUT = 24
KEY_TH = 22
FEATHER = 1.1
TARGET_H = 200
MEME_MAX_W = 300

# (动作名, 显示名, 源文件, 起始帧, 取多少帧)  24fps 源,逐帧取
JOBS = [
    ("idle",  "待机呼吸",        "待机呼吸休闲.webm",       4, 60),
    ("sleep", "打盹",            "原地小憩沉眠.webm",       12, 34),
    ("wave",  "挥手回应",        "点击回应-元气挥手.webm",   12, 34),
    ("walk",  "漂浮踏步",        "原地漂浮踏步.webm",       10, 34),
    ("look",  "东张西望",        "东张西望.webm",            12, 34),
    ("cat",   "撸猫",            "撸猫.webm",                10, 34),
    ("code",  "写代码",          "写代码.webm",              10, 34),
    ("hotpot", "涮火锅",         "涮火锅.webm",              10, 34),
    ("rubik", "玩魔方",          "原地专心玩魔方.webm",       10, 34),
    ("stretch", "伸懒腰",        "超大伸懒腰.webm",          10, 34),
    ("snack", "偷吃零食",        "偷吃零食被抓住.webm",      10, 34),
    ("angry", "傲娇生气",        "点击回应-傲娇生气.webm",   10, 34),
    ("tail",  "拍拍尾巴",        "用鲸鱼尾巴拍打地面.webm",  10, 34),
]


def mask_alpha(rgb):
    """黑底泛洪:返回 0/255 的 alpha(L 模式数组)"""
    a = np.array(rgb, dtype=np.uint8)
    h, w, _ = a.shape
    cand = a.max(axis=2) <= KEY_TH
    bg = np.zeros((h, w), dtype=bool)
    dq = deque()
    for x in range(w):
        for y in (0, h - 1):
            if cand[y, x] and not bg[y, x]:
                bg[y, x] = True
                dq.append((y, x))
    for y in range(h):
        for x in (0, w - 1):
            if cand[y, x] and not bg[y, x]:
                bg[y, x] = True
                dq.append((y, x))
    while dq:
        y, x = dq.popleft()
        for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            ny, nx = y + dy, x + dx
            if 0 <= ny < h and 0 <= nx < w and cand[ny, nx] and not bg[ny, nx]:
                bg[ny, nx] = True
                dq.append((ny, nx))
    return np.where(bg, 0, 255).astype(np.uint8), bg


def extend_colors(rgb, solid):
    """把不透明像素的颜色向边缘扩,避免边缘残留黑底(黑边)"""
    a = np.array(rgb, dtype=np.uint8).copy()
    h, w, _ = a.shape
    filled = solid.copy()
    dq = deque()
    ys, xs = np.where(solid)
    for y, x in zip(ys, xs):
        dq.append((y, x))
    while dq:
        y, x = dq.popleft()
        for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            ny, nx = y + dy, x + dx
            if 0 <= ny < h and 0 <= nx < w and not filled[ny, nx]:
                a[ny, nx] = a[y, x]
                filled[ny, nx] = True
                dq.append((ny, nx))
    return Image.fromarray(a, "RGB")


def key_black(img):
    rgb = img.convert("RGB")
    alpha, bg = mask_alpha(rgb)
    solid = alpha >= 250
    if solid.any():
        rgb = extend_colors(rgb, solid)
    am = Image.fromarray(alpha, mode="L").filter(ImageFilter.GaussianBlur(FEATHER))
    out = rgb.convert("RGBA")
    out.putalpha(am)
    return out


def load_frames(path, start, count):
    c = av.open(path)
    frames, idx = [], 0
    want = set(range(start, start + count))
    last = max(want)
    for f in c.decode(video=0):
        if idx in want:
            frames.append(f.to_image().convert("RGB"))
        if idx >= last:
            break
        idx += 1
    c.close()
    return frames


def content_bbox(frames):
    x0, y0, x1, y1 = 10 ** 9, 10 ** 9, -1, -1
    for img in frames:
        m = np.array(img, dtype=np.uint8).max(axis=2) > KEY_TH
        if not m.any():
            continue
        ys, xs = np.where(m)
        x0 = min(x0, int(xs.min())); x1 = max(x1, int(xs.max()))
        y0 = min(y0, int(ys.min())); y1 = max(y1, int(ys.max()))
    return x0, y0, x1, y1


def build_memes(out_dir):
    if not MEME_SRC or not os.path.isdir(MEME_SRC):
        return []
    dst = os.path.join(out_dir, "memes")
    os.makedirs(dst, exist_ok=True)
    for old in glob.glob(os.path.join(dst, "*.png")):
        os.remove(old)
    memes = []
    for i, p in enumerate(sorted(glob.glob(os.path.join(MEME_SRC, "*.png")))):
        try:
            im = Image.open(p).convert("RGBA")
            if im.width > MEME_MAX_W:
                im = im.resize((MEME_MAX_W, int(round(im.height * MEME_MAX_W / im.width))), Image.LANCZOS)
            name = "%02d.png" % i
            im.save(os.path.join(dst, name), optimize=True)
            memes.append({"id": "%02d" % i, "label": os.path.splitext(os.path.basename(p))[0],
                          "w": im.width, "h": im.height})
        except Exception as e:
            print("meme skip:", p, e)
    return memes


def main():
    os.makedirs(OUT, exist_ok=True)
    manifest = {"source": "https://github.com/PC2005-cloud/dsh-pet",
                "license": "MIT", "fps": FPS_OUT, "anims": [], "memes": []}
    total = 0
    for aid, label, fn, start, count in JOBS:
        path = os.path.join(SRC, fn)
        if not os.path.exists(path):
            print("skip (missing):", fn)
            continue
        frames = load_frames(path, start, count)
        if not frames:
            print("skip (no frames):", fn)
            continue
        x0, y0, x1, y1 = content_bbox(frames)
        pad = 8
        x0 = max(0, x0 - pad); y0 = max(0, y0 - pad); x1 += pad; y1 += pad
        w, h = x1 - x0 + 1, y1 - y0 + 1
        tw = int(round(w * TARGET_H / h))
        d = os.path.join(OUT, aid)
        os.makedirs(d, exist_ok=True)
        for old in os.listdir(d):
            if old.endswith(".png"):
                os.remove(os.path.join(d, old))
        for i, img in enumerate(frames):
            cropped = img.crop((x0, y0, x1 + 1, y1 + 1)).resize((tw, TARGET_H), Image.LANCZOS)
            key_black(cropped).save(os.path.join(d, "%03d.png" % i), optimize=True)
        manifest["anims"].append({"id": aid, "label": label, "w": tw, "h": TARGET_H, "frames": len(frames)})
        total += len(frames)
        print("%-8s %-6s %2d 帧 %dx%d" % (aid, label, len(frames), tw, TARGET_H))

    manifest["memes"] = build_memes(OUT)
    with open(os.path.join(OUT, "pet.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print("总帧数:", total, " 表情包:", len(manifest["memes"]))


main()
