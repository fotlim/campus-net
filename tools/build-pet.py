"""把 dsh-pet 的 webm **整段**(241 帧 = 10.04 秒)抽成模型包用的 PNG 帧。

和老的 build-pet.py 的区别:老脚本每段只截 34 帧(1.4 秒),这里抽完整段,
按 12fps 采样(源 24fps → 取一半),高度统一到 TARGET_H,体积/内存都可控。

用法:
    python build-frames.py <webm 目录> <输出目录> [memes 目录]

依赖: pip install av numpy pillow
"""
import os
import sys
import json
import shutil
import numpy as np
import av
from PIL import Image, ImageFilter
from collections import deque
from multiprocessing import Pool

KEY_TH = 22
FEATHER = 1.1
TARGET_H = 200
FPS_OUT = 12
MEME_MAX_W = 300

# (动作名, 显示名, 源文件)  源文件是 241 帧 / 24fps / 10.04 秒
JOBS = [
    ("idle", "待机呼吸", "待机呼吸休闲.webm"),
    ("sleep", "打盹", "原地小憩沉眠.webm"),
    ("wave", "挥手回应", "点击回应-元气挥手.webm"),
    ("walk", "漂浮踏步", "原地漂浮踏步.webm"),
    ("look", "东张西望", "东张西望.webm"),
    ("cat", "撸猫", "撸猫.webm"),
    ("code", "写代码", "写代码.webm"),
    ("hotpot", "涮火锅", "涮火锅.webm"),
    ("rubik", "玩魔方", "原地专心玩魔方.webm"),
    ("stretch", "伸懒腰", "超大伸懒腰.webm"),
    ("snack", "偷吃零食", "偷吃零食被抓住.webm"),
    ("angry", "傲娇生气", "点击回应-傲娇生气.webm"),
    ("tail", "拍拍尾巴", "用鲸鱼尾巴拍打地面.webm"),
]


def mask_alpha(rgb):
    """黑底泛洪:从四边把连通的近黑区域标成透明(角色内部的黑保留)"""
    a = np.asarray(rgb, dtype=np.uint8)
    h, w, _ = a.shape
    cand = a.max(axis=2) <= KEY_TH
    bg = np.zeros((h, w), dtype=bool)
    dq = deque()
    for y in np.flatnonzero(cand[:, 0]):
        if not bg[y, 0]:
            bg[y, 0] = True; dq.append((y, 0))
    for y in np.flatnonzero(cand[:, w - 1]):
        if not bg[y, w - 1]:
            bg[y, w - 1] = True; dq.append((y, w - 1))
    for x in np.flatnonzero(cand[0, :]):
        if not bg[0, x]:
            bg[0, x] = True; dq.append((0, x))
    for x in np.flatnonzero(cand[h - 1, :]):
        if not bg[h - 1, x]:
            bg[h - 1, x] = True; dq.append((h - 1, x))
    pop, app = dq.popleft, dq.append
    while dq:
        y, x = pop()
        if y > 0 and cand[y - 1, x] and not bg[y - 1, x]:
            bg[y - 1, x] = True; app((y - 1, x))
        if y < h - 1 and cand[y + 1, x] and not bg[y + 1, x]:
            bg[y + 1, x] = True; app((y + 1, x))
        if x > 0 and cand[y, x - 1] and not bg[y, x - 1]:
            bg[y, x - 1] = True; app((y, x - 1))
        if x < w - 1 and cand[y, x + 1] and not bg[y, x + 1]:
            bg[y, x + 1] = True; app((y, x + 1))
    return np.where(bg, 0, 255).astype(np.uint8)


def extend_colors(rgb, solid):
    """把不透明像素的颜色往透明一侧扩,避免边缘残留黑边"""
    a = np.asarray(rgb, dtype=np.uint8).copy()
    h, w, _ = a.shape
    filled = solid.copy()
    dq = deque()
    for y, x in zip(*np.where(solid)):
        dq.append((y, x))
    pop, app = dq.popleft, dq.append
    while dq:
        y, x = pop()
        for ny, nx in ((y - 1, x), (y + 1, x), (y, x - 1), (y, x + 1)):
            if 0 <= ny < h and 0 <= nx < w and not filled[ny, nx]:
                filled[ny, nx] = True
                a[ny, nx] = a[y, x]
                app((ny, nx))
    return Image.fromarray(a, "RGB")


def key_black(img):
    rgb = img.convert("RGB")
    alpha = mask_alpha(rgb)
    solid = alpha >= 250
    if solid.any():
        rgb = extend_colors(rgb, solid)
    am = Image.fromarray(alpha, mode="L").filter(ImageFilter.GaussianBlur(FEATHER))
    out = rgb.convert("RGBA")
    out.putalpha(am)
    return out


def build_one(job):
    aid, label, fn, src_dir, out_dir = job
    path = os.path.join(src_dir, fn)
    if not os.path.exists(path):
        return (aid, label, 0, 0, 0, "missing " + fn)
    c = av.open(path)
    frames = [f.to_image().convert("RGB") for f in c.decode(video=0)]
    c.close()
    if not frames:
        return (aid, label, 0, 0, 0, "no frames")
    src_fps = 24.0
    step = max(1, int(round(src_fps / FPS_OUT)))
    picked = frames[::step]

    # 整段内容并集(裁一次,全段共用同一个裁剪框,保证位置不抖)
    x0, y0, x1, y1 = 10 ** 9, 10 ** 9, -1, -1
    for img in picked:
        m = np.asarray(img, dtype=np.uint8).max(axis=2) > KEY_TH
        if not m.any():
            continue
        ys, xs = np.where(m)
        x0 = min(x0, int(xs.min())); x1 = max(x1, int(xs.max()))
        y0 = min(y0, int(ys.min())); y1 = max(y1, int(ys.max()))
    pad = 8
    x0 = max(0, x0 - pad); y0 = max(0, y0 - pad); x1 = min(frames[0].width - 1, x1 + pad); y1 = min(frames[0].height - 1, y1 + pad)
    w, h = x1 - x0 + 1, y1 - y0 + 1
    tw = int(round(w * TARGET_H / h))

    d = os.path.join(out_dir, aid)
    if os.path.isdir(d):
        shutil.rmtree(d)
    os.makedirs(d)
    total = 0
    for i, img in enumerate(picked):
        cropped = img.crop((x0, y0, x1 + 1, y1 + 1)).resize((tw, TARGET_H), Image.LANCZOS)
        fr = key_black(cropped)
        p = os.path.join(d, "%03d.png" % i)
        fr.save(p, optimize=True, compress_level=9)
        total += os.path.getsize(p)
    return (aid, label, len(picked), tw, total, "")


def build_memes(meme_src, out_dir):
    if not meme_src or not os.path.isdir(meme_src):
        return []
    dst = os.path.join(out_dir, "memes")
    os.makedirs(dst, exist_ok=True)
    for old in os.listdir(dst):
        if old.endswith(".png"):
            os.remove(os.path.join(dst, old))
    memes = []
    for i, p in enumerate(sorted(os.listdir(meme_src))):
        if not p.lower().endswith(".png"):
            continue
        im = Image.open(os.path.join(meme_src, p)).convert("RGBA")
        if im.width > MEME_MAX_W:
            im = im.resize((MEME_MAX_W, int(round(im.height * MEME_MAX_W / im.width))), Image.LANCZOS)
        name = "%02d.png" % i
        im.save(os.path.join(dst, name), optimize=True)
        memes.append({"id": "%02d" % i, "label": os.path.splitext(p)[0], "w": im.width, "h": im.height})
    return memes


def main():
    src = sys.argv[1]
    out = sys.argv[2]
    meme_src = sys.argv[3] if len(sys.argv) > 3 else None
    os.makedirs(out, exist_ok=True)

    jobs = [(aid, label, fn, src, out) for aid, label, fn in JOBS]
    results = []
    with Pool(min(8, len(jobs))) as pool:
        for r in pool.imap_unordered(build_one, jobs):
            results.append(r)
            print("%-8s %-6s %3d 帧 %4dpx  %6.1f MB  %s" % (r[0], r[1], r[2], r[3], r[4] / 1048576, r[5]), flush=True)

    by_id = {r[0]: r for r in results}
    anims = []
    for aid, label, fn in JOBS:
        r = by_id.get(aid)
        if not r or r[2] == 0:
            continue
        anims.append({"id": aid, "label": label, "w": r[3], "h": TARGET_H, "frames": r[2], "fps": FPS_OUT})

    # 统一画布:所有动作共用一块画布,每个动作记 ox/oy(水平居中 + 底部对齐)
    max_w = max(a["w"] for a in anims)
    max_h = max(a["h"] for a in anims)
    for a in anims:
        a["ox"] = int(round((max_w - a["w"]) / 2.0))
        a["oy"] = int(max_h - a["h"])

    manifest = {
        "name": "蓝色女仆(dsh-pet)",
        "author": "PC2005-cloud",
        "source": "https://github.com/PC2005-cloud/dsh-pet",
        "license": "MIT",
        "fps": FPS_OUT,
        "canvas": {"w": max_w, "h": max_h},
        "anims": anims,
        "memes": build_memes(meme_src, out),
        "variants": [
            {"id": "default", "name": "原色"},
            {"id": "night", "name": "夜色", "filters": {"bright": 0.75, "saturate": 0.85}},
            {"id": "vivid", "name": "浓彩", "filters": {"saturate": 1.5}},
            {"id": "mono", "name": "墨白", "filters": {"saturate": 0, "bright": 1.05}},
        ],
    }
    with open(os.path.join(out, "pet.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    total = sum(r[4] for r in results)
    print("总帧数 %d , 总大小 %.1f MB" % (sum(r[2] for r in results), total / 1048576))


if __name__ == '__main__':
    main()
