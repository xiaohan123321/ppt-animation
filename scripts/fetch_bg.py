# -*- coding: utf-8 -*-
"""联网搜索真实风光照片，建背景图库。

来源（免钥，按顺序兜底）：
  1. Openverse API  —— CC 授权过滤，可商用
  2. Wikimedia Commons —— 搜图 API
  3. Lorem Picsum —— 最后兜底（Unsplash 实拍池，随机风景）

缓存目录默认 D:\\ppt_bg_cache（按用户约定：下载与缓存一律放 D 盘）。
用法：
  python fetch_bg.py --query "mountain landscape clouds" --count 12
  python fetch_bg.py --list
"""
import argparse
import glob
import io
import json
import os
import urllib.parse
import urllib.request

from PIL import Image

CACHE = r"D:\ppt_bg_cache"
UA = {"User-Agent": "Mozilla/5.0 (ppt-animation-bg-fetcher/1.0)"}
MIN_W = 1280          # 低于这个宽度的图不要


def _get(url, timeout=25):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def search_openverse(query, n):
    url = ("https://api.openverse.org/v1/images/?format=json"
           "&license_type=commercial&size=large"
           f"&q={urllib.parse.quote(query)}&page_size={n}")
    try:
        data = json.loads(_get(url))
        return [it["url"] for it in data.get("results", []) if it.get("url")]
    except Exception as e:
        print(f"  openverse 不可用: {e}")
        return []


def search_wikimedia(query, n):
    url = ("https://commons.wikimedia.org/w/api.php?action=query&generator=search"
           f"&gsrsearch={urllib.parse.quote(query)}&gsrlimit={n + 4}&gsrnamespace=6"
           "&prop=imageinfo&iiprop=url|size&iiurlwidth=1920&format=json")
    try:
        data = json.loads(_get(url))
        pages = (data.get("query") or {}).get("pages") or {}
        out = []
        for p in pages.values():
            ii = (p.get("imageinfo") or [{}])[0]
            u = ii.get("thumburl") or ii.get("url")
            if u and (ii.get("width") or 1920) >= MIN_W:
                out.append(u)
        return out
    except Exception as e:
        print(f"  wikimedia 不可用: {e}")
        return []


def picsum(n):
    return [f"https://picsum.photos/seed/bg{i}/1920/1080" for i in range(n * 3)]


def search_bing(count):
    """Bing 每日壁纸：实拍高清、国内可达。多市场多日期拉一批，按标题挑风光。"""
    urls, seen = [], set()
    for mkt in ("zh-CN", "en-US", "ja-JP", "de-DE", "fr-FR"):
        for idx in (0, 4):
            try:
                u = (f"https://cn.bing.com/HPImageArchive.aspx?format=js&idx={idx}"
                     f"&n=8&mkt={mkt}")
                data = json.loads(_get(u))
            except Exception:
                continue
            for it in data.get("images", []):
                full = "https://cn.bing.com" + it.get("url", "")
                title = (it.get("title") or "") + (it.get("copyright") or "")
                if full and full not in seen:
                    seen.add(full)
                    urls.append((full, title))
    keys = ("山", "峰", "雪", "湖", "云", "森林", "峡", "岭", "崖", "谷", "瀑",
            "岩", "原", "海", "漠", "河", "岛", "mountain", "lake", "valley",
            "forest", "canyon", "glacier", "peak", "desert", "coast", "river",
            "island", "alps")
    picked = [u for u, t in urls if any(k in t for k in keys)] or [u for u, _ in urls]
    return picked[:count * 2]


def fetch(query, count):
    os.makedirs(CACHE, exist_ok=True)
    have = len(glob.glob(os.path.join(CACHE, "photo_*.jpg")))
    need = count - have
    print(f"图库 {CACHE}：已有 {have} 张，目标 {count} 张")
    if need <= 0:
        return
    urls = (search_bing(need) or search_openverse(query, need * 3)
            or search_wikimedia(query, need * 3) or picsum(need))
    got = have
    for i, u in enumerate(urls):
        if got >= count:
            break
        try:
            raw = _get(u, timeout=40)
            im = Image.open(io.BytesIO(raw)).convert("RGB")
            if im.size[0] < MIN_W:
                continue
            got += 1
            fp = os.path.join(CACHE, f"photo_{got:02d}.jpg")
            im.save(fp, quality=90)
            print(f"  + photo_{got:02d}.jpg  {im.size[0]}x{im.size[1]}  <- {u[:70]}")
        except Exception as e:
            print(f"  跳过（{u[:60]}）: {e}")
    print(f"完成：{got} 张实拍在库")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", default="mountain landscape mist clouds")
    ap.add_argument("--count", type=int, default=12)
    ap.add_argument("--list", action="store_true", help="只列出图库")
    a = ap.parse_args()
    if a.list:
        for f in sorted(glob.glob(os.path.join(CACHE, "*.jpg"))):
            print(" ", f)
    else:
        fetch(a.query, a.count)
