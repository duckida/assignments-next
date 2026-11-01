#!/usr/bin/env python3
"""Vendor LINE Seed JP (Google Fonts, unicode-range split) into public/fonts/.
Generates public/fonts/line-seed-jp.css referencing local files."""
import concurrent.futures as cf
import os
import re
import urllib.request

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "public", "fonts", "line-seed-jp")
OUT_CSS = os.path.join(os.path.dirname(__file__), "..", "public", "fonts", "line-seed-jp.css")
GF_URL = "https://fonts.googleapis.com/css2?family=LINE+Seed+JP:wght@400;700&display=swap"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"}


def fetch(url, dest=None):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=30) as r:
        data = r.read()
    if dest:
        with open(dest, "wb") as f:
            f.write(data)
    return data


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    css = fetch(GF_URL).decode("utf-8")
    blocks = re.findall(r"@font-face \{[^}]+\}", css)
    jobs = []
    for b in blocks:
        weight = re.search(r"font-weight: (\d+);", b).group(1)
        url = re.search(r"url\((https://[^)]+)\)", b).group(1)
        idx = re.search(r"\.(\d+)\.woff2", url)
        base = idx.group(1) if idx else re.search(r"/([^/]+)\.woff2", url).group(1)[-12:]
        fname = f"lsj-{weight}-{base}.woff2"
        jobs.append((b, weight, url, fname))

    def dl(job):
        _, _, url, fname = job
        path = os.path.join(OUT_DIR, fname)
        if not os.path.exists(path):
            fetch(url, path)
        return fname

    with cf.ThreadPoolExecutor(max_workers=12) as ex:
        list(ex.map(dl, jobs))

    out = []
    for b, weight, url, fname in jobs:
        out.append(b.replace(url, f"./line-seed-jp/{fname}"))
    with open(OUT_CSS, "w") as f:
        f.write("\n".join(out) + "\n")

    total = sum(os.path.getsize(os.path.join(OUT_DIR, f)) for f in os.listdir(OUT_DIR))
    print(f"{len(jobs)} subsets, {total/1024:.0f} KiB total -> {OUT_DIR}")


if __name__ == "__main__":
    main()
