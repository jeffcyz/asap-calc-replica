#!/usr/bin/env python3
"""decoder.py: encoded/ 里的 PNG -> 校验 -> 还原 ZIP -> 解压到 recovered/"""
import argparse
import hashlib
import struct
import sys
import zipfile
import zlib
from collections import defaultdict
from pathlib import Path

from PIL import Image

from encoder import HEADER_FMT, HEADER_SIZE, MAGIC


def read_png(path: Path):
    img = Image.open(path)
    if img.mode != "RGB":
        sys.exit(f"错误：{path.name} 模式为 {img.mode}，应为 RGB（图片可能被转换/重新压缩过）")
    raw = img.tobytes()
    magic, ver, sha, total, idx, clen, plen, crc = struct.unpack(HEADER_FMT, raw[:HEADER_SIZE])
    if magic != MAGIC:
        return None
    chunk = raw[HEADER_SIZE:HEADER_SIZE + clen]
    if len(chunk) != clen or zlib.crc32(chunk) != crc:
        sys.exit(f"错误：{path.name} 数据 CRC 校验失败（像素被修改，通常是图片被压缩或缩放）")
    return sha, total, idx, plen, chunk


def safe_extract(zf: zipfile.ZipFile, out: Path):
    root = out.resolve()
    for m in zf.infolist():
        target = (out / m.filename).resolve()
        if root != target and root not in target.parents:
            sys.exit(f"错误：ZIP 含非法路径 {m.filename}")
    zf.extractall(out)


def decode(encoded_dir: Path, output_dir: Path):
    groups = defaultdict(dict)   # sha -> {idx: chunk}
    meta = {}                    # sha -> (total, payload_len)
    for p in sorted(encoded_dir.glob("*.png")):
        r = read_png(p)
        if r is None:
            print(f"跳过 {p.name}（不是本工具生成的图片）")
            continue
        sha, total, idx, plen, chunk = r
        groups[sha][idx] = chunk
        meta[sha] = (total, plen)
    if not groups:
        sys.exit("错误：没有找到可解码的 PNG")

    for sha, chunks in groups.items():
        total, plen = meta[sha]
        missing = [i + 1 for i in range(total) if i not in chunks]
        if missing:
            sys.exit(f"错误：缺少第 {missing} 张图片（共 {total} 张）")
        payload = b"".join(chunks[i] for i in range(total))
        if len(payload) != plen or hashlib.sha256(payload).digest() != sha:
            sys.exit("错误：整包 SHA-256 校验失败")
        print(f"[{sha.hex()[:8]}] {total} 张图片，SHA-256 校验通过 ✔")

        dest = output_dir if len(groups) == 1 else output_dir / sha.hex()[:8]
        dest.mkdir(parents=True, exist_ok=True)
        zip_path = dest / "_recovered.zip"
        zip_path.write_bytes(payload)
        with zipfile.ZipFile(zip_path) as zf:
            safe_extract(zf, dest)
        zip_path.unlink()
        print(f"已恢复到 {dest}/")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="高密度 PNG -> 原始文件")
    ap.add_argument("-i", "--input", default="encoded", type=Path)
    ap.add_argument("-o", "--output", default="recovered", type=Path)
    a = ap.parse_args()
    decode(a.input, a.output)
