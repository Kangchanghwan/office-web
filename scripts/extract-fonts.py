#!/usr/bin/env python3
"""extract-fonts.py — 从 ONLYOFFICE 字体池（vendor/fonts/<id>）提取标准 ttf/otf

背景：Document Server 的 allfontsgen 把标准字体（ttf/otf/ttc）复制为 <id> 文件后，
用 correct16 密钥对**前 32 字节**做 XOR 混淆（ApplicationFontsWorker.cpp SaveWebFonts）。
所以每个 <id> 文件 = 标准字体 + 头 32 字节 XOR。

本脚本：
1. 对每个字体文件做 XOR 解密
2. 识别 ttf / otf(CFF) / ttc(TrueType Collection)
3. 读字体名（name 表 nameID=1），建立「字体名 → 输出文件」映射
4. 输出：
   - <out>/decoded/<id>.<ext>          解密后的标准字体文件
   - <out>/fonts-index.json            「字体名(大小写不敏感) → [文件列表]」索引

用法:
    python3 extract-fonts.py [fonts_dir] [out_dir]
    默认 fonts_dir=vendor/fonts, out_dir=build/fonts
    加 --deploy <frontend_dir> 时把解密字体 + 索引复制到
    <frontend_dir>/vendor/decoded-fonts/（office.html 按需 fetch 的源）。

产物用途：PDF 转换前，前端解析渲染 bin 的字体名 → 查 fonts-index.json →
按需 fetch 对应字体 → 写 wasm FS → wasm 端不再回退，PDF 用页面同款字体（所见即所得）。
"""

import argparse
import json
import os
import shutil
import struct
import sys

# 密钥（ApplicationFontsWorker.cpp SaveWebFonts 的 correct16）
CORRECT16 = bytes([0xA0, 0x66, 0xD6, 0x20, 0x14, 0x96, 0x47, 0xfa,
                   0x95, 0x69, 0xB8, 0x50, 0xB0, 0x41, 0x49, 0x48])


def decode_font(data):
    """对前 32 字节做 XOR，返回还原后的标准字体字节。"""
    d = bytearray(data)
    n = min(32, len(d))
    for k in range(n):
        d[k] ^= CORRECT16[k & 0x0F]
    return bytes(d)


def font_magic(data):
    """返回 (magic_str, kind)：kind in {ttf, otf, ttc, unknown}。"""
    if data[:4] in (b'\x00\x01\x00\x00', b'true'):
        return 'ttf', 'ttf'
    if data[:4] == b'OTTO':
        return 'otf', 'otf'
    if data[:4] == b'ttcf':
        return 'ttc', 'ttc'
    return data[:4].hex(), 'unknown'


def ttc_faces(data):
    """ttc 里的 font offset 列表。"""
    num = struct.unpack('>I', data[8:12])[0]
    offs = []
    for i in range(num):
        off = struct.unpack('>I', data[12 + i * 4:16 + i * 4])[0]
        offs.append(off)
    return offs


def read_family_name(data, offset=0):
    """从 ttf/otf 的 name 表读 family（nameID=1）。ttc 传 face 偏移。"""
    try:
        from fontTools.ttLib import TTFont
        import io
        tf = TTFont(io.BytesIO(data[offset:]), lazy=True)
        return tf['name'].getDebugName(1)
    except Exception:
        return None


def read_full_name(data, offset=0):
    """从 name 表读完整字体名（nameID=4，含字重，如 'DejaVu Sans Bold'），
    供前端在同名多字重文件间优先挑 Regular。"""
    try:
        from fontTools.ttLib import TTFont
        import io
        tf = TTFont(io.BytesIO(data[offset:]), lazy=True)
        return tf['name'].getDebugName(4)
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser(description='提取 ONLYOFFICE 字体池为标准 ttf/otf')
    ap.add_argument('fonts_dir', nargs='?',
                    default='OnlyofficePersonal/9.4.0.1-official/vendor/fonts')
    ap.add_argument('out_dir', nargs='?', default='build/fonts')
    ap.add_argument('--deploy', help='把解密字体+索引复制到 <dir>/vendor/decoded-fonts/')
    args = ap.parse_args()

    fonts_dir = os.path.abspath(args.fonts_dir)
    out_dir = os.path.abspath(args.out_dir)
    os.makedirs(out_dir, exist_ok=True)
    dec_dir = os.path.join(out_dir, 'decoded')
    os.makedirs(dec_dir, exist_ok=True)

    index = {}   # 字体名(小写) -> [{file, family, kind, size}]
    n_ok = n_skip = 0

    for fname in sorted(os.listdir(fonts_dir)):
        if fname.endswith('.br'):
            continue
        src = os.path.join(fonts_dir, fname)
        try:
            raw = open(src, 'rb').read()
        except OSError:
            continue
        data = decode_font(raw)
        _, kind = font_magic(data)

        # 确定输出扩展名
        ext = {'ttf': 'ttf', 'otf': 'otf', 'ttc': 'ttc'}.get(kind)
        if not ext:
            n_skip += 1
            continue

        # 读字体名（ttc 取所有 face）
        families = []
        if kind == 'ttc':
            for off in ttc_faces(data)[:8]:
                fam = read_family_name(data, off)
                if fam:
                    families.append(fam)
        else:
            fam = read_family_name(data)
            if fam:
                families.append(fam)

        # 读完整名（含字重），供前端优先挑 Regular
        full = read_full_name(data)

        # 写解密文件
        out_path = os.path.join(dec_dir, fname + '.' + ext)
        open(out_path, 'wb').write(data)

        entry = {'file': fname, 'kind': kind, 'size': len(data),
                 'families': families, 'full': full}
        for fam in families:
            key = fam.lower()
            index.setdefault(key, []).append(entry)
        n_ok += 1

    with open(os.path.join(out_dir, 'fonts-index.json'), 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=1)

    print(f'成功解密 {n_ok} 个字体，跳过 {n_skip} 个非字体')
    print(f'索引 {len(index)} 个字体名 → {os.path.join(out_dir, "fonts-index.json")}')
    print(f'解密文件 → {dec_dir}/')

    # 汇总覆盖的字体名（抽样）
    total_family = len(index)
    print(f'\n字体名总数: {total_family}')

    # 部署到前端（office.html 按需 fetch 的源）
    if args.deploy:
        deploy_dir = os.path.join(os.path.abspath(args.deploy),
                                  'vendor', 'decoded-fonts')
        os.makedirs(deploy_dir, exist_ok=True)
        for fname in os.listdir(dec_dir):
            shutil.copy2(os.path.join(dec_dir, fname), deploy_dir)
        shutil.copy2(os.path.join(out_dir, 'fonts-index.json'), deploy_dir)
        print(f'已部署到 {deploy_dir}/（{len(os.listdir(dec_dir))} 字体 + 索引）')
        # 预压缩（brotli），减小按需下载体积
        try:
            import subprocess
            # opensource 的 precompress.py 在仓库根目录（scripts/ 同级的上一位）
            subprocess.run([sys.executable,
                            os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                         '..', 'precompress.py'),
                            '--root', deploy_dir, '--min-size', '0', '--quiet'],
                           check=False)
        except Exception:
            pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
