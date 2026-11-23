#!/usr/bin/env python3
"""Cloudflare Pages 한도: 파일 20,000개 미만, 파일 하나 25 MiB 이하."""
import os, sys
root = sys.argv[1]; LIMIT = 25 * 1024 * 1024
files = []; big = []; total = 0
for d, _, fs in os.walk(root):
    for f in fs:
        p = os.path.join(d, f); s = os.path.getsize(p); files.append(p); total += s
        if s > LIMIT: big.append((s, os.path.relpath(p, root)))
print(f'파일 {len(files)}개, 합계 {total/1048576:.1f} MiB, 가장 큰 파일 {max(os.path.getsize(p) for p in files)/1048576:.1f} MiB')
if big: print('25 MiB 초과:', big)
if len(files) >= 20000 or big: sys.exit(1)
