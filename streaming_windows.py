#!/usr/bin/env python3
# streaming_windows_problem.txt — Python 3 only

n, W, L, K = map(int, input().split())
raw = []
for _ in range(n):
    p = input().split()
    raw.append((p[0], int(p[1]), int(p[2]), p[3]))
wm = max(t[1] for t in raw) - L
seen, sums = set(), {}
for key, ts, val, eid in raw:
    if eid in seen:
        continue
    seen.add(eid)
    if ts < wm:
        continue
    k = (key, (ts // W) * W)
    sums[k] = sums.get(k, 0) + val
by_key = {}
for (key, w0), s in sums.items():
    if key not in by_key:
        by_key[key] = []
    by_key[key].append((w0, s))
for key in sorted(by_key):
    rows = sorted(by_key[key], key=lambda r: (-r[1], r[0]))
    for w0, s in rows[:K]:
        print(key, w0, s)
