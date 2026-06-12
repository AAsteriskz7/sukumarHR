#!/usr/bin/env python3
# Python 3.8: never use list[int](...) or list[int] = ... — use list() or [] only.
import sys

n = int(input())
A = [int(x) for x in input().split()]
adj = [[] for _ in range(n)]
for _ in range(n - 1):
    line = input().split()
    if len(line) >= 2:
        u = int(line[0]) - 1
        v = int(line[1]) - 1
    else:
        # e.g. "14" typed instead of "1 4" (single-digit labels only)
        s = line[0]
        if len(s) == 2 and s.isdigit():
            a, b = int(s[0]), int(s[1])
            if 1 <= a <= n and 1 <= b <= n and a != b:
                u, v = a - 1, b - 1
            else:
                raise ValueError("invalid edge line (expected u v): %r" % s)
        else:
            raise ValueError("invalid edge line (expected two integers): %r" % " ".join(line))
    adj[u].append(v)
    adj[v].append(u)

M = max(A)
tau = [0] * (M + 1)
for i in range(1, M + 1):
    for j in range(i, M + 1, i):
        tau[j] += 1

parent = [-1] * n
seen = [False] * n
stack = [0]
seen[0] = True
while stack:
    u = stack.pop()
    for v in adj[u]:
        if not seen[v]:
            seen[v] = True
            parent[v] = u
            stack.append(v)

order = []
stack = [(0, -1, False)]
while stack:
    u, p, done = stack.pop()
    if done:
        order.append(u)
        continue
    stack.append((u, p, True))
    for v in reversed(adj[u]):
        if v != p:
            stack.append((v, u, False))

ans = [0] * n
memo = [None] * n
for u in order:
    mp = None
    for v in adj[u]:
        if v == parent[u]:
            continue
        ch = memo[v]
        if mp is None:
            mp = ch
        else:
            if len(ch) > len(mp):
                mp, ch = ch, mp
            for k, c in ch.items():
                mp[k] = mp.get(k, 0) + c
    if mp is None:
        mp = {}
    x = A[u]
    mp[x] = mp.get(x, 0) + 1
    s = 0
    for k, c in mp.items():
        s += tau[k] * c * c
    ans[u] = s
    memo[u] = mp

sys.stdout.write(" ".join(str(x) for x in ans) + "\n")
