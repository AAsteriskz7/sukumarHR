"""Run: python3 -m pytest test_streaming_windows.py -v   OR   python3 test_streaming_windows.py"""

import os
import subprocess
import sys

DIR = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
PROG = os.path.join(DIR, "streaming_windows.py")


def run_prog(inp: str) -> str:
    p = subprocess.run(
        [PY, PROG],
        input=inp,
        text=True,
        capture_output=True,
        cwd=DIR,
    )
    assert p.returncode == 0, p.stderr
    return p.stdout.strip()


def test_sample_matches_expected():
    with open(os.path.join(DIR, "streaming_windows_sample.in")) as f:
        inp = f.read()
    with open(os.path.join(DIR, "streaming_windows_sample.out")) as f:
        want = f.read().strip()
    got = run_prog(inp)
    assert got == want, "got: %r\nwant: %r" % (got, want)


def test_tiny_single_event():
    # N=1 W=10 L=0 K=1 → one window [0,10), Tmax=5 WM=5, ts>=5 -> ts=5 ok
    inp = "1 10 0 1\nk 5 42 e0\n"
    got = run_prog(inp)
    assert got == "k 0 42"


if __name__ == "__main__":
    test_sample_matches_expected()
    test_tiny_single_event()
    print("OK: sample + tiny test passed.")
