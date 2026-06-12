"""Run: python3 test_squared_divisor_tree.py   or   python3 -m pytest test_squared_divisor_tree.py -v"""

import os
import subprocess
import sys

DIR = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
PROG = os.path.join(DIR, "squared_divisor_tree.py")


def run(inp: str) -> str:
    p = subprocess.run(
        [PY, PROG],
        input=inp,
        text=True,
        capture_output=True,
        cwd=DIR,
    )
    assert p.returncode == 0, p.stderr
    return p.stdout.strip()


def test_sample_star_all_ones():
    inp = """4
1 1 1 1
1 2
1 3
1 4
"""
    assert run(inp) == "16 1 1 1"


def test_sample_five_nodes_explained():
    inp = """5
6 4 3 10 6
1 2
2 4
2 5
1 3
"""
    assert run(inp) == "25 11 2 4 4"


def test_single_node():
    """N=1: no edges; Ans[0] = tau(A[0]) * 1^2."""
    inp = """1
6
"""
    assert run(inp) == "4"


if __name__ == "__main__":
    test_sample_star_all_ones()
    test_sample_five_nodes_explained()
    test_single_node()
    print("OK: squared_divisor_tree tests passed.")
