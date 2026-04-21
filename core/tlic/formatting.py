from __future__ import annotations

import math


def trunc_fixed(value: float, decimals: int = 6) -> str:
    factor = 10 ** decimals
    truncated = math.trunc(value * factor) / factor
    return f"{truncated:.{decimals}f}"
