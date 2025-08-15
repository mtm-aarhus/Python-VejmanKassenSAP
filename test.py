#!/usr/bin/env python3
import json
import sys
from typing import Any, List, Tuple

# Toggle this to treat lists as ordered (True) or compare them item-by-item where possible (False).
STRICT_LIST_ORDER = True

def load_json(path: str) -> Any:
    with open(path, "r", encoding="ANSI") as f:
        return json.load(f)

def type_name(x: Any) -> str:
    return type(x).__name__

def is_scalar(x: Any) -> bool:
    return not isinstance(x, (dict, list))

def join_path(parts: List[str]) -> str:
    if not parts:
        return "$"  # root
    out = "$"
    for p in parts:
        if p.isidentifier():
            out += f".{p}"
        else:
            # Use bracket notation for non-identifier keys or indexes
            if p.isdigit():
                out += f"[{p}]"
            else:
                out += f"[{json.dumps(p)}]"
    return out

def compare(a: Any, b: Any, path: List[str], diffs: List[Tuple[str, str]]):
    """
    Populate diffs with tuples of (kind, message) where kind in:
      - "ADDED", "REMOVED", "CHANGED", "TYPE", "LENGTH", "INFO"
    """
    # Types differ
    if type(a) != type(b):
        diffs.append(("TYPE",
                      f"{join_path(path)} type differs: {type_name(a)} vs {type_name(b)}"))
        return

    # Scalars: compare directly
    if is_scalar(a):
        if a != b:
            diffs.append(("CHANGED",
                          f"{join_path(path)} value differs: {repr(a)} → {repr(b)}"))
        return

    # Dicts: compare keys and recurse
    if isinstance(a, dict):
        a_keys = set(a.keys())
        b_keys = set(b.keys())
        for k in sorted(a_keys - b_keys):
            diffs.append(("REMOVED", f"{join_path(path + [str(k)])} present only in Json1"))
        for k in sorted(b_keys - a_keys):
            diffs.append(("ADDED", f"{join_path(path + [str(k)])} present only in Json2"))
        for k in sorted(a_keys & b_keys):
            compare(a[k], b[k], path + [str(k)], diffs)
        return

    # Lists: compare length and items
    if isinstance(a, list):
        if STRICT_LIST_ORDER:
            if len(a) != len(b):
                diffs.append(("LENGTH",
                              f"{join_path(path)} list length differs: {len(a)} vs {len(b)}"))
            # Compare up to the min length
            for i in range(min(len(a), len(b))):
                compare(a[i], b[i], path + [str(i)], diffs)
            # Report extra tail items explicitly
            for i in range(len(a) - 1, len(b) - 1, -1):
                if i >= len(b):
                    diffs.append(("REMOVED", f"{join_path(path + [str(i)])} present only in Json1"))
            for i in range(len(b) - 1, len(a) - 1, -1):
                if i >= len(a):
                    diffs.append(("ADDED", f"{join_path(path + [str(i)])} present only in Json2"))
        else:
            # Heuristic: try to match dict/list elements by value; fall back to index
            if len(a) != len(b):
                diffs.append(("LENGTH",
                              f"{join_path(path)} list length differs: {len(a)} vs {len(b)}"))

            # Attempt multiset-like comparison for scalars
            if all(is_scalar(x) for x in a + b):
                a_sorted = sorted(a)
                b_sorted = sorted(b)
                if a_sorted != b_sorted:
                    diffs.append(("CHANGED",
                                  f"{join_path(path)} list items differ (order-insensitive)"))
            else:
                # Compare element-by-element up to min length
                for i in range(min(len(a), len(b))):
                    compare(a[i], b[i], path + [str(i)], diffs)
                # Report extras
                for i in range(min(len(a), len(b)), len(a)):
                    diffs.append(("REMOVED", f"{join_path(path + [str(i)])} present only in Json1"))
                for i in range(min(len(a), len(b)), len(b)):
                    diffs.append(("ADDED", f"{join_path(path + [str(i)])} present only in Json2"))
        return

def print_diffs(diffs: List[Tuple[str, str]]):
    if not diffs:
        print("✔ No differences found.")
        return
    # Minimal color if terminal supports ANSI (safe to print regardless)
    COLORS = {
        "ADDED": "\033[32m",    # green
        "REMOVED": "\033[31m",  # red
        "CHANGED": "\033[33m",  # yellow
        "TYPE": "\033[35m",     # magenta
        "LENGTH": "\033[36m",   # cyan
        "INFO": "\033[34m",     # blue
    }
    RESET = "\033[0m"
    for kind, msg in diffs:
        color = COLORS.get(kind, "")
        print(f"{color}{kind:<7}{RESET} {msg}")

def main():
    # Filenames exactly as given, assumed to be in the project folder
    file1 = "Json1.txt"
    file2 = "Json2.txt"

    try:
        a = load_json(file1)
        b = load_json(file2)
    except FileNotFoundError as e:
        print(f"File not found: {e.filename}", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"JSON parse error in {e.doc[:40]}... at pos {e.pos}: {e}", file=sys.stderr)
        sys.exit(2)

    diffs: List[Tuple[str, str]] = []
    compare(a, b, [], diffs)
    print_diffs(diffs)

if __name__ == "__main__":
    main()
