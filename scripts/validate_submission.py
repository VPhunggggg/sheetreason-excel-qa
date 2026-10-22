"""Independent structural CSV check; does not assert semantic answer accuracy."""
import argparse
import csv
import json
import math
import re
from collections import Counter
from decimal import Decimal, InvalidOperation


def id_sort_key(qid):
    text = str(qid)
    try:
        number = Decimal(text)
        if number.is_finite():
            return (0, number, text)
    except InvalidOperation:
        pass
    parts = tuple((1, int(p)) if p.isdecimal() else (0, p) for p in re.split(r"(\d+)", text))
    return (1, parts, text)


def validate_submission(queries_path, csv_path):
    with open(queries_path, encoding="utf-8-sig") as f:
        queries = json.load(f)
    if not isinstance(queries, list) or not queries:
        raise ValueError("queries must be a nonempty list")
    expected = []
    for q in queries:
        qid = q.get("id") if isinstance(q, dict) else None
        if (isinstance(qid, bool) or not isinstance(qid, (str, int, float))
                or (isinstance(qid, str) and not qid.strip())
                or (isinstance(qid, float) and not math.isfinite(qid))):
            raise ValueError("invalid query id")
        expected.append(str(qid))
    if len(set(expected)) != len(expected):
        raise ValueError("duplicate query IDs in CSV representation")
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, strict=True)
        if next(reader, None) != ["id", "answer"]:
            raise ValueError("CSV header must be exactly id,answer")
        rows = list(reader)
    if any(len(row) != 2 for row in rows):
        raise ValueError("CSV rows must contain exactly two fields")
    actual = [row[0] for row in rows]
    if Counter(actual) != Counter(expected):
        raise ValueError("CSV has missing, extra or duplicate IDs")
    if actual != sorted(expected, key=id_sort_key):
        raise ValueError("CSV IDs must be sorted in increasing order")
    return {"rows": len(rows), "blank_answers": sum(not row[1].strip() for row in rows),
            "input_order_preserved": actual == expected, "ids_sorted": True}


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--queries", required=True)
    ap.add_argument("--submission", required=True)
    args = ap.parse_args()
    print(json.dumps(validate_submission(args.queries, args.submission), ensure_ascii=False))
