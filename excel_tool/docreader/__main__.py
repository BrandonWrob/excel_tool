"""python -m docreader FILE [FILE ...] [--sql out.sqlite] [--query "SELECT ..."]"""
import argparse
import sys

from . import ConversionError, ReadOnlyDB, build_sqlite, convert, to_prompt_block


def main() -> int:
    ap = argparse.ArgumentParser(prog="docreader")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--sql", help="write every extracted table to this SQLite file")
    ap.add_argument("--query", action="append", default=[], help="run a read-only query (needs --sql)")
    ap.add_argument("--max-tokens", type=int, default=20_000)
    args = ap.parse_args()

    results = []
    for f in args.files:
        try:
            r = convert(f, max_inline_tokens=args.max_tokens)
        except ConversionError as e:
            print(f"!! {f}: {e}", file=sys.stderr)
            continue
        results.append(r)
        print(f"===== {r.filename}  kind={r.kind}  strategy={r.strategy}  ~{r.est_tokens:,} tokens  "
              f"{len(r.tables)} SQL table(s)")
        print(to_prompt_block(r))
        print()
    if args.sql and results:
        build_sqlite(results, args.sql)
        db = ReadOnlyDB(args.sql)
        for q in args.query:
            print(f"--- {q}\n{db.query(q)}\n")
        db.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
