"""AI2026 six-argument entrypoint. No credential is written to environment or disk."""
import argparse
import sys

import excelqa


class Parser(argparse.ArgumentParser):
    def error(self, message):
        # Invalid command lines may contain a credential among unexpected arguments.
        super().error(excelqa.redact(message))


def parser():
    ap = Parser(description="AI2026 Excel QA (CPU)", allow_abbrev=False)
    ap.add_argument("--xlsx-dir", required=True)
    ap.add_argument("--queries", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--openrouter-key", required=True)
    ap.add_argument("--openrouter-model", required=True)
    ap.add_argument("--openrouter-base-url", required=True)
    return ap


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    for i, value in enumerate(argv):
        if value == "--openrouter-key" and i + 1 < len(argv):
            excelqa._SECRETS.add(argv[i + 1])
        elif value.startswith("--openrouter-key="):
            excelqa._SECRETS.add(value.split("=", 1)[1])
    excelqa._SECRETS.discard("")
    args = parser().parse_args(argv)
    try:
        if not args.openrouter_key.strip():
            raise ValueError("--openrouter-key must not be empty")
        if not args.openrouter_model.strip():
            raise ValueError("--openrouter-model must not be empty")
        client = excelqa.OpenRouterClient(args.openrouter_key, args.openrouter_model,
                                          args.openrouter_base_url)
        answers = excelqa.run(args.xlsx_dir, args.queries, args.output, client)
        # Validate the file independently from the writer before declaring success.
        from scripts.validate_submission import validate_submission
        validate_submission(args.queries, args.output)
        unresolved = sum(not answer or answer == "UNVALUE" for answer in answers.values())
        unresolved += len(excelqa.load_queries(args.queries)) - len(answers)
        if unresolved:
            excelqa.log(f"{unresolved} unresolved question(s); UNVALUE/blank answers, see stderr warnings")
        return 0
    except (OSError, ValueError, RuntimeError, ImportError) as e:
        excelqa.log("ERROR: " + excelqa.redact(e))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
