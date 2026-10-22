"""Create a <=100 MB source ZIP, with Dockerfile at its root and no local data."""
import argparse
import ast
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEAM_CODE = "AI26C0024"
MAX_ZIP_BYTES = 100_000_000
REQUIRED = {"Dockerfile", "main.py", "excelqa.py", "requirements.txt", "README.md",
            ".dockerignore", "scripts/__init__.py", "scripts/validate_submission.py",
            "scripts/package_submission.py"}
DENIED = {".venv", ".git", "__pycache__", "debug", "output", "dist", "node_modules"}


def check_secrets(text):
    if re.search(r"\bsk-(?:or-v1-)?[A-Za-z0-9_-]{20,}\b", text):
        raise ValueError("Possible API key found; packaging aborted (value withheld)")
    if re.search(r"-----BEGIN (?:[A-Z0-9 ]+ )?PRIVATE KEY(?: BLOCK)?-----", text):
        raise ValueError("Possible private key found; packaging aborted (value withheld)")
    for name, secret in os.environ.items():
        if re.search(r"KEY|TOKEN|SECRET|PASSWORD|PASSWD|CREDENTIAL", name, re.I) and len(secret) >= 16 and secret in text:
            raise ValueError("Environment credential found; packaging aborted (value withheld)")


def validate_sources(contents):
    missing = REQUIRED - contents.keys()
    if missing:
        raise ValueError("Missing required files: " + ", ".join(sorted(missing)))
    third_party = {re.split(r"[<=>!~\[]", line.strip())[0].replace("-", "_")
                   for line in contents["requirements.txt"].decode().splitlines()
                   if line.strip() and not line.startswith("#")}
    local = {"main", "excelqa", "scripts"}
    for name, raw in contents.items():
        check_secrets(raw.decode("utf-8"))
        if not name.endswith(".py"):
            continue
        tree = ast.parse(raw, filename=name)
        for node in ast.walk(tree):
            modules = [a.name for a in node.names] if isinstance(node, ast.Import) else [node.module] if isinstance(node, ast.ImportFrom) and node.module else []
            for module in modules:
                first = module.split(".")[0]
                if first in local:
                    target = module.replace(".", "/")
                    if target + ".py" not in contents and target + "/__init__.py" not in contents:
                        raise ValueError("Missing local import: " + module)
                elif first not in sys.stdlib_module_names and first not in third_party:
                    raise ValueError("Undeclared dependency: " + first)


def validate_zip(path):
    if Path(path).stat().st_size > MAX_ZIP_BYTES:
        raise ValueError("ZIP exceeds 100 MB")
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)) or "Dockerfile" not in names:
            raise ValueError("ZIP must have unique members and Dockerfile at root")
        if set(names) - REQUIRED:
            raise ValueError("ZIP contains files outside the submission allowlist")
        if archive.testzip() is not None:
            raise ValueError("ZIP CRC check failed")
        contents = {}
        for name in names:
            p = Path(name)
            if (p.is_absolute() or ".." in p.parts or any(part in DENIED for part in p.parts)
                    or p.suffix.lower() in {".xlsx", ".ipynb", ".zip", ".pyc", ".tar"}
                    or p.name.startswith(".env")):
                raise ValueError("Forbidden archive member: " + name)
            contents[name] = archive.read(name)
        validate_sources(contents)
    return {"bytes": Path(path).stat().st_size, "files": len(names), "dockerfile_at_root": True}


def package(team=TEAM_CODE, output_dir=None, root=ROOT):
    if (not re.fullmatch(r'[^<>:"/\\|?*\x00-\x1f]+', team) or team in {".", ".."}
            or team[-1] in ". " or team.split(".")[0].upper() in
            {"CON", "PRN", "AUX", "NUL", *[f"COM{i}" for i in range(1, 10)], *[f"LPT{i}" for i in range(1, 10)]}):
        raise ValueError("Team name must be a safe filename")
    root = Path(root).resolve()
    paths = [root / name for name in REQUIRED]
    contents = {}
    for path in sorted(set(paths)):
        if not path.is_file() or path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError("Missing or unsafe source file: " + path.name)
        contents[path.relative_to(root).as_posix()] = path.read_bytes()
    validate_sources(contents)
    destination = Path(output_dir) if output_dir is not None else root / "dist"
    destination.mkdir(parents=True, exist_ok=True)
    final = destination / (team + ".zip")
    fd, temporary = tempfile.mkstemp(suffix=".zip", dir=destination)
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for name, raw in sorted(contents.items()):
                archive.writestr(name, raw)
        result = validate_zip(temporary)
        os.replace(temporary, final)
        return final, result
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--team", default=TEAM_CODE)
    ap.add_argument("--output-dir")
    args = ap.parse_args()
    zip_path, info = package(args.team, args.output_dir)
    print(f"Created {zip_path} ({info['bytes']} bytes, {info['files']} files)")
