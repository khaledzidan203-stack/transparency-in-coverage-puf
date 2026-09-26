"""Read-only release checks; writes only ignored evidence and Python bytecode.

Uses the standard library. Does not run analytical pipelines or PBIR builders.
Run from any working directory. Exit 1 means a release gate failed.
"""
from pathlib import Path
from urllib.parse import unquote
import hashlib
import json
import py_compile
import re
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[2]
SOURCE_HASH = "27379dc76590027b0e6d21d736331f87376f0ac87810d68e9f7cca555eef4b1f"


def main() -> int:
    evidence = ROOT / ".local-review"
    evidence.mkdir(exist_ok=True)
    failures = []
    def check(ok, message):
        if not ok:
            failures.append(message)

    result = subprocess.run(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
        cwd=ROOT, capture_output=True, check=True,
    )
    paths = sorted({ROOT / p.decode("utf-8") for p in result.stdout.split(b"\0") if p})
    check(all(p.is_file() for p in paths), "Missing publication candidate")
    paths = [p for p in paths if p.is_file()]
    for p in paths:
        check(p.stat().st_size < 100 * 1024**2, f"Oversize file: {p.relative_to(ROOT)}")
        check(not any(x in p.parts for x in (".pbi", "__pycache__", ".local-release", ".local-review")), "Local artifact in publication candidates")
        check(p.suffix != ".zip", "Backup ZIP in publication candidates")
        check(p.suffix != ".xlsx" or p.parent == ROOT / "data/raw", "Review workbook in publication candidates")

    py_files = [p for p in paths if p.suffix == ".py"]
    for p in py_files:
        try:
            py_compile.compile(str(p), doraise=True)
        except py_compile.PyCompileError:
            failures.append(f"Python syntax: {p.relative_to(ROOT)}")

    report = ROOT / "powerbi/TransparencyInCoverage.Report"
    model = ROOT / "powerbi/TransparencyInCoverage.SemanticModel"
    parsed = {}
    for p in paths:
        if p.suffix in (".json", ".pbip", ".pbir", ".pbism", ".platform"):
            try:
                parsed[p] = json.loads(p.read_text(encoding="utf-8-sig"))
            except (ValueError, UnicodeError):
                failures.append(f"JSON parse: {p.relative_to(ROOT)}")
    pbip = ROOT / "powerbi/TransparencyInCoverage.pbip"
    check(pbip in parsed, "PBIP entry missing")
    if pbip in parsed:
        for artifact in parsed[pbip]["artifacts"]:
            check((pbip.parent / artifact["report"]["path"]).is_dir(), "Broken PBIP report path")
    definition = report / "definition.pbir"
    if definition in parsed:
        check((report / parsed[definition]["datasetReference"]["byPath"]["path"]).resolve() == model.resolve(), "Broken report-model reference")
    else:
        failures.append("Report definition missing")
    check((model / "definition/relationships.tmdl").is_file(), "Relationships missing")
    check((model / "definition/tables/_Measures.tmdl").is_file(), "Measures missing")
    page_root = report / "definition/pages"
    page_files = sorted(page_root.glob("*/page.json"))
    page_ids = {p.parent.name for p in page_files}
    order = parsed.get(page_root / "pages.json", {})
    check(len(page_ids) == 11, "Expected 11 pages")
    check(set(order.get("pageOrder", [])) == page_ids and len(order.get("pageOrder", [])) == 11, "Page order mismatch")
    check(order.get("activePageName") in page_ids, "Active page missing")
    visual_files = sorted(page_root.glob("*/visuals/*/visual.json"))
    nav_count = 0
    def walk(value):
        nonlocal nav_count
        if isinstance(value, dict):
            for key, item in value.items():
                if key in ("destination", "navigationSection") and isinstance(item, dict):
                    target = item.get("expr", {}).get("Literal", {}).get("Value")
                    if target:
                        nav_count += 1
                        check(target.strip("'") in page_ids, "Navigation destination missing")
                walk(item)
        elif isinstance(value, list):
            for item in value:
                walk(item)
    for p in visual_files:
        obj = parsed.get(p, {})
        check(obj.get("name") == p.parent.name, f"Visual identity mismatch: {p.parent.name}")
        walk(obj)
    check(nav_count >= 20, "Expected INDEX and return navigation targets")
    images = sorted((ROOT / "screenshots").glob("*.png"))
    check(len(images) == 11, "Expected 11 screenshots")
    for p in images:
        check(p.read_bytes().startswith(b"\x89PNG\r\n\x1a\n"), "Invalid PNG signature")

    link_count = 0
    for p in [x for x in paths if x.suffix == ".md"]:
        text = p.read_text(encoding="utf-8-sig")
        targets = re.findall(r"\]\(([^)]+)\)", text) + re.findall(r'(?:src|href)="([^"]+)"', text)
        for target in targets:
            if re.match(r"^[a-zA-Z]+:", target) or target.startswith("#"):
                continue
            target = unquote(target.split("#")[0].strip("<>"))
            if target:
                link_count += 1
                check((p.parent / target).exists(), f"Broken link in {p.relative_to(ROOT)}: {target}")

    # Findings identify files only; never print potential secret values.
    patterns = {
        "private key": r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----",
        "credential": r"(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}|AKIA[A-Z0-9]{16})",
        "assigned secret": r"(?i)(?:password|passwd|api_key|apikey|client_secret)\s*[:=]\s*[\"'][^\"'\s]{8,}[\"']",
        "authorization value": r"(?i)authorization\s*[:=]\s*[\"']?(?:Bearer|Basic)\s+[A-Za-z0-9+/=_-]{12,}",
        "personal path": r"(?i)[A-Z]:\\Users\\[A-Za-z0-9._-]+",
    }
    findings = []
    email_files = set()
    text_extensions = {".py", ".md", ".txt", ".json", ".tmdl", ".pbip", ".pbir", ".pbism", ".platform", ".csv", ".ps1", ".sql"}
    def scan(label, text):
        for kind, pattern in patterns.items():
            if re.search(pattern, text):
                findings.append({"file": label, "kind": kind})
        if re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text):
            email_files.add(label)
    for p in paths:
        if p.suffix in text_extensions or p.name.startswith(".env"):
            scan(p.relative_to(ROOT).as_posix(), p.read_text(encoding="utf-8-sig"))
        elif p.suffix == ".xlsx":
            with zipfile.ZipFile(p) as archive:
                for name in archive.namelist():
                    if name.endswith((".xml", ".rels")):
                        scan(p.relative_to(ROOT).as_posix() + "::" + name, archive.read(name).decode("utf-8"))
    check(not findings, "Security scan has unresolved findings (see local JSON)")
    source = ROOT / "data/raw/Transparency_in_Coverage_PUF.xlsx"
    check(source.is_file() and hashlib.sha256(source.read_bytes()).hexdigest() == SOURCE_HASH, "Raw source hash mismatch")
    baseline_path = ROOT / ".local-release/baseline.json"
    preservation = "NOT_RUN: local pre-release checkpoint is not part of a clone"
    if baseline_path.exists():
        baseline = json.loads(baseline_path.read_text())
        protected = [name for name in baseline if name.startswith(("powerbi/", "data/", "screenshots/"))]
        changed = [name for name in protected if not (ROOT / name).exists() or hashlib.sha256((ROOT / name).read_bytes()).hexdigest() != baseline[name]]
        check(not changed, "Protected file hash changed")
        preservation = {"checked": len(protected), "changed": changed}
    summary = {
        "status": "PASS" if not failures else "FAIL", "failures": failures,
        "publication_files": len(paths), "publication_bytes": sum(p.stat().st_size for p in paths),
        "largest_file_bytes": max(p.stat().st_size for p in paths),
        "python_compiled": len(py_files), "json_parsed": len(parsed),
        "pages": len(page_files), "visuals": len(visual_files), "navigation_destinations": nav_count,
        "screenshots": len(images), "local_links_checked": link_count,
        "security_findings": findings, "email_files_for_review": sorted(email_files),
        "preservation": preservation,
    }
    (evidence / "release_audit.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, indent=2))
    return int(bool(failures))


if __name__ == "__main__":
    sys.exit(main())
