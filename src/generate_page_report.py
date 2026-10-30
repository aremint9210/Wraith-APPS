import sys
import os
import re
import subprocess
from pathlib import Path
from collections import defaultdict

SCRIPT_PATH = Path(__file__).resolve()
SCRIPT_DIR = SCRIPT_PATH.parent

sys.path.append(str(SCRIPT_DIR))
from generate_report import process_csv, MAPPING_DIR, BASE_DIR

MERGE_SCRIPT_PATH = SCRIPT_DIR / "dynamic_merge_tool.ps1"


def parse_page_targets(inputs, default_mapping_dir=None):
    """
    Parses strings, paths, filenames, or pasted blocks.
    Returns:
      targets: dict mapping Path(csv_file) -> list of unique page strings
      unmatched: list of strings that could not be parsed
    """
    if isinstance(inputs, str):
        raw_text = inputs
    else:
        raw_text = "\n".join(str(x) for x in inputs)

    lines = [line.strip().strip('"\'') for line in raw_text.splitlines() if line.strip()]

    pattern = re.compile(
        r'(?P<full_path>(?:[A-Za-z]:[\\/][^\r\n"\'<>*|?]+[\\/])?(?P<filename>(?P<prefix>0*\d+[a-zA-Z]?)[.\s_-]+(?P<page>page\d+)(?:\.docx)?))',
        re.IGNORECASE
    )

    targets = defaultdict(list)
    unmatched = []

    for line in lines:
        matched = False
        for m in pattern.finditer(line):
            matched = True
            prefix = m.group('prefix')
            page = m.group('page').lower()
            full_path_str = m.group('full_path')

            csv_file = None
            if os.path.isabs(full_path_str):
                p = Path(full_path_str)
                for parent in p.parents:
                    cand_map = parent / "06. EXCELL MAPPING"
                    if cand_map.exists():
                        cand_csv = cand_map / f"{prefix}.csv"
                        if cand_csv.exists():
                            csv_file = cand_csv
                            break

            if not csv_file:
                if default_mapping_dir:
                    csv_file = Path(default_mapping_dir) / f"{prefix}.csv"
                else:
                    csv_file = Path(f"{prefix}.csv")

            if page not in targets[csv_file]:
                targets[csv_file].append(page)

        if not matched:
            csv_m = re.match(r'^(?P<prefix>0*\d+[a-zA-Z]?)(?:\.csv)?\s+(?P<pages>.+)$', line, re.IGNORECASE)
            if csv_m:
                prefix = csv_m.group('prefix')
                pages_str = csv_m.group('pages')
                if default_mapping_dir:
                    csv_file = Path(default_mapping_dir) / f"{prefix}.csv"
                else:
                    csv_file = Path(f"{prefix}.csv")
                for p in re.findall(r'page\d+|all', pages_str, re.IGNORECASE):
                    if p.lower() not in targets[csv_file]:
                        targets[csv_file].append(p.lower())
            else:
                unmatched.append(line)

    return targets, unmatched


def run_page_batch(csv_path, pages_arg):
    print("=" * 60)
    print(f" WRAITH Page-Only Generator: {csv_path.name} (Pages: {pages_arg})")
    print("=" * 60)
    process_csv(csv_path, target_pages_input=pages_arg, no_merge=True)


def merge_substation_reports(csv_path):
    print("\n" + "=" * 60)
    print(f" Auto-Merging Word Reports for Substation: {csv_path.name}")
    print("=" * 60)
    if MERGE_SCRIPT_PATH.exists():
        try:
            env = os.environ.copy()
            parent_base = csv_path.parent.parent if csv_path.parent.name == "06. EXCELL MAPPING" else BASE_DIR
            env["BASE_DIR"] = str(parent_base)
            env["MAPPING_DIR"] = str(csv_path.parent)
            env["OUTPUT_DIR"] = str(parent_base / "07. OUTPUT FILE")

            subprocess.run([
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", str(MERGE_SCRIPT_PATH),
                "-csvName", csv_path.name,
                "-targetGroupsRaw", "all"
            ], check=True, env=env, cwd=str(parent_base))
            print(f" Successfully merged Word reports for {csv_path.name}")
        except subprocess.CalledProcessError as merge_err:
            print(f"Error during Word report merging for {csv_path.name}: {merge_err}")
    else:
        print(f"Warning: Merge script not found at {MERGE_SCRIPT_PATH}")


def run_page_only(csv_input, pages_arg="all", auto_merge=False):
    # If csv_input looks like a docx path or contains page references
    if ".docx" in str(csv_input).lower() or ("page" in str(csv_input).lower() and not str(csv_input).lower().endswith(".csv")):
        targets, _ = parse_page_targets([csv_input], MAPPING_DIR)
        for csv_path, pages in targets.items():
            run_page_batch(csv_path, " ".join(pages))
            if auto_merge:
                merge_substation_reports(csv_path)
        return

    csv_p = Path(csv_input)
    if csv_p.is_absolute() and csv_p.exists():
        csv_path = csv_p
    else:
        csv_path = MAPPING_DIR / csv_p.name

    if not csv_path.exists():
        print(f"Error: CSV file not found at {csv_path}")
        return

    run_page_batch(csv_path, pages_arg)
    if auto_merge:
        merge_substation_reports(csv_path)


def main():
    should_merge = False
    args = sys.argv[1:]

    if "--merge" in args:
        should_merge = True
        args = [a for a in args if a != "--merge"]
    if "--no-merge" in args:
        should_merge = False
        args = [a for a in args if a != "--no-merge"]

    inputs = []
    if "--file" in args:
        idx = args.index("--file")
        if idx + 1 < len(args):
            file_path = args[idx + 1]
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                    inputs.extend(f.readlines())
            args = args[:idx] + args[idx + 2:]

    inputs.extend(args)

    if inputs:
        # Check if first arg is csv and second arg is pages
        if len(inputs) == 2 and inputs[0].lower().endswith(".csv") and ".docx" not in inputs[0].lower():
            run_page_only(inputs[0], inputs[1], auto_merge=should_merge)
        else:
            targets, unmatched = parse_page_targets(inputs, MAPPING_DIR)
            if not targets:
                print("Could not parse any valid docx page paths or CSV entries from arguments.")
                if unmatched:
                    print(f"Unmatched inputs: {unmatched}")
                return

            print(f"Parsed {len(targets)} substation(s) for regeneration (Auto-Merge: {should_merge}):")
            for csv_f, pages in targets.items():
                print(f" - {csv_f.name}: {', '.join(pages)}")
            print()

            for csv_path, pages in targets.items():
                run_page_batch(csv_path, " ".join(pages))
                if should_merge:
                    merge_substation_reports(csv_path)
    else:
        print("=" * 60)
        print(" WRAITH Interactive Page Regenerator")
        print(" Paste docx paths (one or more) or enter '<substation>.csv <pages>':")
        print(" (Press Enter twice or EOF when finished)")
        print("=" * 60)
        lines = []
        try:
            while True:
                line = input()
                if not line:
                    break
                lines.append(line)
        except EOFError:
            pass

        if not lines:
            print("No input provided. Exiting.")
            return

        targets, _ = parse_page_targets(lines, MAPPING_DIR)
        for csv_path, pages in targets.items():
            run_page_batch(csv_path, " ".join(pages))
            if should_merge:
                merge_substation_reports(csv_path)


if __name__ == "__main__":
    main()
