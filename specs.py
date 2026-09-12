import argparse
import os
import re
import sys
import uuid
import zipfile

try:
    import openpyxl
except ImportError:
    print("ERROR: the 'openpyxl' package is required.")
    print("Install it with:  pip install openpyxl")
    if sys.stdin.isatty():
        input("\nPress Enter to exit...")
    sys.exit(1)

def get_app_dir() -> str:

    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def xml_escape(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def rand_id() -> str:
    """8-char uppercase hex id, like Word's w14:paraId values.
    Word requires the high bit clear (value < 0x80000000)."""
    val = uuid.uuid4().int & 0x7FFFFFFE
    return f"{val:08X}"


# ---------------------------------------------------------------------------
# Template parsing
# ---------------------------------------------------------------------------

class Template:
    """Parses word/document.xml once and exposes pieces needed to rebuild it."""

    def __init__(self, document_xml: str):
        tbl_match = re.search(r"<w:tbl>.*?</w:tbl>", document_xml, re.S)
        if not tbl_match:
            raise ValueError("No <w:tbl> table found in template document.xml")
        self.tbl_start, self.tbl_end = tbl_match.span()
        self.before_tbl = document_xml[: self.tbl_start]
        self.after_tbl = document_xml[self.tbl_end :]

        tbl_xml = tbl_match.group(0)
        rows = re.findall(r"<w:tr .*?</w:tr>", tbl_xml, re.S)
        if len(rows) < 2:
            raise ValueError("Expected a header row plus at least one data row")

        first_tr_pos = tbl_xml.index(rows[0])
        self.tbl_head = tbl_xml[:first_tr_pos]  # "<w:tbl>" + tblPr + tblGrid
        self.tbl_tail = "</w:tbl>"

        self.header_row = rows[0]
        # Use the "Line 2" row (index 2: header, Line1, Line2, ...) as the
        # master template for every data row -- Line 1's row has extra
        # paragraphs unique to it, while Line 2/3/4 share simple, consistent
        # formatting. Fall back sensibly if the template has fewer rows.
        if len(rows) > 2:
            self.row_template = rows[2]
        elif len(rows) > 1:
            self.row_template = rows[1]
        else:
            self.row_template = rows[0]

        m = re.search(r"<w:t>Line \d+</w:t>", self.row_template)
        if not m:
            raise ValueError("Could not find a 'Line N' label in the template row")
        self._label_pattern = m.group(0)

        if 'w:color w:val="000000"' not in self.row_template:
            raise ValueError("Could not find the description cell placeholder in the template row")

    def build_row(self, line_number: int, description: str) -> str:
        row_xml = self.row_template

        row_xml = row_xml.replace(self._label_pattern, f"<w:t>Line {line_number}</w:t>", 1)

        def _fresh_id(m):
            return f'{m.group(1)}="{rand_id()}"'

        row_xml = re.sub(r'(w14:paraId|w14:textId|w:rsidR)="[0-9A-Fa-f]{8}"',
                          _fresh_id, row_xml)

        run_rpr = (
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
            'w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="22"/>'
            '<w:szCs w:val="22"/></w:rPr>'
        )
        lines = str(description).split("\n")
        runs = []
        for i, ln in enumerate(lines):
            if i > 0:
                runs.append(f"<w:r>{run_rpr}<w:br/></w:r>")
            runs.append(f'<w:r>{run_rpr}<w:t xml:space="preserve">{xml_escape(ln)}</w:t></w:r>')
        runs_xml = "".join(runs)

        placeholder_para = re.search(
            r'<w:p [^>]*><w:pPr><w:rPr><w:color w:val="000000"/>'
            r'<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>',
            row_xml,
        )
        if not placeholder_para:
            raise ValueError("Could not locate description placeholder paragraph in row")
        old_para = placeholder_para.group(0)
        new_para = old_para[: -len("</w:p>")] + runs_xml + "</w:p>"
        row_xml = row_xml.replace(old_para, new_para, 1)

        return row_xml

    def build_document(self, descriptions) -> str:
        data_rows = "".join(
            self.build_row(i + 1, desc) for i, desc in enumerate(descriptions)
        )
        new_tbl = self.tbl_head + self.header_row + data_rows + self.tbl_tail
        return self.before_tbl + new_tbl + self.after_tbl


# ---------------------------------------------------------------------------
# Excel reading
# ---------------------------------------------------------------------------

def read_descriptions(xlsx_path: str):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    header = [str(h).strip().upper() if h else "" for h in header_row]

    try:
        desc_col = header.index("DESCRIPTION")
    except ValueError:
        raise ValueError(f"No 'DESCRIPTION' column found (headers: {header})")

    descriptions = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if desc_col >= len(row):
            continue
        val = row[desc_col]
        if val is None or str(val).strip() == "":
            continue
        descriptions.append(str(val).strip())

    return descriptions


# ---------------------------------------------------------------------------
# docx read/write
# ---------------------------------------------------------------------------

def load_document_xml(docx_path: str) -> str:
    with zipfile.ZipFile(docx_path) as z:
        return z.read("word/document.xml").decode("utf-8")


def write_docx(template_docx_path: str, new_document_xml: str, out_path: str):
    """Copy the template docx and swap out word/document.xml."""
    with zipfile.ZipFile(template_docx_path) as zin:
        names = zin.namelist()
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name == "word/document.xml":
                    data = new_document_xml.encode("utf-8")
                zout.writestr(name, data)


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

NUM_RE = re.compile(r"(\d+)(?=\.xlsx$)", re.IGNORECASE)


def find_docx_files_recursive(folder: str):
    found = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".docx") and not f.startswith("~$"):
                found.append(os.path.join(root, f))
    return sorted(found)


def find_xlsx_files_recursive(folder: str, exclude=None):
    exclude = {os.path.abspath(p) for p in (exclude or [])}
    found = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                full = os.path.join(root, f)
                if os.path.abspath(full) not in exclude:
                    found.append(full)
    return sorted(found)


def find_xlsx_by_pos(folder: str, pos_numbers, exclude=None):
    """Return .xlsx files under folder whose filename ends with one of the
    given PO numbers (e.g. List of items - 1020.xlsx matches 1020)."""
    pos_set = {str(n).strip() for n in pos_numbers if str(n).strip()}
    if not pos_set:
        return [], []
    all_files = find_xlsx_files_recursive(folder, exclude=exclude)
    matched = []
    seen_pos = set()
    for path in all_files:
        m = NUM_RE.search(os.path.basename(path))
        if m and m.group(1) in pos_set:
            matched.append(path)
            seen_pos.add(m.group(1))
    missing = sorted(pos_set - seen_pos, key=lambda x: int(x) if x.isdigit() else x)
    return matched, missing


# ---------------------------------------------------------------------------
# Interactive selection (GUI with console fallback)
# ---------------------------------------------------------------------------

def _gui_root():
    """Try to create a hidden Tk root. Returns the root, or None if no
    display / tkinter is available (e.g. headless server)."""
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass
        return root
    except Exception:
        return None


def select_template_interactive() -> str:
    root = _gui_root()
    if root is not None:
        from tkinter import filedialog
        print("A window should have opened -- select your Specs Word template (.docx)...")
        path = filedialog.askopenfilename(
            title="Select the Specs Word template (.docx)",
            initialdir=APP_DIR,
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        root.destroy()
        if not path:
            print("No template selected. Exiting.")
            sys.exit(1)
        return path

    # console fallback
    candidates = find_docx_files_recursive(APP_DIR)
    print("\nLooking for the Word template (.docx) in:")
    print(f"  {APP_DIR}\n")
    if candidates:
        for i, c in enumerate(candidates, 1):
            print(f"  {i}. {os.path.relpath(c, APP_DIR)}")
        print("  0. Type a path manually")
        choice = input("\nSelect the template by number: ").strip()
        if choice.isdigit() and choice != "0":
            idx = int(choice)
            if 1 <= idx <= len(candidates):
                return candidates[idx - 1]
    typed = input("Enter the full path to the template .docx: ").strip().strip('"')
    if not typed or not os.path.isfile(typed):
        print("Template file not found. Exiting.")
        sys.exit(1)
    return typed


def select_input_folder_interactive(prompt_title: str = None) -> str:
    root = _gui_root()
    if root is not None:
        from tkinter import filedialog
        title = prompt_title or "Select the folder containing your List of items Excel files"
        print(f"A window should have opened -- {title.lower()}...")
        folder = filedialog.askdirectory(
            title=title,
            initialdir=APP_DIR,
        )
        root.destroy()
        if not folder:
            print("No folder selected. Exiting.")
            sys.exit(1)
        return folder

    # console fallback
    typed = input(
        f"\nEnter the folder containing your List of items Excel files "
        f"[default: {APP_DIR}]: "
    ).strip().strip('"')
    folder = typed or APP_DIR
    if not os.path.isdir(folder):
        print("Folder not found. Exiting.")
        sys.exit(1)
    return folder


def select_mode_interactive() -> str:
    """Return 'folder' or 'pos'."""
    print("\nChoose:")
    print("  1. Whole folder")
    print("  2. Write PO numbers individually or in batch")
    while True:
        choice = input("\nEnter 1 or 2 [default: 1]: ").strip()
        if choice in ("", "1"):
            return "folder"
        if choice == "2":
            return "pos"
        print("Please enter 1 or 2.")


def select_pos_numbers_interactive() -> list:
    """Prompt for PO numbers like 1020, 1533. Returns list of number strings."""
    print("\nEnter the PO number(s), separated by commas or spaces.")
    print("Example:  1020, 80230   or   1020 80230")
    while True:
        typed = input("PO numbers: ").strip()
        if not typed:
            print("Please enter at least one PO number.")
            continue
        # Split on commas and/or whitespace
        parts = re.split(r"[\s,;]+", typed)
        numbers = []
        for p in parts:
            p = p.strip()
            if not p:
                continue
            if not p.isdigit():
                print(f"  Warning: ignoring non-numeric value: {p}")
                continue
            numbers.append(p)
        if numbers:
            return numbers
        print("No valid PO numbers found. Try again.")


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def process_one(template: Template, template_docx_path: str, xlsx_path: str, out_path: str, label: str):
    descriptions = read_descriptions(xlsx_path)
    if not descriptions:
        print(f"  [skip] {label}: no descriptions found")
        return False
    new_xml = template.build_document(descriptions)
    write_docx(template_docx_path, new_xml, out_path)
    print(f"  [ok]   {label} -> {os.path.basename(out_path)}  ({len(descriptions)} item(s))")
    return True


def run(template_path: str, xlsx_files, output_dir: str, prefix: str = "Specs"):
    os.makedirs(output_dir, exist_ok=True)
    document_xml = load_document_xml(template_path)
    template = Template(document_xml)

    print(f"\nTemplate : {template_path}")
    print(f"Output   : {output_dir}")
    print(f"Files    : {len(xlsx_files)} Excel file(s)\n")

    ok, failed = 0, 0
    for xlsx_path in xlsx_files:
        base = os.path.basename(xlsx_path)
        m = NUM_RE.search(base)
        suffix = m.group(1) if m else os.path.splitext(base)[0]
        out_name = f"{prefix} {suffix}.docx"
        out_path = os.path.join(output_dir, out_name)
        try:
            if process_one(template, template_path, xlsx_path, out_path, base):
                ok += 1
            else:
                failed += 1
        except Exception as e:
            print(f"  [ERROR] {base}: {e}")
            failed += 1

    print(f"\nDone. {ok} document(s) created, {failed} skipped/failed.")
    print(f"All output saved in: {output_dir}")
    return ok, failed


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", help="Path to the template .docx (skips the picker)")
    ap.add_argument(
        "--input-dir",
        help="Folder containing List_of_items_*.xlsx files, searched recursively",
    )
    ap.add_argument(
        "--pos",
        nargs="+",
        metavar="N",
        help="PO number(s) to process (e.g. --pos 1020 1533). "
             "Requires --input-dir (or uses the program folder).",
    )
    ap.add_argument("--output-dir", help="Folder to write Specs_*.docx into (default: this program's own folder)")
    ap.add_argument("--prefix", default="Specs", help="Output filename prefix (default: Specs)")
    ap.add_argument("--yes", action="store_true", help="Don't ask for confirmation before generating")
    ap.add_argument("--no-pause", action="store_true", help="Don't wait for Enter at the end")
    args = ap.parse_args()

    output_dir = args.output_dir or APP_DIR

    print("=" * 60)
    print(" Specs Document Generator")
    print("=" * 60)

    # 1. Template
    template_path = args.template or select_template_interactive()

    # 2. Folder
    if args.input_dir:
        input_dir = args.input_dir
        if not os.path.isdir(input_dir):
            print(f"\nFolder not found: {input_dir}")
            _pause(args)
            sys.exit(1)
    else:
        input_dir = select_input_folder_interactive()

    # 3. Mode: whole folder OR specific PO numbers
    if args.pos is not None:
        mode = "pos"
        pos_numbers = [str(n) for n in args.pos]
    elif args.input_dir and args.yes:
        # CLI folder-only with --yes: process whole folder
        mode = "folder"
        pos_numbers = None
    else:
        mode = select_mode_interactive()
        pos_numbers = select_pos_numbers_interactive() if mode == "pos" else None

    if mode == "folder":
        xlsx_files = find_xlsx_files_recursive(input_dir, exclude=[template_path])
        if not xlsx_files:
            print(
                f"\nNo List of items .xlsx files found in:\n  {input_dir}\n"
                "(searched all subfolders too)"
            )
            _pause(args)
            sys.exit(1)
    else:
        xlsx_files, missing = find_xlsx_by_pos(input_dir, pos_numbers, exclude=[template_path])
        if missing:
            print(f"\nWarning: no Excel file found for PO number(s): {', '.join(missing)}")
            print(f"(searched under: {input_dir})")
        if not xlsx_files:
            print("\nNo matching Excel files found.")
            _pause(args)
            sys.exit(1)

    # Safety: never process the template itself
    template_abs = os.path.abspath(template_path)
    xlsx_files = [f for f in xlsx_files if os.path.abspath(f) != template_abs]

    if not xlsx_files:
        print("\nNo List of items .xlsx files to process.")
        _pause(args)
        sys.exit(1)

    print(f"\nSelected {len(xlsx_files)} Excel file(s):")
    for f in xlsx_files:
        try:
            print(f"   - {os.path.relpath(f, input_dir)}")
        except ValueError:
            print(f"   - {f}")

    if not args.yes:
        answer = input(
            f"\nGenerate {len(xlsx_files)} Specs document(s) into:\n  {output_dir}\n"
            "Proceed? [Y/n]: "
        ).strip().lower()
        if answer not in ("", "y", "yes"):
            print("Cancelled.")
            _pause(args)
            sys.exit(0)

    try:
        run(template_path, xlsx_files, output_dir, prefix=args.prefix)
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        _pause(args)
        sys.exit(1)

    _pause(args)


def _pause(args):
    if not args.no_pause and sys.stdin.isatty():
        try:
            input("\nPress Enter to exit...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()