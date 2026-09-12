from pathlib import Path
import shutil
import re


# ============================================================
# SETTINGS
# ============================================================

# The folder containing:
# List of items
# Pricing
# Specs
# Recived
# Delivery Notes
#
# Change this if necessary.
from pathlib import Path
import sys

if getattr(sys, "frozen", False):
    WORK_FOLDER = Path(sys.executable).resolve().parent
else:
    WORK_FOLDER = Path(__file__).resolve().parent


# Output folder
OUTPUT_FOLDER = WORK_FOLDER / "Grouped by PO"


# These are the ONLY folders that will be scanned.
DOCUMENT_TYPES = [
    "List of items",
    "Pricing",
    "Specs",
    "Recived",
    "Delivery Notes",
]


# Companies
COMPANIES = [
    "Abraj",
    "Ameedat",
    "Athab",
    "Wardat",
    "Wasmat",
]


# ============================================================
# FUNCTIONS
# ============================================================

def find_company(folder_name):
    """
    Finds the company from folder names such as:

        List of items Abraj
        Pricing Ameedat
        Specs Athab
        Recived Wardat
        Delivery Notes Wasmat
    """

    folder_lower = folder_name.strip().lower()

    for company in COMPANIES:
        if folder_lower.endswith(company.lower()):
            return company

    return None


def find_po_number(filename):
    """
    Finds the PO number.

    Supports 4 or 5 digit numbers.

    Examples:

        List of items 1020.xlsx
        Pricing 1040.xlsx
        Specs 1503.xlsx
        Delivery Notes 12345.pdf

    The LAST 4/5 digit number is used.
    """

    matches = re.findall(r"(?<!\d)\d{4,5}(?!\d)", filename)

    if not matches:
        return None

    return matches[-1]


def copy_file(source, destination):
    """
    Copy a file.

    If the destination already exists, don't overwrite it.
    Instead create:

        file.xlsx
        file (1).xlsx
        file (2).xlsx
        etc.
    """

    if not destination.exists():
        shutil.copy2(source, destination)
        return destination, False

    stem = destination.stem
    extension = destination.suffix

    counter = 1

    while True:

        new_destination = (
            destination.parent
            / f"{stem} ({counter}){extension}"
        )

        if not new_destination.exists():
            shutil.copy2(source, new_destination)
            return new_destination, True

        counter += 1


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 70)
    print("DOCUMENT GROUPING")
    print("=" * 70)
    print()

    # --------------------------------------------------------
    # Check Work folder
    # --------------------------------------------------------

    if not WORK_FOLDER.exists():

        print("ERROR:")
        print()
        print(f"Work folder was not found:")
        print(WORK_FOLDER)
        print()

        input("Press ENTER to exit...")
        return


    # --------------------------------------------------------
    # Create output folder
    # --------------------------------------------------------

    OUTPUT_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )


    print(f"Work folder:")
    print(f"  {WORK_FOLDER}")

    print()

    print(f"Output folder:")
    print(f"  {OUTPUT_FOLDER}")

    print()

    print("Scanning only:")

    for folder in DOCUMENT_TYPES:
        print(f"  - {folder}")

    print()


    # --------------------------------------------------------
    # Counters
    # --------------------------------------------------------

    files_found = 0
    files_copied = 0
    duplicates = 0
    skipped = 0

    skipped_files = []


    # ========================================================
    # PROCESS EACH DOCUMENT TYPE
    # ========================================================

    for document_type in DOCUMENT_TYPES:

        main_folder = WORK_FOLDER / document_type

        print("-" * 70)
        print(f"Scanning: {document_type}")
        print("-" * 70)


        # ----------------------------------------------------
        # If the main folder doesn't exist
        # ----------------------------------------------------

        if not main_folder.exists():

            print(
                f"  WARNING: Folder not found: "
                f"{main_folder}"
            )

            print()

            continue


        # ----------------------------------------------------
        # Find company folders
        # ----------------------------------------------------

        for company_folder in main_folder.iterdir():

            if not company_folder.is_dir():
                continue


            # ------------------------------------------------
            # Find company name
            # ------------------------------------------------

            company = find_company(
                company_folder.name
            )

            if company is None:

                print(
                    f"  SKIPPED folder: "
                    f"{company_folder.name}"
                )

                continue


            print()
            print(f"  Company: {company}")


            # ------------------------------------------------
            # Process files inside company folder
            # ------------------------------------------------

            for file in company_folder.rglob("*"):

                if not file.is_file():
                    continue

                files_found += 1


                # --------------------------------------------
                # Find PO number
                # --------------------------------------------

                po_number = find_po_number(
                    file.name
                )


                if po_number is None:

                    skipped += 1

                    skipped_files.append(
                        (
                            file,
                            "No 4/5 digit PO number"
                        )
                    )

                    print(
                        f"    SKIPPED: "
                        f"{file.name}"
                    )

                    continue


                # --------------------------------------------
                # Destination
                #
                # Grouped
                #   Company
                #       PO
                # --------------------------------------------

                destination_folder = (
                    OUTPUT_FOLDER
                    / company
                    / po_number
                )


                destination_folder.mkdir(
                    parents=True,
                    exist_ok=True
                )


                destination_file = (
                    destination_folder
                    / file.name
                )


                # --------------------------------------------
                # Copy
                # --------------------------------------------

                _, was_duplicate = copy_file(
                    file,
                    destination_file
                )


                files_copied += 1


                if was_duplicate:
                    duplicates += 1

                    print(
                        f"    COPIED (duplicate): "
                        f"{company} / {po_number} / "
                        f"{file.name}"
                    )

                else:

                    print(
                        f"    COPIED: "
                        f"{company} / {po_number} / "
                        f"{file.name}"
                    )


    # ========================================================
    # SUMMARY
    # ========================================================

    print()
    print()
    print("=" * 70)
    print("FINISHED")
    print("=" * 70)
    print()

    print(f"Files found     : {files_found}")
    print(f"Files copied    : {files_copied}")
    print(f"Duplicates      : {duplicates}")
    print(f"Files skipped   : {skipped}")

    print()

    print("Output:")
    print(OUTPUT_FOLDER)


    # ========================================================
    # SKIPPED FILES
    # ========================================================

    if skipped_files:

        print()
        print("=" * 70)
        print("SKIPPED FILES")
        print("=" * 70)
        print()

        for file, reason in skipped_files:

            print(file)
            print(f"  Reason: {reason}")
            print()


    print()
    input("Press ENTER to exit...")


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    main()