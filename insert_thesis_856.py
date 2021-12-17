from tkinter import filedialog
from pathlib import Path
from openpyxl import load_workbook
# from openpyxl.workbook import Workbook
from pymarc import MARCReader, Record, Field


# def lookup_col_index(ws, search_string, col_idx=1):
#     for row in range(1, ws.max_row + 1):
#         if ws[row][col_idx].value == search_string:
#             return col_idx, row
#     return col_idx, None


def spreadsheet_lookup(lookup_file_path=None, search_string=""):
    return_val = None

    wb = load_workbook(lookup_file_path)
    ws = wb.active

    col_idx = 0

    for row in range(1, ws.max_row + 1):
        cell_value = ws[row][col_idx].value

        if cell_value == search_string:
            return_val = ws[row][1].value
            return return_val
    return return_val


def process_marc(marc_file_path=None, output_file_path=None, lookup_file_path=None):
    # Read the MARC File (mf)
    with open(marc_file_path, 'rb') as mf:
        reader = MARCReader(mf)

        # Loop through records
        for record in reader:

            f_001 = record['001'].data
            print(f_001)

            # Get 856s to loop through to update/create
            for f_856 in record.get_fields('856'):
                # u_856 = f_856['u'].data
                u_856 = f_856['u']
                print(u_856)

                # Check to make sure it is the handle.
                if 'hdl.handle.net' in u_856:
                    # Remove field if matched:
                    record.remove_field(f_856)

            # Add field based on data from lookup.
            new_856 = spreadsheet_lookup(
                lookup_file_path=lookup_file_path,
                search_string=f_001
            )

            if new_856:
                record.add_field(
                    Field(
                        tag='856',
                        indicators=['4', '1'],
                        subfields=[
                            'u', new_856,
                            'z', 'Link to OAKTrust copy'
                        ]
                    )
                )

            # Write changes to new output file.
            with open(output_file_path, 'ab') as out:
                out.write(record.as_marc())


def main():
    # Get MARC file to update.
    mrc = filedialog.askopenfile(title="Select input MRC file")

    if mrc == "" or mrc is None:
        print("User canceled operations.")
        exit()

    marc_file_path = Path(mrc.name)

    # Establish output MARC file.
    output_file_path = Path.joinpath(marc_file_path.parent.absolute(), 'output.mrc')

    # Get spreadsheet with identifier and value to insert.
    lookup = filedialog.askopenfile(title="Select the Lookup Spreadsheet")

    if lookup == "" or lookup is None:
        print("User canceled operations.")
        exit()

    lookup_file_path = Path(lookup.name)

    # Run the processing.
    process_marc(
        marc_file_path=marc_file_path,
        output_file_path=output_file_path,
        lookup_file_path=lookup_file_path
    )


if __name__ == "__main__":
    main()
