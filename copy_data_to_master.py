import os
import glob
import win32com.client


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dest_path = os.path.join(script_dir, "Ventilation_master_file.xlsx")

    pattern = os.path.join(
        script_dir,
        "Ventilation Surveys (Air Quantity - Temperature) - Olympias Mine*.xlsx",
    )
    source_files = sorted(glob.glob(pattern))

    if not source_files:
        print("No source files found in current folder.")
        return

    print(f"Found {len(source_files)} source file(s).")

    if not os.path.exists(dest_path):
        print(f"ERROR: Destination file not found: {dest_path}")
        print("Please create an empty Ventilation_master_file.xlsx first.")
        return

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        dst_wb = excel.Workbooks.Open(os.path.abspath(dest_path))
        dst_ws = dst_wb.Sheets("Sheet1")

        # Clear existing data in destination (keep row 1 for header)
        last_row = dst_ws.UsedRange.Row + dst_ws.UsedRange.Rows.Count - 1
        if last_row >= 2:
            dst_ws.Range(dst_ws.Cells(2, 1), dst_ws.Cells(last_row, 31)).Clear()

        # Copy header from the first source file
        first_wb = excel.Workbooks.Open(os.path.abspath(source_files[0]))
        first_ws = first_wb.Sheets("Data")
        for col in range(1, 32):  # A to AE = 1 to 31
            val = first_ws.Cells(1, col).Value2
            if val is not None:
                dst_ws.Cells(1, col).Value2 = val
        first_wb.Close(False)

        dst_row = 2
        total_copied = 0

        for src_file in source_files:
            print(f"Processing: {os.path.basename(src_file)}")
            src_wb = excel.Workbooks.Open(os.path.abspath(src_file))
            src_ws = src_wb.Sheets("Data")

            src_row = 2
            file_count = 0

            while True:
                val_a = src_ws.Cells(src_row, 1).Value2
                if val_a is None or val_a == "":
                    break

                for col in range(1, 32):  # A to AE
                    val = src_ws.Cells(src_row, col).Value2
                    if val is not None and val != "":
                        try:
                            dst_ws.Cells(dst_row, col).Value2 = val
                        except Exception:
                            dst_ws.Cells(dst_row, col).Value2 = str(val)

                dst_row += 1
                src_row += 1
                file_count += 1

            total_copied += file_count
            print(f"  -> {file_count} rows copied.")
            src_wb.Close(False)

        dst_wb.Save()
        print(f"\nDone. Total {total_copied} rows written to Ventilation_master_file.xlsx")

    finally:
        try:
            dst_wb.Close(False)
        except Exception:
            pass
        excel.Quit()


if __name__ == "__main__":
    main()
