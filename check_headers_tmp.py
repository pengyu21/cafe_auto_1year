def check():
    from sheet_manager import GoogleSheetManager
    mgr = GoogleSheetManager()
    headers = mgr.task_sheet.row_values(1)
    if not headers:
        headers = mgr.task_sheet.row_values(2)
    print("HEADERS:")
    for i, h in enumerate(headers):
        print(f"{i} (Col {chr(65+i)}): {h}")
check()
