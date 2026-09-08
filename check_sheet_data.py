import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

from sheet_manager import GoogleSheetManager

def main():
    sm = GoogleSheetManager()
    tasks = sm.get_tasks()
    print(f"Total tasks: {len(tasks)}")
    for t in tasks[-30:]:
        if "김다운" in t['name'] or "김다미" in t['name'] or "김가영" in t['name'] or "강승남" in t['name']:
            print(f"Name: {t['name']}")
            print(f"ID: '{t['id']}'")
            print(f"Cafe: '{t['cafe_name']}'")
            print(f"Next Run: '{t['next_run']}'")
            print("-" * 20)

if __name__ == "__main__":
    main()
