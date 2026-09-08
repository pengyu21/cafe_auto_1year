import os

path = 'c:/antigravity/navercafe_auto/sheet_manager.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    new_lines.append(line)
    if 'elif not processed_first_pending:' in line:
        # Check if the next line is a comment
        pass
    if "elif not processed_first_pending:" in lines[i-1] if i > 0 else False:
        # Ensure we add the missing line
        new_lines.append("                         should_create = True\n")

with open(path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Fixed!")
