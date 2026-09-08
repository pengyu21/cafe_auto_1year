import os

path = 'c:/antigravity/navercafe_auto/sheet_manager.py'
with open(path, 'r', encoding='cp949', errors='ignore') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if 'elif not processed_first_pending:' in line:
        new_lines.append(line)
    elif 'port_val = ""' in line and '#' in line and 'if hasattr' not in line:
        parts = line.split('port_val = ""')
        new_lines.append(parts[0].rstrip() + '\n')
        new_lines.append('                    port_val = ""\n')
    else:
        new_lines.append(line)

with open(path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Fixed!")
