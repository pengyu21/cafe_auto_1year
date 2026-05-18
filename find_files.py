import os

target = 'port_mapping.json'
found = []
for root, dirs, files in os.walk('c:\\'):
    if 'Windows' in root or 'Program Files' in root:
        continue
    if target in files:
        found.append(os.path.join(root, target))
print(f"Found: {found}")

target2 = 'zerooo007'
found2 = []
for root, dirs, files in os.walk('c:\\'):
    if 'Windows' in root or 'Program Files' in root:
        continue
    if target2 in dirs:
        found2.append(os.path.join(root, target2))
print(f"Found dirs: {found2}")
