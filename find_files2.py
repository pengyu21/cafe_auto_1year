import os

target2 = 'zerooo007'
found2 = []
for root, dirs, files in os.walk('c:\\'):
    if 'Windows' in root or 'Program Files' in root:
        continue
    for d in dirs:
        if target2 in d:
            found2.append(os.path.join(root, d))
print(f"Found dirs: {found2}")
