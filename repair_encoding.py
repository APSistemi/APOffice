import os

def repair_project_encoding():
    root_dir = r"C:\ApProject\APOffice"
    
    # Map of UTF-8 byte sequences (as seen in latin-1) to the actual characters
    # This is a robust way to fix "double-encoding" or "misinterpreted encoding"
    replacements = [
        (b'\xc3\xa0', 'à'.encode('cp1252')),
        (b'\xc3\xa8', 'è'.encode('cp1252')),
        (b'\xc3\xa9', 'é'.encode('cp1252')),
        (b'\xc3\xac', 'ì'.encode('cp1252')),
        (b'\xc3\xb2', 'ò'.encode('cp1252')),
        (b'\xc3\xb9', 'ù'.encode('cp1252')),
        (b'\xc3\xa0', 'à'.encode('cp1252')),
        # Add common symbols often used in this project
        (b'\xe2\x86\x92', '→'.encode('cp1252', errors='replace')),
        (b'\xe2\x86\x90', '←'.encode('cp1252', errors='replace')),
        (b'\xe2\x86\x91', '↑'.encode('cp1252', errors='replace')),
        (b'\xe2\x86\x93', '↓'.encode('cp1252', errors='replace')),
        (b'\xe2\x9c\x94', '✔'.encode('cp1252', errors='replace')),
        (b'\xe2\x9c\x96', '✖'.encode('cp1252', errors='replace')),
        (b'\xf0\x9f\x96\xbc', '?'.encode('cp1252')), # frames / pictures
        (b'\xf0\x9f\x92\xa1', '?'.encode('cp1252')), # bulb
        (b'\xf0\x9f\x92\xbe', '?'.encode('cp1252')), # save
    ]

    count = 0
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.endswith(".cs"):
                if file.endswith(".bak"): continue
                path = os.path.join(root, file)
                
                with open(path, "rb") as f:
                    raw = f.read()
                
                new_raw = raw
                for utf8_seq, target_seq in replacements:
                    new_raw = new_raw.replace(utf8_seq, target_seq)
                
                if new_raw != raw:
                    with open(path, "wb") as f:
                        f.write(new_raw)
                    print(f"Repaired: {file}")
                    count += 1

    print(f"Finished repairing {count} files.")

if __name__ == "__main__":
    repair_project_encoding()
