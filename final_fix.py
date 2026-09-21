import os

def final_fix():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        raw = f.read()

    # We know the handler name is "ControlloProprietà_Changed"
    # and "à" in cp1252 is 0xE0.
    
    # We replace any form of "ControlloPropriet" + noise + "_Changed"
    # We'll use a byte search and replace
    
    prefix = b"ControlloPropriet"
    suffix = b"_Changed"
    
    new_raw = bytearray()
    i = 0
    while i < len(raw):
        if raw.startswith(prefix, i):
            new_raw.extend(prefix)
            new_raw.append(0xE0) # à
            new_raw.extend(suffix)
            # Find the end of the original corrupted name to skip it
            end_idx = raw.find(suffix, i + len(prefix))
            if end_idx != -1:
                i = end_idx + len(suffix)
            else:
                i += len(prefix)
        else:
            new_raw.append(raw[i])
            i += 1
            
    # Also fix some other corrupted symbols manually in bytes
    new_raw = new_raw.replace(b'\xc3\xa2\xe2\x80\xa1\xe2\x80\xa2', b'\x12') # placeholder
    # Actually, let's just use string replacement for the symbols now that we cleaned the identifiers
    content = new_raw.decode('cp1252', errors='ignore')
    content = content.replace('â‡•', '↕')
    content = content.replace('â‡”', '↔')
    content = content.replace('â†‘', '↑')
    content = content.replace('â† ', '←')
    content = content.replace('nÂ°', 'n°')
    content = content.replace('âœ”', '✔')
    content = content.replace('Ã ', 'à')
    
    with open(path, "wb") as f:
        f.write(content.encode('cp1252'))
    print("Final fix applied.")

if __name__ == "__main__":
    final_fix()
