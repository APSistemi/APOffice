import os

def fix_eti_formati():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        raw = f.read()

    # The corrupted sequence seems to be 0xC3 0x20
    # or just 0xC3 followed by space.
    # We want to replace it with 0xE0 (à in cp1252)
    
    # Let's be very specific: "ControlloPropriet" + corruption + "_Changed"
    # Corruption seems to be b'\xc3 ' or b'\xc3'
    
    new_raw = raw.replace(b'ControlloPropriet\xc3 ', b'ControlloPropriet\xe0')
    new_raw = new_raw.replace(b'ControlloPropriet\xc3', b'ControlloPropriet\xe0')
    
    # Also fix other symbols in this file that are broken
    new_raw = new_raw.replace(b'\xc3\xa2\xe2\x80\xa1\xe2\x80\xa2', '↕'.encode('cp1252')) # â‡•
    new_raw = new_raw.replace(b'\xc3\xa2\xe2\x80\xa1\xe2\x80\x9d', '↔'.encode('cp1252')) # â‡”
    new_raw = new_raw.replace(b'\xc3\xa2\xe2\x80\xa0\xe2\x80\x98', '↑'.encode('cp1252')) # â†‘
    
    if new_raw != raw:
        with open(path, "wb") as f:
            f.write(new_raw)
        print("Fixed frmUtyEtiFormati.Designer.cs")
    else:
        print("No changes needed or bytes didn't match.")

if __name__ == "__main__":
    fix_eti_formati()
