import os
import re

def aggressive_repair():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        raw = f.read()
    
    # Very aggressive: anything that looks like ControlloPropriet...Changed
    # we replace with the clean version.
    # [^\w] matches any non-word byte (roughly)
    new_raw = re.sub(b'ControlloPropriet[^a-zA-Z0-9_]+_?Changed', b'ControlloPropriet\xe0_Changed', raw)
    
    # Also fix some other obvious ones
    new_raw = new_raw.replace(b'\xc3\x83\xc2\xa0', b'\xe0') # Double encoded à
    new_raw = new_raw.replace(b'\xc3\x83 ', b'\xe0')       # Ã space
    
    if new_raw != raw:
        with open(path, "wb") as f:
            f.write(new_raw)
        print("Aggressive repair applied.")
    else:
        print("No matches for aggressive repair.")

if __name__ == "__main__":
    aggressive_repair()
