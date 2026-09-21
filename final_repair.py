import os
import re

def final_repair():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        raw = f.read()
    
    # Replace anything between "ControlloPropriet" and "_Changed" with "à"
    # We use a regex that matches the prefix and suffix and replaces the middle
    new_raw = re.sub(b'ControlloPropriet.*?_Changed', b'ControlloPropriet\xe0_Changed', raw)
    
    # Also fix the weird "nÂ°" which is n°
    new_raw = new_raw.replace(b'n\xc2\xb0', b'n\xb0')
    new_raw = new_raw.replace(b'n\xc3\x82\xc2\xb0', b'n\xb0')
    
    if new_raw != raw:
        with open(path, "wb") as f:
            f.write(new_raw)
        print("Final repair applied to frmUtyEtiFormati.Designer.cs")
    else:
        print("No changes applied.")

if __name__ == "__main__":
    final_repair()
