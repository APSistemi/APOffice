import os

def check_bytes():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        raw = f.read()
    
    # Search for the string "ControlloPropriet"
    idx = raw.find(b"ControlloPropriet")
    if idx != -1:
        # print the 20 bytes after it
        snippet = raw[idx:idx+30]
        print(f"Bytes: {snippet}")
        # print hex
        print(f"Hex: {snippet.hex(' ')}")

if __name__ == "__main__":
    check_bytes()
