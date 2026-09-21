import os

def check_bytes():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        raw = f.read()
    
    # Search for the string "ControlloPropriet"
    idx = raw.find(b"ControlloPropriet")
    if idx != -1:
        snippet = raw[idx:idx+30]
        # write to a text file
        with open("byte_output.txt", "w") as f2:
            f2.write(f"Bytes: {snippet}\n")
            f2.write(f"Hex: {snippet.hex(' ')}\n")
        print(f"Check byte_output.txt")

if __name__ == "__main__":
    check_bytes()
