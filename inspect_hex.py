import os

def inspect_hex():
    path = r"C:\ApProject\APOffice\frmUtyEtiFormati.Designer.cs"
    with open(path, "rb") as f:
        content = f.read()
    
    # Find all occurrences of ControlloPropriet
    import re
    for match in re.finditer(b'ControlloPropriet.{1,10}Changed', content):
        start, end = match.span()
        snippet = content[start:end]
        print(f"Match at {start}: {snippet}")
        print(f"Hex: {snippet.hex(' ')}")

if __name__ == "__main__":
    inspect_hex()
