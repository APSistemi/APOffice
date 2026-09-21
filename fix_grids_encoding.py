import os
import re

def fix_and_upgrade():
    root_dir = r"C:\ApProject\APOffice"
    
    # Common UTF-8 to Windows-1252 corruptions
    replacements = {
        "Ã\xa0": "à",
        "Ã¨": "è",
        "Ã©": "é",
        "Ã¬": "ì",
        "Ã²": "ò",
        "Ã¹": "ù",
        "Ã€": "À",
        "Ãˆ": "È",
        "Ã‰": "É",
        "ÃŒ": "Ì",
        "Ã’": "Ò",
        "Ãš": "Ù",
        "â†’": "→",
        "â† ": "←",
        "â†‘": "↑",
        "â†“": "↓",
        "âœ”": "✔",
        "âœ–": "✖",
        "âž•": "✚",
        "â— ": "●",
        "â‡•": "↕",
        "â‡”": "↔",
        "ðŸ–¼": "🖼",
        "ðŸ’¡": "💡",
        "ðŸ’¾": "💾",
        "â§‰": "❐",
        "âœ–": "✖",
        "âž—": "➖",
        "â† ": "←",
        "â†’": "→"
    }

    count = 0
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.endswith(".Designer.cs") or file.endswith(".designer.cs"):
                if file.endswith(".bak"): continue
                
                path = os.path.join(root, file)
                
                # We read as bytes to handle the corruption manually or try to decode correctly
                with open(path, "rb") as f:
                    raw = f.read()
                
                try:
                    # If it's valid UTF-8, it might be the corrupted version we just wrote
                    content = raw.decode("utf-8")
                    
                    # Fix the corruptions
                    for corrupted, fixed in replacements.items():
                        content = content.replace(corrupted, fixed)
                    
                    # Also ensure the Grid replacement is done (in case it missed some or we are re-running)
                    # 1. New instance
                    content = re.sub(r'new System\.Windows\.Forms\.DataGridView\(\)', 'new APOffice.APDataGridView()', content)
                    # 2. Variable declaration
                    content = re.sub(r'private System\.Windows\.Forms\.DataGridView (\w+);', r'private APOffice.APDataGridView \1;', content)
                    # 3. Type cast or other mentions
                    content = content.replace('System.Windows.Forms.DataGridView ', 'APOffice.APDataGridView ')

                    # Write back as Windows-1252 (ANSI) which is common for these old files
                    # or just UTF-8 WITHOUT BOM if that's what we want, but MSBuild might prefer ANSI for these chars.
                    # Actually, let's use utf-8-sig (with BOM) as it helps Visual Studio recognize it.
                    with open(path, "w", encoding="utf-8-sig") as f:
                        f.write(content)
                    
                    print(f"Fixed and Upgraded: {file}")
                    count += 1
                except Exception as e:
                    print(f"Error processing {file}: {e}")

    print(f"Finished processing {count} files.")

if __name__ == "__main__":
    fix_and_upgrade()
