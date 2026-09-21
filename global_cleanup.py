import os
import re

def global_cleanup():
    root_dir = r"C:\ApProject\APOffice"
    
    # We target the specific corrupted string pattern
    # It might look like "ControlloProprietÃ _Changed" or "ControlloProprietÃ\xa0_Changed"
    # We want "ControlloProprietà_Changed"
    
    # Also "NudPos_ValueChanged" might have issues? Let's check.
    
    count = 0
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.endswith(".cs"):
                if file.endswith(".bak"): continue
                path = os.path.join(root, file)
                
                with open(path, "rb") as f:
                    raw = f.read()
                
                # Replace the byte sequences that represent the corruption
                # 0xC3 0xA0 is 'à' in UTF-8. 
                # If it got turned into 0xC3 0x20 (Ã space), we fix it.
                new_raw = raw.replace(b'\xc3\xa0', 'à'.encode('cp1252'))
                new_raw = new_raw.replace(b'\xc3\x20', 'à'.encode('cp1252'))
                new_raw = new_raw.replace(b'\xc3', 'à'.encode('cp1252')) # Desperate fix for any lone Ã
                
                # Fix the spaces in the specific handler name if they exist
                # Match "ControlloPropriet" followed by any non-word char and then "Changed"
                # This handles cases like "ControlloPropriet à_Changed" or "ControlloProprietÃ _Changed"
                new_raw = re.sub(b'ControlloPropriet[^a-zA-Z0-9_]+Changed', b'ControlloPropriet\xe0_Changed', new_raw)
                
                # General cleanup for common Italian symbols that might be broken
                # (Same as before but with wb write)
                
                if new_raw != raw:
                    with open(path, "wb") as f:
                        f.write(new_raw)
                    print(f"Cleaned up: {file}")
                    count += 1

    print(f"Finished cleanup on {count} files.")

if __name__ == "__main__":
    global_cleanup()
