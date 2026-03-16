
import re

def inspect_conflicts(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Regex for conflict blocks
    # Note: git conflicts can be complex. This is a simplified regex.
    # <<<<<<< HEAD
    # ...
    # =======
    # ...
    # >>>>>>> ...
    
    pattern = re.compile(r'<<<<<<< HEAD\n(.*?)\n=======\n(.*?)\n>>>>>>> .*?\n', re.DOTALL)
    
    matches = pattern.finditer(content)
    
    for i, match in enumerate(matches):
        print(f"--- Conflict {i+1} ---")
        head_content = match.group(1)
        remote_content = match.group(2)
        
        print(f"[HEAD] Length: {len(head_content)}")
        print(f"[REMOTE] Length: {len(remote_content)}")
        
        # Show start of HEAD
        print(f"[HEAD Start]: {head_content[:100].strip()}")
        # Show start of REMOTE
        print(f"[REMOTE Start]: {remote_content[:100].strip()}")
        
        if "<ApplicationVersion>" in head_content:
            print(f"[HEAD Version]: {re.search(r'<ApplicationVersion>(.*?)</ApplicationVersion>', head_content).group(1)}")
        if "<ApplicationVersion>" in remote_content:
            print(f"[REMOTE Version]: {re.search(r'<ApplicationVersion>(.*?)</ApplicationVersion>', remote_content).group(1)}")

inspect_conflicts(r'e:\Code\QSBar\QSBar\QSBar.csproj')
