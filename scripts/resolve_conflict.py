
import re

file_path = r'e:\Code\QSBar\QSBar\QSBar.csproj'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Logic:
# The file has nested conflicts.
# We want to extract the "Inner HEAD" block (which contains the detailed AfterBuild target).
# And update its version to 1.0.0.5.

# Find the start of Inner HEAD
# It starts after the first ======= and then <<<<<<< HEAD
# Or simpler: search for the first <Project ...> that is followed by <Import ...>
# Line 5 starts with <Project ...
# Line 6 starts with <Import ...

# Let's use regex to find the block between <<<<<<< HEAD (inner) and ======= (inner)
# But we have outer markers too.

# content structure:
# <<<<<<< HEAD
# ...
# =======
# <<<<<<< HEAD
# [Target Content]
# =======
# ...
# >>>>>>> ...
# >>>>>>> ...

# We want [Target Content].

pattern = re.compile(r'=======\s+<<<<<<< HEAD\s+(.*?)\s+=======', re.DOTALL)
match = pattern.search(content)

if not match:
    print("Could not find Inner HEAD block with the expected pattern.")
    # Fallback: Try to find the block that contains "AfterBuild" and "taskkill"
    # and ends before =======
    pass
else:
    target_content = match.group(1)
    
    # Update version
    new_content = target_content.replace('<ApplicationVersion>1.0.0.1</ApplicationVersion>', '<ApplicationVersion>1.0.0.5</ApplicationVersion>')
    
    # Ensure it ends with </Project>
    # The Inner HEAD block in the file ended with </Target> (based on my read).
    # So we need to add </Project>.
    
    if not new_content.strip().endswith('</Project>'):
        new_content += '\n</Project>'
        
    # Write back
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print("Successfully resolved conflict in QSBar.csproj")
