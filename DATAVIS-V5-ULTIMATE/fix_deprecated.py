import re

with open("app.py", "r", encoding="utf-8") as f:
    content = f.read()

old_true = "use_container_width=True"
new_true = "width='stretch'"
old_false = "use_container_width=False"
new_false = "width='content'"

count_true = content.count(old_true)
count_false = content.count(old_false)

content = content.replace(old_true, new_true)
content = content.replace(old_false, new_false)

with open("app.py", "w", encoding="utf-8") as f:
    f.write(content)

print(f"Replaced {count_true} occurrences of use_container_width=True")
print(f"Replaced {count_false} occurrences of use_container_width=False")
remaining = content.count("use_container_width")
print(f"Remaining use_container_width occurrences: {remaining}")
