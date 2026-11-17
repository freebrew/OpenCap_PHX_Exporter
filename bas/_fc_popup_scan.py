import re
import urllib.request

site = open(r"C:\Users\User\AppData\Local\Temp\fc_site.js", encoding="utf-8", errors="replace").read()
vals = set(re.findall(r'windowClass:"([^"]+)"', site))
print("windowClass", len(vals))
for v in sorted(vals):
    print(" ", v)

print("--- backdrop / modal class mentions ---")
for pat in ["modal-backdrop", "uib-modal-window", "modal-window", "backdropClass"]:
    print(pat, site.count(pat))
