"""讀 Google 表單「姓名」的選項，記錄每個名字被加進／拿掉的日期，存到 form-members.json。"""
import datetime
import json
import os
import re
import urllib.request

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSc0phHwYqcs0WPKklQhrpbEVePiLPX9lcODbEjhHI1a0TAIJw/viewform"
PATH = "form-members.json"

html = urllib.request.urlopen(FORM_URL, timeout=30).read().decode("utf-8")
m = re.search(r"FB_PUBLIC_LOAD_DATA_ = (.*?);</script>", html, re.S)
if not m:
    raise SystemExit("找不到表單資料")
questions = json.loads(m.group(1))[1][1]
names = None
for q in questions:
    if (q[1] or "").strip() in ("姓名", "名字") and q[4]:
        names = [o[0].strip() for o in q[4][0][1] if o and o[0] and o[0].strip()]
        break
if not names:
    raise SystemExit("找不到「姓名」題目的選項")

today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).date()
db = json.load(open(PATH, encoding="utf-8")) if os.path.exists(PATH) else {"baseline": today.isoformat(), "members": {}}
first_run = not db["members"]

for n in names:
    periods = db["members"].setdefault(n, [])
    if not periods:
        periods.append({"from": db["baseline"] if first_run else today.isoformat(), "to": None})
    elif periods[-1]["to"]:  # 拿掉後又加回來
        periods.append({"from": today.isoformat(), "to": None})
for n, periods in db["members"].items():
    if n not in names and periods and not periods[-1]["to"]:
        periods[-1]["to"] = (today - datetime.timedelta(days=1)).isoformat()  # 拿掉的前一天是最後一天
db["options"] = names

with open(PATH, "w", encoding="utf-8") as f:
    json.dump(db, f, ensure_ascii=False, indent=1)
    f.write("\n")
