import json, os
from playwright.sync_api import sync_playwright

root = os.path.dirname(os.path.abspath(__file__))
tpl = open(os.path.join(root, "dashboard.template.html")).read()
data = open(os.path.join(root, "data.json")).read()
html = tpl.replace("/*DATA*/null", data)
open(os.path.join(root, "dashboard.html"), "w").write(html)
os.makedirs(os.path.join(root, "shots"), exist_ok=True)
with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={"width": 1440, "height": 810}, device_scale_factor=2)
    for v in ["map", "tasks", "verify", "log"]:
        pg.goto(f"file://{root}/dashboard.html#{v}")
        pg.reload()
        pg.wait_for_timeout(500)
        pg.screenshot(path=os.path.join(root, "shots", f"{v}.png"))
        print("shot", v)
    b.close()
