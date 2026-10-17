"""DrainProof AI - prototype backend.

Pipeline: (1) vision module measures debris coverage from drain-inlet images,
(2) ML model predicts waterlogging risk per drain, (3) agent prioritises tasks,
assigns crews, schedules before the rain window and verifies completion from
after-photos. All data is SYNTHETIC (demo only).
"""
import json, math, os
import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFilter
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.metrics import roc_auc_score, precision_score, recall_score
from sklearn.model_selection import train_test_split

OUT = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(OUT, "img")
os.makedirs(IMG, exist_ok=True)
rng = np.random.default_rng(42)

# ---------------------------------------------------------------- vision
def render_drain(blockage, path, seed):
    """Draw a synthetic drain inlet photo with debris covering ~blockage of the grate."""
    r = np.random.default_rng(seed)
    W, H = 360, 240
    img = Image.new("RGB", (W, H), (96, 96, 92))  # asphalt
    px = np.array(img).astype(int)
    px += r.integers(-10, 10, px.shape)
    img = Image.fromarray(np.clip(px, 0, 255).astype("uint8"))
    d = ImageDraw.Draw(img)
    gx0, gy0, gx1, gy1 = 70, 50, 290, 190
    d.rectangle([gx0 - 8, gy0 - 8, gx1 + 8, gy1 + 8], fill=(60, 60, 60))   # frame
    d.rectangle([gx0, gy0, gx1, gy1], fill=(20, 20, 24))                    # opening
    for x in range(gx0 + 12, gx1, 22):                                      # grate bars
        d.rectangle([x, gy0, x + 8, gy1], fill=(120, 124, 128))
    # debris blobs (brown/green/plastic) until coverage ~ blockage
    area = (gx1 - gx0) * (gy1 - gy0)
    target = blockage * area
    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    covered = 0
    cols = [(110, 80, 40), (70, 100, 40), (150, 120, 70), (200, 60, 60)]
    tries = 0
    while covered < target and tries < 4000:
        tries += 1
        cx, cy = r.integers(gx0, gx1), r.integers(gy0, gy1)
        rad = int(r.integers(8, 26))
        md.ellipse([cx - rad, cy - rad, cx + rad, cy + rad], fill=255)
        covered = (np.array(mask)[gy0:gy1, gx0:gx1] > 0).sum()
    layer = Image.new("RGB", (W, H), cols[0])
    ld = ImageDraw.Draw(layer)
    for _ in range(260):
        x, y = r.integers(0, W), r.integers(0, H)
        ld.ellipse([x, y, x + int(r.integers(6, 20)), y + int(r.integers(6, 20))],
                   fill=cols[int(r.integers(0, len(cols)))])
    img.paste(layer.filter(ImageFilter.GaussianBlur(1)), (0, 0), mask)
    img.save(path, quality=88)

def measure_debris(path):
    """Vision module: % of grate opening covered by debris (colour segmentation)."""
    a = np.array(Image.open(path).convert("RGB")).astype(int)
    g = a[50:190, 70:290]
    R, G, B = g[..., 0], g[..., 1], g[..., 2]
    dark = (R < 45) & (G < 45) & (B < 50)                         # open (black) slot
    bar = (abs(R - 120) < 14) & (abs(G - 124) < 14) & (abs(B - 128) < 14)  # steel bar
    debris = ~(dark | bar)
    return float(round(100 * debris.sum() / debris.size, 1))

# ---------------------------------------------------------------- data
ZONES = ["Central Market", "Riverside", "Old Town", "Tech Park", "Station Road", "Hospital Belt"]
ZC = {z: (28.60 + 0.04 * (i // 3) + rng.normal(0, .002), 77.20 + 0.05 * (i % 3) + rng.normal(0, .002))
      for i, z in enumerate(ZONES)}

def make(n, hist=False):
    rows = []
    for i in range(n):
        z = ZONES[int(rng.integers(0, len(ZONES)))]
        la, lo = ZC[z]
        rows.append(dict(
            silt_depth_cm=float(np.clip(rng.gamma(2.2, 6), 0, 55)),
            blockage_pct=float(np.clip(rng.beta(2, 3) * 100, 0, 100)),
            cover_damaged=int(rng.random() < .12),
            days_since_clean=int(rng.integers(5, 300)),
            elevation_m=float(rng.normal(215, 4)),
            imperviousness=float(np.clip(rng.normal(.72, .15), .2, 1)),
            rain_mm_24h=float(rng.gamma(3, 14) if hist else 60.0),
            past_events=int(rng.poisson(1.2)),
            zone=z, lat=la + rng.normal(0, .009), lon=lo + rng.normal(0, .009),
            critical=int(z in ("Hospital Belt", "Station Road") and rng.random() < .5),
        ))
    return pd.DataFrame(rows)

def true_logit(d):
    return (-6.2 + 0.035 * d.silt_depth_cm + 0.034 * d.blockage_pct + 0.9 * d.cover_damaged
            + 0.006 * d.days_since_clean - 0.09 * (d.elevation_m - 215) + 1.6 * d.imperviousness
            + 0.032 * d.rain_mm_24h + 0.22 * d.past_events)

FEATS = ["silt_depth_cm", "blockage_pct", "cover_damaged", "days_since_clean",
         "elevation_m", "imperviousness", "rain_mm_24h", "past_events"]

hist = make(4000, hist=True)
hist["y"] = (rng.random(len(hist)) < 1 / (1 + np.exp(-true_logit(hist)))).astype(int)
Xtr, Xte, ytr, yte = train_test_split(hist[FEATS], hist.y, test_size=.25, random_state=1, stratify=hist.y)
model = GradientBoostingClassifier(n_estimators=200, max_depth=3, learning_rate=.07, random_state=1).fit(Xtr, ytr)
p = model.predict_proba(Xte)[:, 1]
metrics = dict(auc=round(roc_auc_score(yte, p), 3),
               precision=round(precision_score(yte, p > .5), 3),
               recall=round(recall_score(yte, p > .5), 3),
               n_train=len(Xtr), n_test=len(Xte), base_rate=round(float(hist.y.mean()), 3))
importances = sorted(zip(FEATS, model.feature_importances_.round(3).tolist()), key=lambda t: -t[1])

# ---------------------------------------------------------------- live network
N = 60
net = make(N)
net["id"] = [f"DR-{100 + i}" for i in range(N)]
# render field photos; vision module measures blockage from the pixels
meas = []
for i, r in net.iterrows():
    f = os.path.join(IMG, f"{r.id}_before.jpg")
    render_drain(r.blockage_pct / 100, f, seed=i)
    meas.append(measure_debris(f))
net["vision_blockage_pct"] = meas
vis_err = float(np.mean(np.abs(net.vision_blockage_pct - net.blockage_pct)))
net["blockage_pct"] = net["vision_blockage_pct"]            # model uses what vision sees
net["before_blockage_pct"] = net["blockage_pct"]
net["risk"] = model.predict_proba(net[FEATS])[:, 1].round(3)
net["priority"] = (net.risk * (1 + 0.6 * net.critical) * (1 + net.past_events * .05)).round(3)

# ---------------------------------------------------------------- agent
RAIN_ETA_H = 36
CREWS = [dict(id="Crew A", home="Central Market"), dict(id="Crew B", home="Hospital Belt"),
         dict(id="Crew C", home="Riverside")]
CAP_H, CLEAN_H = 12, {True: 1.0, False: 0.6}   # shift hours available, hours per drain
THRESH = .45
todo = net[net.risk >= THRESH].sort_values("priority", ascending=False).reset_index(drop=True)
log = [dict(t="T-36h", step="PREDICT", msg=f"Model scored {N} drains for 60 mm forecast; {len(todo)} above risk threshold {THRESH}.")]

def dist(a, b):
    return math.hypot(a.lat - b.lat, a.lon - b.lon)

zone_home = {c["id"]: net[net.zone == c["home"]].iloc[0] for c in CREWS}
crew_state = {c["id"]: dict(pos=zone_home[c["id"]], hours=0.0, tasks=[]) for c in CREWS}
tasks = []
for _, d in todo.iterrows():
    best = min(CREWS, key=lambda c: dist(crew_state[c["id"]]["pos"], d) + crew_state[c["id"]]["hours"] * .004)
    cs = crew_state[best["id"]]
    dur = CLEAN_H[bool(d.cover_damaged == 0)] + (0.5 if d.silt_depth_cm > 25 else 0)
    if cs["hours"] + dur > CAP_H:
        cand = [c for c in CREWS if crew_state[c["id"]]["hours"] + dur <= CAP_H]
        if not cand:
            log.append(dict(t="T-34h", step="ESCALATE", msg=f"{d.id} deferred - all crews at capacity; flagged to supervisor."))
            continue
        best = cand[0]; cs = crew_state[best["id"]]
    cs["hours"] += dur; cs["pos"] = d
    tasks.append(dict(drain=d.id, crew=best["id"], zone=d.zone, risk=float(d.risk), priority=float(d.priority),
                      slot=f"T-{RAIN_ETA_H - 2 - int(cs['hours'] * 1.5)}h", dur_h=round(dur, 1),
                      reason=("critical site; " if d.critical else "") + f"blockage {d.blockage_pct:.0f}%, silt {d.silt_depth_cm:.0f} cm, "
                              f"{int(d.days_since_clean)} d since cleaned"))
    cs["tasks"].append(d.id)
log.append(dict(t="T-35h", step="PLAN", msg=f"Assigned {len(tasks)} tasks to {len(CREWS)} crews; optimised sequence per crew; all slots end before rain onset."))
for c in CREWS:
    log.append(dict(t="T-35h", step="DISPATCH", msg=f"{c['id']}: {len(crew_state[c['id']]['tasks'])} drains, {crew_state[c['id']]['hours']:.1f} h shift."))

# field execution + verification
after_p = []
for t in tasks:
    i = net.index[net.id == t["drain"]][0]
    done = rng.random() > .12                                  # 12% of jobs are incomplete / faked
    after_b = float(rng.uniform(1, 7)) if done else float(rng.uniform(35, 60))
    f = os.path.join(IMG, f"{t['drain']}_after.jpg")
    render_drain(after_b / 100, f, seed=1000 + i)
    m = measure_debris(f)
    t.update(after_blockage=m, verified=bool(m < 15), status="VERIFIED" if m < 15 else "REOPENED")
    net.loc[i, "blockage_pct"] = m if m < 15 else net.loc[i, "blockage_pct"]
    net.loc[i, "silt_depth_cm"] = 2.0 if m < 15 else net.loc[i, "silt_depth_cm"]
    net.loc[i, "days_since_clean"] = 0 if m < 15 else net.loc[i, "days_since_clean"]
    log.append(dict(t="T-20h", step="VERIFY" if m < 15 else "REOPEN",
                    msg=f"{t['drain']} after-photo: debris {m:.0f}% -> " + ("closed." if m < 15 else "NOT cleared, task reopened & crew re-dispatched.")))
    if m >= 15:   # re-dispatch cleans it properly
        f2 = os.path.join(IMG, f"{t['drain']}_after.jpg")
        render_drain(float(rng.uniform(1, 6)) / 100, f2, seed=2000 + i)
        m2 = measure_debris(f2)
        t.update(after_blockage=m2, verified=True, status="REOPENED -> VERIFIED", redispatch=True)
        net.loc[i, ["blockage_pct", "silt_depth_cm", "days_since_clean"]] = [m2, 2.0, 0]

net["risk_after"] = model.predict_proba(net[FEATS])[:, 1].round(3)
before_hi = int((net.risk >= THRESH).sum()); after_hi = int((net.risk_after >= THRESH).sum())
log.append(dict(t="T-12h", step="REPORT", msg=f"High-risk drains {before_hi} -> {after_hi}. Citywide mean risk {net.risk.mean():.2f} -> {net.risk_after.mean():.2f}."))

reopened = sum(1 for t in tasks if t.get("redispatch"))
summary = dict(drains=N, high_risk_before=before_hi, high_risk_after=after_hi,
               tasks=len(tasks), reopened=reopened, verified=sum(1 for t in tasks if t["verified"]),
               mean_risk_before=round(float(net.risk.mean()), 3), mean_risk_after=round(float(net.risk_after.mean()), 3),
               vision_mae_pct=round(vis_err, 1), rain_mm=60, rain_eta_h=RAIN_ETA_H)

out = dict(summary=summary, model=dict(metrics=metrics, importances=importances),
           drains=json.loads(net.drop(columns=["vision_blockage_pct"]).to_json(orient="records")),
           tasks=tasks, log=log, crews=[dict(id=c, hours=round(s["hours"], 1), n=len(s["tasks"])) for c, s in crew_state.items()])
json.dump(out, open(os.path.join(OUT, "data.json"), "w"), indent=1)
print(json.dumps(dict(summary=summary, metrics=metrics, importances=importances[:4], crews=out["crews"]), indent=1))
