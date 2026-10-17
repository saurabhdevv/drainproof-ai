const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const ROOT = "/home/claude/drainproof";
const D = JSON.parse(fs.readFileSync(path.join(ROOT, "data.json"), "utf8"));
const S = D.summary;

const THEME = {
  name: "DrainProof",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "0F1F33", lt1: "FFFFFF", dk2: "0B2540", lt2: "F4F7FB",
    accent1: "0FA3A3", accent2: "E8A317", accent3: "D64545",
    accent4: "2F9E6B", accent5: "2E6FD8", accent6: "5B6B80",
    hlink: "2E6FD8", folHlink: "5B6B80",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "DrainProof AI - Concept note and pitch";
pres.author = "DrainProof AI";
const C = pres.SchemeColor;

// ------------------------------------------------------------ layouts
const FOOT = { x: 0.6, y: 7.02, w: 8, h: 0.3, fontSize: 10, color: HEX.accent6 };
pres.defineSlideMaster({
  title: "TITLE_DARK", background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.05, w: 7.2, h: 1.2, fontSize: 60, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.3, w: 7.2, h: 0.8, fontSize: 28, color: C.accent1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: HEX.lt2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.95, fontSize: 32, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "DrainProof AI  |  SDG 11 concept note and pitch", options: { ...FOOT, isTextBox: true, margin: 0 } } },
  ],
  slideNumber: { x: 12.1, y: 7.02, w: 0.6, h: 0.3, fontSize: 10, color: HEX.accent6, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSE_DARK", background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.9, w: 11.7, h: 1.3, fontSize: 50, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
});

// ------------------------------------------------------------ helpers
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 8, offset: 2, angle: 90 });
function card(slide, x, y, w, h, name, fill = C.background1) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill, width: 0 }, shadow: shadow(), objectName: name });
}
async function iconCircle(slide, iconName, x, y, d, circleColor, iconColor, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: circleColor }, line: { color: circleColor, width: 0 }, objectName: name + " circle" });
  const p = d * 0.5;
  slide.addImage({ data: await icon(iconName, iconColor), x: x + (d - p) / 2, y: y + (d - p) / 2, w: p, h: p, objectName: name + " icon", altText: name });
}
function bullets(items, opts = {}) {
  return items.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 8, ...opts } }));
}
function arrow(slide, x, y, w, h, color, name) {
  slide.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color }, line: { color, width: 0 }, objectName: name });
}
const pct = (a, b) => Math.round(((a - b) / a) * 100);

async function main() {
  // ======================================================== 1. TITLE
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  s.addText("SDG 11  |  Sustainable Cities and Communities", { x: 0.8, y: 1.35, w: 7.2, h: 0.4, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "SDG kicker" });
  s.addText("DrainProof AI", { placeholder: "title" });
  s.addText("Fix the drain before the rain", { placeholder: "body" });
  s.addText("An agentic AI and machine learning platform that predicts which urban drains will cause waterlogging, then makes sure the maintenance actually happens, with photo proof, before heavy rain arrives.", { x: 0.8, y: 4.2, w: 6.8, h: 1.3, fontSize: 16, color: C.background2, margin: 0, valign: "top", isTextBox: true, objectName: "Description" });
  s.addText("Internship project  |  Concept note and pitch  |  October 2026", { x: 0.8, y: 6.6, w: 7, h: 0.35, fontSize: 14, color: C.background2, margin: 0, isTextBox: true, objectName: "Context line" });
  const hero = "DR-106";
  const heroB = D.drains.find((d) => d.id === hero).before_blockage_pct;
  const heroA = D.tasks.find((t) => t.drain === hero).after_blockage;
  [["before", 1.3, "BEFORE  |  " + Math.round(heroB) + "% blocked", HEX.accent3], ["after", 4.2, "AFTER  |  " + Math.round(heroA) + "% blocked, verified", HEX.accent4]].forEach(([k, y, label, col]) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.45, y: y - 0.1, w: 4.2, h: 2.9, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.background1, width: 0 }, objectName: "Photo frame " + k });
    s.addImage({ path: path.join(ROOT, "img", `${hero}_${k}.jpg`), x: 8.55, y: y, w: 3.9, h: 2.6, objectName: "Drain photo " + k, altText: `Drain inlet photo ${k} cleaning` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.7, y: y + 2.0, w: 2.9, h: 0.4, rectRadius: 0.2, fill: { color: col }, line: { color: col, width: 0 }, objectName: "Chip " + k });
    s.addText(label, { x: 8.7, y: y + 2.0, w: 2.9, h: 0.4, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Chip text " + k });
  });
  s.addNotes("Open with the one-line promise: fix the drain before the rain. The two photos are from the prototype (synthetic demo images): the same drain before the crew visit and after, with the AI measuring the debris on the grate.");

  // ======================================================== 2. CONCEPT NOTE
  pres.addSection({ title: "Concept note" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Concept note" });
  s.addText("Concept note at a glance: why, what, whom", { placeholder: "title" });
  const cols = [
    { tag: "WHY", ic: "FaWater", col: HEX.accent3, head: "Waterlogging from blocked drains is preventable", b: ["Drains clog with silt, plastic and debris between cleaning rounds", "Teams lack one live view of drain condition and urgency", "Jobs get marked done with no proof the drain is actually clear"] },
    { tag: "WHAT", ic: "FaBrain", col: HEX.accent1, head: "An agent that closes the loop from risk to fix", b: ["ML predicts waterlogging risk for every drain", "The agent ranks, assigns and schedules cleaning before rain arrives", "After-photos are checked, and failed jobs are reopened automatically"] },
    { tag: "WHOM", ic: "FaCity", col: HEX.accent5, head: "Cities and large facilities that own drains", b: ["Municipal drainage and stormwater departments", "Campuses, airports, industrial parks and malls", "Field crews and supervisors use it daily; residents get drier streets"] },
  ];
  for (let i = 0; i < 3; i++) {
    const c = cols[i], x = 0.6 + i * 4.165;
    card(s, x, 1.6, 3.8, 4.6, "Card " + c.tag);
    await iconCircle(s, c.ic, x + 0.3, 1.9, 0.75, c.col, "FFFFFF", c.tag);
    s.addText(c.tag, { x: x + 1.2, y: 1.9, w: 2.3, h: 0.75, fontSize: 24, bold: true, fontFace: "Cambria", color: c.col, valign: "middle", margin: 0, isTextBox: true, objectName: c.tag + " label" });
    s.addText(c.head, { x: x + 0.3, y: 2.85, w: 3.2, h: 1.0, fontSize: 18, bold: true, color: C.text2, valign: "top", margin: 0, isTextBox: true, objectName: c.tag + " heading" });
    s.addText(bullets(c.b), { x: x + 0.3, y: 3.95, w: 3.2, h: 2.65, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: c.tag + " points" });
  }
  s.addNotes("This is the concept note on one slide. WHY: the problem is preventable and keeps happening because follow-through is weak. WHAT: not another dashboard, an agent that owns the work until it is verified. WHOM: organisations that own drainage assets and the crews who maintain them.");

  // ======================================================== 3. THE GAP
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Concept note" });
  s.addText("Dashboards flag risk. DrainProof gets it fixed", { placeholder: "title" });
  s.addText("TODAY", { x: 0.6, y: 1.6, w: 4, h: 0.35, fontSize: 14, bold: true, color: C.accent6, margin: 0, isTextBox: true, objectName: "Today label" });
  const today = [["Risk spotted", C.accent6], ["Report filed", C.accent6], ["Cleaning maybe done", C.accent2], ["Street floods", C.accent3]];
  today.forEach(([t, col], i) => {
    const x = 0.6 + i * 3.19;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.05, w: 2.55, h: 1.15, rectRadius: 0.12, fill: { color: col }, line: { color: col, width: 0 }, objectName: "Today step " + (i + 1) });
    s.addText(t, { x, y: 2.05, w: 2.55, h: 1.15, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", margin: 6, isTextBox: true, objectName: "Today text " + (i + 1) });
    if (i < 3) arrow(s, x + 2.65, 2.45, 0.4, 0.35, HEX.accent6, "Today arrow " + (i + 1));
  });
  s.addText("WITH DRAINPROOF", { x: 0.6, y: 3.65, w: 5, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "DrainProof label" });
  const dp = [["Predict", C.accent1], ["Prioritise", C.accent1], ["Dispatch", C.accent1], ["Verify", C.accent1], ["Dry street", C.accent4]];
  dp.forEach(([t, col], i) => {
    const x = 0.6 + i * 2.53;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.1, w: 2.0, h: 1.15, rectRadius: 0.12, fill: { color: col }, line: { color: col, width: 0 }, objectName: "DrainProof step " + (i + 1) });
    s.addText(t, { x, y: 4.1, w: 2.0, h: 1.15, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", margin: 6, isTextBox: true, objectName: "DrainProof text " + (i + 1) });
    if (i < 4) arrow(s, x + 2.08, 4.5, 0.37, 0.35, HEX.accent1, "DrainProof arrow " + (i + 1));
  });
  card(s, 0.6, 5.65, 12.13, 1.1, "Insight card");
  s.addText([
    { text: "The missing piece is follow-through. ", options: { bold: true, color: HEX.dk2 } },
    { text: "An agent that assigns the work, books it before the rain, checks the photo evidence and reopens anything left undone." },
  ], { x: 0.9, y: 5.65, w: 11.55, h: 1.1, fontSize: 18, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "Insight text" });
  s.addNotes("Most flood-risk tools end at 'here is a map of risky places'. Nobody checks whether the drain was cleaned before the storm. That gap between knowing and fixing is where DrainProof sits.");

  // ======================================================== 4. HOW IT WORKS
  pres.addSection({ title: "Solution" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution" });
  s.addText("One loop: see, predict, act, verify", { placeholder: "title" });
  const steps = [
    { n: "1", ic: "FaCamera", t: "See", d: "A vision model reads each field photo and measures how much of the grate is covered by debris.", col: HEX.accent5 },
    { n: "2", ic: "FaChartLine", t: "Predict", d: "A gradient-boosting model scores waterlogging risk per drain, using blockage, silt, history and the rain forecast.", col: HEX.accent2 },
    { n: "3", ic: "FaRoute", t: "Act", d: "The agent ranks drains, assigns the nearest crew with capacity and books every slot before rain onset.", col: HEX.accent1 },
    { n: "4", ic: "FaClipboardCheck", t: "Verify", d: "After-photos are checked. Anything still blocked is rejected and the crew is re-dispatched.", col: HEX.accent4 },
  ];
  for (let i = 0; i < 4; i++) {
    const c = steps[i], x = 0.6 + i * 3.06;
    card(s, x, 1.7, 2.85, 4.2, "Step card " + c.t);
    await iconCircle(s, c.ic, x + 0.3, 2.0, 0.85, c.col, "FFFFFF", c.t);
    s.addText(c.n, { x: x + 2.1, y: 2.0, w: 0.5, h: 0.5, fontSize: 28, bold: true, fontFace: "Cambria", color: HEX.accent6, align: "right", margin: 0, isTextBox: true, objectName: "Step number " + c.n });
    s.addText(c.t, { x: x + 0.3, y: 3.15, w: 2.3, h: 0.5, fontSize: 24, bold: true, fontFace: "Cambria", color: C.text2, margin: 0, isTextBox: true, objectName: "Step title " + c.t });
    s.addText(c.d, { x: x + 0.3, y: 3.75, w: 2.3, h: 2.0, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "Step text " + c.t });
  }
  s.addText("The loop repeats before every forecast heavy-rain event, and each pass makes the risk model's picture of the city fresher.", { x: 0.6, y: 6.15, w: 12.13, h: 0.6, fontSize: 16, italic: true, color: C.accent6, margin: 0, isTextBox: true, objectName: "Loop note" });
  s.addNotes("Walk the four steps. Emphasise that steps 3 and 4 are the agentic part: the system does not just inform, it plans, dispatches, checks evidence and reopens work.");

  // ======================================================== 5. ARCHITECTURE
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution" });
  s.addText("Architecture: from raw signals to verified work", { placeholder: "title" });
  const arch = [
    { h: "INPUTS", col: HEX.accent6, items: [["Field photos", "crew phone or CCTV"], ["Rain forecast", "weather API, rain gauges"], ["Drain register", "location, size, elevation"], ["History", "cleaning logs, complaints"]] },
    { h: "INTELLIGENCE", col: HEX.accent5, items: [["Vision model", "reads debris % from a photo"], ["Risk model", "gradient boosting predicts waterlogging risk"]] },
    { h: "AGENT", col: HEX.accent1, items: [["Prioritise", "risk x criticality"], ["Assign and schedule", "nearest crew, before rain"], ["Verify and reopen", "after-photo check"], ["Escalate", "no capacity: tell supervisor"]] },
    { h: "OUTPUTS", col: HEX.accent4, items: [["Ops dashboard", "map, ranking, KPIs"], ["Crew task list", "slot, place, reason"], ["Audit trail", "photo proof per drain"]] },
  ];
  const top0 = 2.3, avail = 3.45;
  arch.forEach((c, i) => {
    const x = 0.6 + i * 3.14;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 2.7, h: 0.5, rectRadius: 0.25, fill: { color: c.col }, line: { color: c.col, width: 0 }, objectName: c.h + " header" });
    s.addText(c.h, { x, y: 1.6, w: 2.7, h: 0.5, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: c.h + " header text" });
    const n = c.items.length, ih = (avail - (n - 1) * 0.12) / n;
    c.items.forEach(([t, sub], j) => {
      const y = top0 + j * (ih + 0.12);
      card(s, x, y, 2.7, ih, `${c.h} item ${j + 1}`);
      s.addText([{ text: t, options: { bold: true, fontSize: 14, color: HEX.dk2, breakLine: true } }, { text: sub, options: { fontSize: 12, color: HEX.accent6 } }], { x: x + 0.15, y, w: 2.4, h: ih, valign: "middle", margin: 0, isTextBox: true, objectName: `${c.h} item text ${j + 1}` });
    });
    if (i < 3) arrow(s, x + 2.74, 3.75, 0.36, 0.3, HEX.accent6, "Flow arrow " + (i + 1));
  });
  card(s, 0.6, 6.0, 12.13, 0.8, "Stack card");
  s.addText([
    { text: "Prototype: ", options: { bold: true, color: HEX.dk2 } },
    { text: "Python, scikit-learn, image analysis, HTML dashboard.   " },
    { text: "Production path: ", options: { bold: true, color: HEX.dk2 } },
    { text: "crew mobile app, live weather feed, deployable on IBM watsonx.ai or any cloud." },
  ], { x: 0.85, y: 6.0, w: 11.65, h: 0.8, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "Stack text" });
  s.addNotes("Four layers. Inputs are cheap to get: crew phone photos, a weather feed and the drain register the city already has. The agent layer is where the value is. The prototype runs on a laptop; the production path is a crew mobile app and a live weather feed, on IBM watsonx.ai or any cloud.");

  // ======================================================== 6-9 SCREENSHOTS
  pres.addSection({ title: "Live demo screens" });
  async function shotSlide(title, file, cards, notes) {
    const sl = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Live demo screens" });
    sl.addText(title, { placeholder: "title" });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.55, y: 1.5, w: 8.3, h: 4.7, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.background1, width: 0 }, shadow: shadow(), objectName: "Screenshot frame" });
    sl.addImage({ path: path.join(ROOT, "shots", file), x: 0.65, y: 1.57, w: 8.1, h: 4.556, objectName: "Prototype screenshot", altText: title });
    sl.addText("Prototype screenshot. All data shown is synthetic demo data.", { x: 0.6, y: 6.35, w: 8.2, h: 0.3, fontSize: 11, italic: true, color: C.accent6, margin: 0, isTextBox: true, objectName: "Screenshot caption" });
    cards.forEach(([lead, body], i) => {
      const y = 1.5 + i * 1.6;
      card(sl, 9.15, y, 3.58, 1.5, "Callout card " + (i + 1));
      sl.addText([{ text: lead, options: { bold: true, fontSize: 16, color: HEX.dk2, breakLine: true } }, { text: body, options: { fontSize: 14, color: HEX.dk1 } }], { x: 9.35, y, w: 3.2, h: 1.5, valign: "middle", margin: 0, isTextBox: true, objectName: "Callout text " + (i + 1) });
    });
    sl.addNotes(notes);
  }
  const hi = S.high_risk_before;
  await shotSlide("See the whole city's flood risk at once", "map.png", [
    [`${hi} of ${S.drains} drains`, `flagged high risk for the ${S.rain_mm} mm rain forecast`],
    ["Critical sites rank first", "Hospital and station drains outrank equal-risk drains elsewhere"],
    ["A clear first-to-fix list", "The top-10 table tells crews exactly where to go first"],
  ], "Screen 1: the risk map. Each dot is a drain, coloured by predicted waterlogging risk. Dashed rings are critical sites. The table on the right is the agent's first-to-fix list.");
  await shotSlide("The agent turns risk into a booked work plan", "tasks.png", [
    ["Priority = risk x criticality", "plus flood history, so the worst drains go first"],
    ["Crews balanced", `${D.crews.map((c) => c.hours).join(", ")} h booked on a 12 h shift`],
    ["Done before the rain", "Every slot is scheduled to finish ahead of rain onset"],
  ], "Screen 2: the plan the agent produced without a human. Each row has a crew, a time slot and the evidence for why that drain was chosen.");
  await shotSlide("Photo evidence proves the drain is clear", "verify.png", [
    ["Photo, not a tick-box", "Debris % is measured from before and after photos"],
    ["Failures get reopened", `${S.reopened} of ${S.tasks} jobs failed the first check and were re-dispatched`],
    [`About ${S.vision_mae_pct} points of error`, "average gap between measured and true debris in the demo"],
  ], "Screen 3: verification. Before and after photos side by side with the measured debris percentage. A job above 15% debris is rejected and the crew goes back.");
  await shotSlide("Every step is logged and measurable", "log.png", [
    ["Full audit trail", "Forecast, plan, dispatch and proof recorded per drain"],
    [`Mean risk down ${pct(S.mean_risk_before, S.mean_risk_after)}%`, `from ${S.mean_risk_before.toFixed(2)} to ${S.mean_risk_after.toFixed(2)} across the city`],
    [`${S.high_risk_before} to ${S.high_risk_after} high-risk drains`, "re-scored by the same model after cleaning"],
  ], "Screen 4: the agent's activity log and the before-and-after risk. This is the evidence an operations manager would take to a review meeting.");

  // ======================================================== 10. RESULTS
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Live demo screens" });
  s.addText(`In the demo, ${S.high_risk_before} high-risk drains fall to ${S.high_risk_after} before the rain`, { placeholder: "title" });
  const stats = [
    [`${S.high_risk_before} to ${S.high_risk_after}`, "high-risk drains", HEX.accent4],
    [`${S.mean_risk_before.toFixed(2)} to ${S.mean_risk_after.toFixed(2)}`, "mean citywide risk", HEX.accent1],
    [`${S.verified}/${S.tasks}`, `jobs verified by photo, ${S.reopened} after a reopen`, HEX.accent5],
    [`${D.crews.length} crews`, "all slots finish before rain onset", HEX.accent2],
  ];
  stats.forEach(([big, lab, col], i) => {
    const x = 0.6 + (i % 2) * 2.95, y = 1.6 + Math.floor(i / 2) * 2.45;
    card(s, x, y, 2.8, 2.3, "Stat card " + (i + 1));
    s.addText(big, { x: x + 0.2, y: y + 0.2, w: 2.4, h: 0.85, fontSize: 30, bold: true, fontFace: "Cambria", color: col, valign: "middle", margin: 0, isTextBox: true, objectName: "Stat value " + (i + 1) });
    s.addText(lab, { x: x + 0.2, y: y + 1.1, w: 2.4, h: 1.0, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "Stat label " + (i + 1) });
  });
  const zones = [...new Set(D.drains.map((d) => d.zone))];
  const zb = [], za = [];
  zones.forEach((z) => {
    const a = D.drains.filter((d) => d.zone === z);
    zb.push(+(a.reduce((t, d) => t + d.risk, 0) / a.length).toFixed(2));
    za.push(+(a.reduce((t, d) => t + d.risk_after, 0) / a.length).toFixed(2));
  });
  card(s, 6.7, 1.6, 6.03, 5.15, "Chart card");
  s.addChart(pres.charts.BAR, [
    { name: "Before", labels: zones, values: zb },
    { name: "After", labels: zones, values: za },
  ], {
    x: 6.85, y: 1.7, w: 5.75, h: 4.95, barDir: "col", barGrouping: "clustered",
    chartColors: [HEX.accent3, HEX.accent4],
    showTitle: true, title: "Mean waterlogging risk by zone", titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", dataLabelFontSize: 11, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt",
    valAxisLabelColor: HEX.accent6, valAxisLabelFontSize: 11, valAxisLabelFontFace: "+mn-lt", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0.0",
    valGridLine: { color: "E1E8F0", size: 0.75 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: HEX.dk1, legendFontFace: "+mn-lt",
    chartArea: { fill: { color: "FFFFFF" } }, plotArea: { fill: { color: "FFFFFF" } },
  });
  s.addNotes(`Simulated city of ${S.drains} drains and a ${S.rain_mm} mm forecast, ${S.rain_eta_h} hours out. These numbers come from the prototype on synthetic data, so say clearly that they show how the system behaves, not field results. Real results come from the pilot.`);

  // ======================================================== 11. MODEL
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Live demo screens" });
  s.addText("Under the hood: an honest look at the model", { placeholder: "title" });
  const M = D.model.metrics;
  [[M.auc.toFixed(2), "AUC on held-out data", HEX.accent5], [Math.round(M.precision * 100) + "%", "precision at 0.5 threshold", HEX.accent1], [Math.round(M.recall * 100) + "%", "recall at 0.5 threshold", HEX.accent2]].forEach(([v, l, col], i) => {
    const y = 1.6 + i * 1.5;
    card(s, 0.6, y, 3.6, 1.35, "Metric card " + (i + 1));
    s.addText(v, { x: 0.85, y, w: 1.5, h: 1.35, fontSize: 34, bold: true, fontFace: "Cambria", color: col, valign: "middle", margin: 0, isTextBox: true, objectName: "Metric value " + (i + 1) });
    s.addText(l, { x: 2.4, y, w: 1.65, h: 1.35, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "Metric label " + (i + 1) });
  });
  card(s, 0.6, 6.1, 3.6, 0.7, "Dataset card");
  s.addText(`${(M.n_train + M.n_test).toLocaleString("en-US")} simulated drain-events`, { x: 0.8, y: 6.1, w: 3.3, h: 0.7, fontSize: 14, bold: true, color: HEX.dk2, valign: "middle", margin: 0, isTextBox: true, objectName: "Dataset text" });
  const imp = D.model.importances;
  const nice = (k) => ({ rain_mm_24h: "Rain forecast (24 h)", blockage_pct: "Inlet blockage", days_since_clean: "Days since cleaned", silt_depth_cm: "Silt depth", past_events: "Past flood events", imperviousness: "Paved catchment", elevation_m: "Elevation", cover_damaged: "Damaged cover" }[k] || k);
  card(s, 4.45, 1.6, 8.28, 4.0, "Importance chart card");
  s.addChart(pres.charts.BAR, [{ name: "Importance", labels: imp.map((i) => nice(i[0])), values: imp.map((i) => Math.round(i[1] * 100)) }], {
    x: 4.6, y: 1.7, w: 8.0, h: 3.8, barDir: "bar", chartColors: [HEX.accent5],
    catAxisOrientation: "maxMin",
    showTitle: true, title: "What drives predicted risk (feature importance, %)", titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0", dataLabelFontSize: 11, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt",
    valAxisLabelColor: HEX.accent6, valAxisLabelFontSize: 11, valAxisLabelFontFace: "+mn-lt", valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    chartArea: { fill: { color: "FFFFFF" } }, plotArea: { fill: { color: "FFFFFF" } },
  });
  card(s, 4.45, 5.8, 8.28, 1.0, "Caveat card");
  s.addText([{ text: "Caveat: ", options: { bold: true, color: HEX.accent3 } }, { text: "trained on simulated data, so these scores show the pipeline works, not real-world accuracy. The pilot replaces them with measured results." }], { x: 4.7, y: 5.8, w: 7.8, h: 1.0, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "Caveat text" });
  s.addNotes("Be upfront: the data is synthetic, so the AUC says the pipeline learns what it was taught, nothing more. The feature ranking is still useful: it shows the model leans on rain, blockage and time since cleaning, which is what field engineers would expect. Real validation happens in the pilot.");

  // ======================================================== 12. WHOM / BUSINESS
  pres.addSection({ title: "Business" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Business" });
  s.addText("Who buys it, who uses it, how it earns", { placeholder: "title" });
  const biz = [
    { ic: "FaCity", col: HEX.accent5, t: "Who buys", b: ["City and municipal drainage departments", "Smart-city programmes", "Large facilities: campuses, airports, industrial parks, malls"] },
    { ic: "FaHardHat", col: HEX.accent1, t: "Who uses it daily", b: ["Field crews get a ranked task list with photo upload", "Supervisors see plan, progress and proof", "Residents benefit from fewer flooded streets"] },
    { ic: "FaCoins", col: HEX.accent2, t: "How it earns", b: ["Annual subscription per monitored drain", "One-time onboarding: drain register and photo set-up", "Pilot first, pricing validated with the partner"] },
  ];
  for (let i = 0; i < 3; i++) {
    const c = biz[i], x = 0.6 + i * 4.165;
    card(s, x, 1.6, 3.8, 4.6, "Business card " + c.t);
    await iconCircle(s, c.ic, x + 0.3, 1.9, 0.75, c.col, "FFFFFF", c.t);
    s.addText(c.t, { x: x + 0.3, y: 2.9, w: 3.2, h: 0.6, fontSize: 22, bold: true, fontFace: "Cambria", color: C.text2, margin: 0, isTextBox: true, objectName: c.t + " heading" });
    s.addText(bullets(c.b), { x: x + 0.3, y: 3.65, w: 3.2, h: 2.95, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: c.t + " points" });
  }
  s.addNotes("For employers: the buyer is whoever is accountable for flooding and drain maintenance. Pricing is a hypothesis to test in the pilot, not a quote.");

  // ======================================================== 13. ROADMAP
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Business" });
  s.addText("Roadmap: from prototype to a real pilot", { placeholder: "title" });
  const rm = [
    { w: "Done", t: "Working prototype", d: "Risk model, agent loop, photo check and dashboard on synthetic data", col: HEX.accent4 },
    { w: "Months 1-2", t: "Real data", d: "Drain register, rain history and a labelled photo set from a partner zone", col: HEX.accent5 },
    { w: "Months 3-4", t: "Field pilot", d: "Crew mobile app and live weather feed in one pilot zone", col: HEX.accent1 },
    { w: "Months 5-6", t: "Measure and scale", d: "Compare against untreated zones, then widen to the whole city", col: HEX.accent2 },
  ];
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 2.3, w: 11.3, h: 0, line: { color: HEX.accent6, width: 2, dashType: "dash" }, objectName: "Timeline line" });
  rm.forEach((r, i) => {
    const x = 0.6 + i * 3.06;
    s.addShape(pres.shapes.OVAL, { x: x + 1.2, y: 2.05, w: 0.5, h: 0.5, fill: { color: r.col }, line: { color: "FFFFFF", width: 2 }, objectName: "Milestone dot " + (i + 1) });
    s.addText(r.w, { x, y: 1.55, w: 2.85, h: 0.4, fontSize: 14, bold: true, color: r.col, align: "center", margin: 0, isTextBox: true, objectName: "Phase time " + (i + 1) });
    card(s, x, 2.85, 2.85, 2.2, "Phase card " + (i + 1));
    s.addText([{ text: r.t, options: { bold: true, fontSize: 18, fontFace: "Cambria", color: HEX.dk2, breakLine: true } }, { text: r.d, options: { fontSize: 14, color: HEX.dk1 } }], { x: x + 0.2, y: 2.85, w: 2.45, h: 2.2, valign: "middle", margin: 0, isTextBox: true, objectName: "Phase text " + (i + 1) });
  });
  card(s, 0.6, 5.35, 12.13, 1.4, "Success metrics card");
  s.addText([
    { text: "Pilot success metrics", options: { bold: true, fontSize: 16, color: HEX.dk2, breakLine: true } },
    { text: "Share of high-risk drains cleaned before the rain  |  Share of completed jobs verified by photo  |  Waterlogging complaints versus untreated zones  |  Time from alert to verified fix", options: { fontSize: 14, color: HEX.dk1 } },
  ], { x: 0.9, y: 5.35, w: 11.55, h: 1.4, valign: "middle", margin: 0, isTextBox: true, objectName: "Success metrics text" });
  s.addNotes("The roadmap is deliberately short. The first real milestone is getting a partner zone's drain register and a few hundred labelled photos so the vision model can be trained on real drains.");

  // ======================================================== 14. CLOSING
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "CLOSE_DARK", sectionTitle: "Close" });
  s.addText("Fix the drain before the rain", { placeholder: "title" });
  s.addText("Aligned with the UN Sustainable Development Goals", { x: 0.8, y: 2.5, w: 11.7, h: 0.4, fontSize: 16, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "SDG heading" });
  const sdg = [["11.5", "Reduce losses from water-related disasters"], ["9.1", "Resilient urban infrastructure"], ["13.1", "Resilience to climate-related hazards"]];
  sdg.forEach(([n, t], i) => {
    const x = 0.8 + i * 3.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.05, w: 3.7, h: 1.5, rectRadius: 0.12, fill: { color: HEX.accent1, transparency: 80 }, line: { color: HEX.accent1, width: 1 }, objectName: "SDG card " + n });
    s.addText("Target " + n, { x: x + 0.25, y: 3.2, w: 3.2, h: 0.5, fontSize: 24, bold: true, fontFace: "Cambria", color: C.background1, margin: 0, isTextBox: true, objectName: "SDG number " + n });
    s.addText(t, { x: x + 0.25, y: 3.75, w: 3.2, h: 0.7, fontSize: 14, color: C.background2, valign: "top", margin: 0, isTextBox: true, objectName: "SDG text " + n });
  });
  s.addText("What I am asking for", { x: 0.8, y: 4.95, w: 11.7, h: 0.4, fontSize: 16, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "Ask heading" });
  s.addText(bullets(["A pilot zone and a drain asset register to test on real data", "Feedback from the field and operations teams who would use it", "A conversation about where this fits in your organisation"]), { x: 0.8, y: 5.4, w: 11.7, h: 1.4, fontSize: 18, color: C.background1, valign: "top", margin: 0, isTextBox: true, objectName: "Ask list" });
  s.addNotes("Close with the ask. Add your name and contact details here before presenting. Thank the audience and invite questions.");

  await pres.writeFile({ fileName: path.join(ROOT, "DrainProof_AI_Pitch.pptx") });
  await applyTheme(path.join(ROOT, "DrainProof_AI_Pitch.pptx"), THEME);
  console.log("deck written");
}
main().catch((e) => { console.error(e); process.exit(1); });
