#!/usr/bin/env python3
"""
make_form_script.py: generate a ready-to-paste Apps Script (.gs) from content.json.

Usage:
    python3 make_form_script.py content.json out/

Writes out/<slug>-form-builder.gs: assets/form-builder-template.gs with everything between
the SPEC BEGIN and SPEC END markers replaced by a spec generated from content.json. The
template's own validation runs again inside Google before anything is created.

Runs check_cluster.py first and refuses to generate while it reports errors.
"""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "form-builder-template.gs")


def js(v):
    """JSON is valid JavaScript; keep Unicode readable."""
    return json.dumps(v, ensure_ascii=False, indent=2)


def fallback_text(block):
    lines = []
    if block.get("kind") == "table":
        cols = block["columns"]
        for r in block["rows"]:
            lines.append(r[0] + ": " + "; ".join(f"{cols[i]} {r[i]}" for i in range(1, len(r))))
    else:
        lines += block.get("text", [])
    if block.get("note"):
        lines.append(block["note"])
    if block.get("source"):
        lines.append("Source: " + block["source"])
    return "\n".join(lines)


def build_spec(spec):
    m = spec["meta"]
    form = m.get("form") or {}
    feats = {f["id"]: f for f in spec["features"]}
    flabel = lambda fid: f"{fid} ({feats[fid]['ref']}): {feats[fid]['text']}"
    duration = f"Suggested time: {m.get('duration_min', 40)} minutes."
    mode = m.get("distinction_mode", "probe")

    config = {
        "pe": m["standard"],
        "formTitle": f"{m['title']} · {m['standard']}",
        "sheetTitle": f"{m['title']} · Scoring ({m['standard']})",
        "durationNote": duration,
        "collectEmails": bool(form.get("collect_emails", True)),
        "teachers": form.get("teachers", []),
        "accessCode": form.get("access_code", ""),
        "accessCodeHelp": form.get("access_code_help", "Enter the code your teacher gives you at the start of the period."),
        "nameHelp": form.get("name_help", "First and last name, as it appears on your class list."),
        "images": {b["key"]: b["image"] for b in spec["stimulus"]["blocks"] if b.get("image")},
    }
    context = "\n\n".join(spec["stimulus"]["context"]) + f"\n\nResources: {m.get('resources', '')} {duration}"
    glossary = "\n".join(f"{g['term']}: {g['def']}" for g in spec["stimulus"].get("glossary", []))
    stimulus = [{"key": b["key"], "title": b["title"], "fallback": fallback_text(b)} for b in spec["stimulus"]["blocks"]]
    features = [flabel(f["id"]) for f in spec["features"] if f.get("in_scope")]

    crits = spec["criteria"]
    buckets = []
    for b in spec["buckets"]:
        es = [c["id"] for c in crits if not c.get("beyond") and b["name"] in c["buckets"]]
        probe = next((c["id"] for c in crits if c.get("beyond") and b["name"] in c["buckets"]), None)
        entry = {"name": b["name"], "label": b["label"], "crits": es, "reflection": b["reflection"]}
        if probe:
            entry["mwdCrit"] = probe
        buckets.append(entry)

    criteria = []
    for c in sorted(crits, key=lambda x: x["order"]):
        e = {
            "crit": c["id"], "buckets": c["buckets"], "skill": c["skill"], "order": c["order"],
            "type": "mc" if c["type"] == "mc" else "para",
            "title": f"{c['item']}. {c['stem']}", "stem": "", "help": c.get("help", ""),
            "part": c["part"], "criterion": c["criterion"], "earns": c["earns"],
        }
        if c.get("beyond"):
            e["beyond"] = True
            e["extends"] = flabel(c["extends"])
        else:
            e["feature"] = flabel(c["feature"])
        if c["type"] == "mc":
            e["choices"] = [[o["text"], bool(o.get("key")), "" if o.get("key") else "misconception: " + o.get("misconception", "")]
                            for o in c["choices"]]
        criteria.append(e)

    parts = [{"n": p["n"], "title": p["title"], "help": p.get("help", "")} for p in spec["parts"]]

    return f"""/* >>> SPEC BEGIN (generated from content.json by make_form_script.py; do not edit by hand) */

var CONFIG = {js(config)};

var CONTEXT_BLURB = {js(context)};

var GLOSSARY = {js(glossary)};

var STIMULUS = {js(stimulus)};

var FEATURES = {js(features)};

var BUCKETS = {js(buckets)};

var MWD_REQUIRES_TOP_GATE = {js(bool(m.get("mwd_requires_top_gate", False)))};

var DISTINCTION_MODE = {js(mode)};

var CRITERIA = {js(criteria)};

var PARTS = {js(parts)};

var MIN_ITEMS_PER_BUCKET = {int(m.get("min_items_per_bucket", 3))};

/* Scoring sheet layout: Timestamp, Name, Block, Teacher occupy columns 1-4. */
var FIRST_CRIT_COL_GLOBAL = 5;
/* <<< SPEC END */"""


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    src, out = sys.argv[1], sys.argv[2]
    chk = subprocess.run([sys.executable, os.path.join(HERE, "check_cluster.py"), src, "--json"],
                         capture_output=True, text=True)
    result = json.loads(chk.stdout)
    if result["errors"]:
        print("Checker reports errors. Fix them first:\n  - " + "\n  - ".join(result["errors"]))
        sys.exit(1)
    spec = json.load(open(src, encoding="utf-8"))
    form = spec["meta"].get("form") or {}
    if not form.get("teachers") or any(re.fullmatch(r"<.*>", t) for t in form["teachers"]) \
            or not form.get("access_code") or re.fullmatch(r"<.*>", form.get("access_code", "")):
        print("WARNING: meta.form.teachers or access_code is still a placeholder. "
              "The script will refuse to run in Google until they are set.")
    tpl = open(TEMPLATE, encoding="utf-8").read()
    new, n = re.subn(r"/\* >>> SPEC BEGIN.*?/\* <<< SPEC END \*/", lambda _: build_spec(spec), tpl, flags=re.S)
    if n != 1:
        sys.exit("Template markers not found exactly once.")
    os.makedirs(out, exist_ok=True)
    slug = spec["meta"].get("slug") or spec["meta"]["standard"].lower()
    path = os.path.join(out, f"{slug}-form-builder.gs")
    open(path, "w", encoding="utf-8").write(new)
    print("wrote " + path)
    imgs = [b["image"] for b in spec["stimulus"]["blocks"] if b.get("image")]
    if imgs:
        print("Upload to Drive before running: " + ", ".join(imgs))


if __name__ == "__main__":
    main()
