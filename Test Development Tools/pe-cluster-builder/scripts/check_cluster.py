#!/usr/bin/env python3
"""
check_cluster.py: structural, item, data and text-hygiene checks for a PE cluster spec.

Usage:
    python3 check_cluster.py content.json          # human-readable report
    python3 check_cluster.py content.json --json   # machine-readable, used by build_docs.js

Exit status is 1 if any error is found, 0 otherwise. Warnings never fail the run; they belong
in the draft's open decisions unless fixed.
"""
import json
import math
import re
import sys
from collections import Counter, defaultdict

LETTERS = "ABCDEFGH"
DEFAULT_PROPS = {"developing": 0.45, "meeting": 0.70, "top": 0.85}
RESIDUE = [r"\bupdated\b", r"\bnew this year\b", r"\brevised version\b", r"\(revised\)",
           r"\bv\d\b", r"\bdraft\b", r"\bTODO\b", r"\bTBD\b"]
FORMULA = re.compile(r"\b(?:[A-Z][a-z]?\d*)*[A-Z][a-z]?\d+(?:[A-Z][a-z]?\d*)*\b")
LATEX = re.compile(r"\$|\\frac|\\text|\^\{|_\{|\\times|\\cdot")


def ceil_prop(n, p):
    # integer arithmetic on percentages avoids float drift (0.7 * 10 etc.)
    return -(-n * round(p * 100) // 100)


def gates_for(n, props):
    return {
        "n": n,
        "emerging": 1 if n else 0,
        "developing": ceil_prop(n, props["developing"]),
        "meeting": ceil_prop(n, props["meeting"]),
        "top": max(ceil_prop(n, props["top"]), ceil_prop(n, props["meeting"])),
    }


def words(s):
    return len([w for w in re.split(r"\s+", s.strip()) if w])


def sentences(s):
    return [x for x in re.split(r"(?<=[.?!])\s+|(?<=[.?!][”’)])\s+", s.strip()) if x]


def safe_eval(expr):
    if not re.fullmatch(r"[0-9.\s+\-*/()]+", expr):
        raise ValueError("expr may contain numbers and + - * / ( ) only")
    return eval(expr, {"__builtins__": {}}, {})


def check(spec):
    E, W = [], []
    meta = spec.get("meta", {})
    props = {**DEFAULT_PROPS, **(meta.get("gate_proportions") or {})}
    mode = meta.get("distinction_mode", "probe")
    floor = int(meta.get("min_items_per_bucket", 3))
    cap_words = int(meta.get("sentence_cap_words", 25))

    # ---------------- meta ----------------
    std = meta.get("standard", "")
    if not std:
        E.append("meta.standard is empty.")
    elif re.search(r"[,;&]|\band\b", std):
        E.append(f"meta.standard looks like more than one standard: {std!r}. A cluster assesses exactly one.")
    if mode not in ("probe", "perfect", "none"):
        E.append(f"meta.distinction_mode must be probe, perfect or none, not {mode!r}.")
    if meta.get("delivery") not in ("paper", "form"):
        E.append("meta.delivery must be 'paper' or 'form'.")
    if meta.get("delivery") == "form":
        f = meta.get("form") or {}
        t = f.get("teachers") or []
        if not t or any(re.fullmatch(r"<.*>", x or "") for x in t):
            E.append("Form route: meta.form.teachers is empty or still a placeholder.")
        if not f.get("access_code") or re.fullmatch(r"<.*>", f.get("access_code", "")):
            E.append("Form route: meta.form.access_code is unset or still a placeholder.")

    # ---------------- scope and features ----------------
    scope_codes = [s.get("code") for s in spec.get("scope", [])]
    for code, n in Counter(scope_codes).items():
        if n > 1:
            E.append(f"Scope code {code} appears {n} times.")
    scope_set = set(scope_codes)

    feats = {f["id"]: f for f in spec.get("features", [])}
    if len(feats) != len(spec.get("features", [])):
        E.append("Feature IDs are not unique.")
    for f in feats.values():
        if f.get("in_scope"):
            if f.get("skill") not in scope_set:
                E.append(f"{f['id']} is in scope but its skill code {f.get('skill')!r} is not on the scope list.")
        elif not f.get("exclusion_reason"):
            E.append(f"{f['id']} is out of scope but gives no exclusion_reason.")
    in_scope = {k for k, f in feats.items() if f.get("in_scope")}

    # ---------------- buckets ----------------
    buckets = [b["name"] for b in spec.get("buckets", [])]
    for name, n in Counter(buckets).items():
        if n > 1:
            E.append(f"Bucket {name} is declared {n} times.")
    for b in spec.get("buckets", []):
        if not b.get("reflection"):
            E.append(f"Bucket {b['name']} has no grading reflection.")

    # ---------------- criteria ----------------
    crits = spec.get("criteria", [])
    parts = {p["n"] for p in spec.get("parts", [])}
    for key in ("id", "item", "order"):
        for v, n in Counter(c.get(key) for c in crits).items():
            if n > 1:
                E.append(f"Duplicate criterion {key}: {v}.")

    by_item = {c["item"]: c for c in crits}
    es = [c for c in crits if not c.get("beyond")]
    probes = [c for c in crits if c.get("beyond")]

    for c in crits:
        cid = c.get("id", "?")
        for req in ("item", "order", "part", "type", "buckets", "skill", "stem", "criterion", "earns", "cue"):
            if c.get(req) in (None, "", []):
                E.append(f"{cid} is missing {req}.")
        if c.get("part") not in parts:
            E.append(f"{cid} is in part {c.get('part')}, which is not declared in parts.")
        if c.get("type") not in ("mc", "cr"):
            E.append(f"{cid} type must be 'mc' or 'cr'.")
        for t in c.get("buckets", []):
            if t not in buckets:
                E.append(f"{cid} is tagged {t}, which is not a bucket.")
        if len(c.get("buckets", [])) >= 3:
            W.append(f"{cid} is tagged to {len(c['buckets'])} buckets. List it in open decisions for approval.")
        if c.get("skill") and c["skill"] not in scope_set:
            E.append(f"{cid} skill code {c['skill']} is not on the scope list.")

        # anchoring
        if c.get("beyond"):
            if c.get("feature"):
                E.append(f"{cid} is a probe and must not also claim a feature.")
            if c.get("extends") not in in_scope:
                E.append(f"{cid} is a probe but extends {c.get('extends')!r}, which is not an in-scope feature.")
        else:
            if not c.get("feature"):
                E.append(f"{cid} is anchored to nothing. Name a feature or declare it a probe.")
            elif c["feature"] not in feats:
                E.append(f"{cid} cites feature {c['feature']}, which does not exist.")
            elif c["feature"] not in in_scope:
                E.append(f"{cid} assesses {c['feature']}, which is out of scope.")

        if c.get("compound") and not c.get("criterion", "").rstrip().endswith("Both required."):
            E.append(f"{cid} is compound; its criterion text must end with \"Both required.\"")
        if c.get("compound"):
            W.append(f"{cid} is a compound criterion. List it in open decisions.")
        if c.get("or_pathway"):
            W.append(f"{cid} is an OR-pathway criterion. List it in open decisions.")
        if len(c.get("buckets", [])) >= 2 and not c.get("beyond"):
            W.append(f"{cid} is cross-scored to {', '.join(c['buckets'])}. Confirm its criterion names what each dimension requires.")

        # constructed response
        if c.get("type") == "cr":
            cap = c.get("sentence_cap")
            if not cap:
                E.append(f"{cid} is constructed response with no sentence_cap.")
            else:
                need = math.ceil(cap * 18 / 12) + 1
                if meta.get("delivery") == "paper" and (c.get("lines") or 0) < need:
                    W.append(f"{cid} has {c.get('lines')} answer lines; a {cap}-sentence cap calibrates to at least {need}.")

        # selected response
        if c.get("type") == "mc":
            ch = c.get("choices") or []
            if len(ch) != 4:
                W.append(f"{cid} has {len(ch)} options; four is the default.")
            keys = [i for i, o in enumerate(ch) if o.get("key")]
            if len(keys) != 1:
                E.append(f"{cid} has {len(keys)} keyed options; must be exactly 1.")
            for i, o in enumerate(ch):
                if not o.get("key") and not o.get("misconception"):
                    E.append(f"{cid} option {LETTERS[i]} is a distractor with no named misconception.")
                if re.search(r"\b(all|none) of the above\b", o.get("text", ""), re.I):
                    E.append(f"{cid} option {LETTERS[i]} uses all/none of the above.")
            if ch and len(keys) == 1:
                lens = [words(o["text"]) for o in ch]
                if max(lens) - min(lens) > 8:
                    E.append(f"{cid} option lengths span {min(lens)} to {max(lens)} words. Match them.")
                dl = [l for i, l in enumerate(lens) if i != keys[0]]
                if lens[keys[0]] - (sum(dl) / len(dl)) > 3:
                    W.append(f"{cid} key is {lens[keys[0]]} words against a distractor mean of {sum(dl)/len(dl):.1f}. A visibly longer key leaks.")

    # ---------------- key distribution across the set ----------------
    mc = [c for c in crits if c.get("type") == "mc" and c.get("choices")]
    key_pos = {}
    longest_key = 0
    for c in mc:
        ks = [i for i, o in enumerate(c["choices"]) if o.get("key")]
        if len(ks) == 1:
            key_pos[c["id"]] = LETTERS[ks[0]]
            lens = [words(o["text"]) for o in c["choices"]]
            if lens[ks[0]] == max(lens) and lens.count(max(lens)) == 1:
                longest_key += 1
    hist = Counter(key_pos.values())
    if len(mc) >= 3 and hist:
        cap = max(2, math.ceil(len(mc) / 2))
        pos, n = hist.most_common(1)[0]
        if len(hist) == 1:
            E.append(f"Every selected-response key is option {pos}. Distribute the keys.")
        elif n > cap:
            E.append(f"Option {pos} is the key on {n} of {len(mc)} selected-response items; cap is {cap}.")
        if len(mc) >= 4 and len(hist) < 3:
            E.append(f"Keys occupy only {len(hist)} positions ({', '.join(sorted(hist))}). Use at least three.")
        if longest_key > len(mc) / 2:
            W.append(f"The key is the single longest option on {longest_key} of {len(mc)} items.")

    # ---------------- feature coverage ----------------
    cover = defaultdict(list)
    for c in es:
        if c.get("feature"):
            cover[c["feature"]].append(c["id"])
    for fid in sorted(in_scope, key=lambda x: int(re.sub(r"\D", "", x) or 0)):
        hits = cover.get(fid, [])
        if not hits:
            E.append(f"In-scope feature {fid} is not assessed.")
        elif len(hits) > 2:
            E.append(f"Feature {fid} is assessed {len(hits)} times ({', '.join(hits)}); two is the maximum.")
        elif len(hits) == 2:
            W.append(f"Feature {fid} is assessed twice ({', '.join(hits)}). Confirm the two exposures are different demands.")

    # ---------------- buckets: counts, exclusivity, probes ----------------
    counts, gates = {}, {}
    for b in buckets:
        mine = [c for c in es if b in c.get("buckets", [])]
        excl = [c for c in mine if len(c["buckets"]) == 1]
        counts[b] = {"es": len(mine), "exclusive": len(excl), "ids": [c["id"] for c in mine]}
        gates[b] = gates_for(len(mine), props)
        if not mine:
            E.append(f"Bucket {b} has no evidence-statement criteria. A mapped bucket must be scored.")
        elif len(mine) < floor:
            W.append(f"Bucket {b} has {len(mine)} criteria, below the floor of {floor}. Resolve by cross-scoring, extending a prompt, or a provisional-level note; never padding.")
        need = max(2, math.ceil(len(mine) / 2))
        if len(mine) >= 2 and len(excl) < need:
            E.append(f"Bucket {b} has {len(excl)} exclusive criteria; needs at least {need}.")

        bp = [c for c in probes if b in c.get("buckets", [])]
        if mode == "probe":
            if len(bp) != 1:
                E.append(f"Bucket {b} has {len(bp)} Distinction probes; mode 'probe' needs exactly one.")
            for p in bp:
                if len(p["buckets"]) != 1:
                    E.append(f"Probe {p['id']} is tagged to {len(p['buckets'])} buckets; probes are never cross-scored.")
                if p.get("type") != "cr":
                    W.append(f"Probe {p['id']} is selected response; probes are normally constructed response.")
                m = re.fullmatch(r"(\d+)b", str(p.get("item", "")))
                parent = by_item.get(f"{m.group(1)}a") if m else None
                if not parent:
                    E.append(f"Probe {p['id']} is item {p.get('item')}; it must be part (b) of an existing item (e.g. 4b after 4a).")
                elif b not in parent.get("buckets", []) or parent.get("beyond"):
                    E.append(f"Probe {p['id']} extends item {parent['item']}, which is not an evidence-statement item in {b}.")
                elif parent.get("order", 0) + 1 != p.get("order"):
                    W.append(f"Probe {p['id']} does not sit directly after item {parent['item']} in student order.")
                if (p.get("sentence_cap") or 0) > 2:
                    W.append(f"Probe {p['id']} allows {p['sentence_cap']} sentences; two is the design cap.")
    if mode != "probe" and probes:
        E.append(f"Distinction mode is {mode!r} but the spec contains probes ({', '.join(c['id'] for c in probes)}).")

    # ---------------- time ----------------
    if len(crits) > 13:
        W.append(f"{len(crits)} scored judgments exceeds the guide of about 13 for {meta.get('duration_min', 40)} minutes. Raise it at the draft; do not compress items.")

    # ---------------- text hygiene ----------------
    student_texts = []  # (where, text)
    stim = spec.get("stimulus", {})
    for i, p in enumerate(stim.get("context", [])):
        student_texts.append((f"context[{i}]", p))
    for b in stim.get("blocks", []):
        student_texts.append((f"{b.get('key')}.title", b.get("title", "")))
        student_texts.append((f"{b.get('key')}.note", b.get("note", "")))
        for r in b.get("rows", []):
            student_texts.append((f"{b.get('key')}.row", " | ".join(r)))
        for c in b.get("columns", []):
            student_texts.append((f"{b.get('key')}.col", c))
        for t in b.get("text", []):
            student_texts.append((f"{b.get('key')}.text", t))
    for g in stim.get("glossary", []):
        student_texts.append(("glossary", f"{g.get('term')}: {g.get('def')}"))
    for p in spec.get("parts", []):
        student_texts.append((f"part{p['n']}", p.get("title", "") + " " + p.get("help", "")))
    for c in crits:
        student_texts.append((f"{c['id']}.stem", c.get("stem", "")))
        student_texts.append((f"{c['id']}.help", c.get("help", "")))
        for i, o in enumerate(c.get("choices") or []):
            student_texts.append((f"{c['id']}.{LETTERS[i]}", o.get("text", "")))
    teacher_texts = []
    for c in crits:
        for k in ("criterion", "earns", "cue"):
            teacher_texts.append((f"{c['id']}.{k}", c.get(k, "")))
        for i, o in enumerate(c.get("choices") or []):
            teacher_texts.append((f"{c['id']}.{LETTERS[i]}.misc", o.get("misconception", "")))
    for b in spec.get("buckets", []):
        teacher_texts.append((f"{b['name']}.reflection", b.get("reflection", "")))

    held = [h["term"] for h in spec.get("held_back_terms", [])]
    for where, t in student_texts + teacher_texts:
        if not t:
            continue
        if "\u2014" in t:
            E.append(f"Em-dash in {where}.")
        if '"' in t or re.search(r"(?<=\w)'(?=\w)|(?<!\w)'|'(?!\w)", t):
            E.append(f"Straight quote or apostrophe in {where}. Use curly quotes.")
        if LATEX.search(t):
            E.append(f"LaTeX or dollar-sign notation in {where}.")
        for term in held:
            if re.search(r"(?<!\w)" + re.escape(term) + r"(?!\w)", t, re.I):
                E.append(f"Held-back term {term!r} appears in {where}.")
    for where, t in student_texts:
        for pat in RESIDUE:
            if re.search(pat, t, re.I):
                E.append(f"Revision residue ({pat}) in {where}.")
        for m in FORMULA.findall(t):
            W.append(f"{where}: {m!r} may be a formula missing Unicode subscripts.")
        if ".stem" in where or re.search(r"\.[A-D]$", where):
            for s in sentences(t):
                if words(s) > cap_words:
                    W.append(f"{where} has a {words(s)}-word sentence; cap is {cap_words}.")

    # ---------------- data ----------------
    nums = []
    for b in stim.get("blocks", []):
        for r in b.get("rows", []):
            for cell in r[1:]:
                nums += re.findall(r"-?\d+(?:\.\d+)?", cell)
    dup = [n for n, k in Counter(nums).items() if k > 1]
    if dup:
        W.append(f"Stimulus tables repeat the values {', '.join(sorted(dup))}. Confirm no two different quantities share a value by accident.")
    all_text = " ".join(t for _, t in student_texts) + " " + " ".join(t for _, t in teacher_texts)
    for d in spec.get("derived_checks", []):
        try:
            got = safe_eval(d["expr"])
        except Exception as ex:
            E.append(f"derived_check {d.get('label')}: {ex}")
            continue
        exp = float(d["expected"])
        if abs(got - exp) > float(d.get("tolerance", 1e-6)):
            E.append(f"derived_check {d.get('label')}: {d['expr']} = {got:g}, not {d['expected']}.")
        if str(d["expected"]) not in all_text:
            W.append(f"derived_check {d.get('label')}: expected value {d['expected']} does not appear in the stimulus or keys.")
    constructed = [b for b in stim.get("blocks", []) if b.get("constructed")]
    for b in constructed:
        if "illustrative" not in (b.get("note", "") + b.get("source", "")).lower():
            E.append(f"Stimulus block {b.get('key')} is constructed but not labelled illustrative.")
    for b in stim.get("blocks", []):
        if not b.get("constructed") and not b.get("source"):
            E.append(f"Stimulus block {b.get('key')} is real data or text with no source.")

    return {
        "errors": E,
        "warnings": W,
        "counts": counts,
        "gates": gates,
        "key_positions": dict(sorted(hist.items())),
        "key_by_criterion": key_pos,
        "coverage": {k: cover.get(k, []) for k in feats},
        "n_criteria": len(crits),
        "n_es": len(es),
        "n_probes": len(probes),
        "mode": mode,
        "props": props,
    }


def report(r):
    out = []
    out.append(f"Criteria: {r['n_criteria']} ({r['n_es']} evidence-statement, {r['n_probes']} probes). Distinction mode: {r['mode']}.")
    out.append("Buckets:")
    for b, c in r["counts"].items():
        g = r["gates"][b]
        out.append(f"  {b}: {c['es']} ES criteria ({c['exclusive']} exclusive) [{', '.join(c['ids'])}]"
                   f"  gates E {g['emerging']}+ / D {g['developing']}+ / M {g['meeting']}+ / top {g['top']}+")
    if r["key_positions"]:
        out.append("Key positions: " + "  ".join(f"{k}={v}" for k, v in r["key_positions"].items()))
    out.append("Feature coverage:")
    for f, ids in r["coverage"].items():
        out.append(f"  {f}: {', '.join(ids) if ids else '-'}")
    out.append(f"\nERRORS ({len(r['errors'])})")
    out += [f"  - {e}" for e in r["errors"]] or ["  none"]
    out.append(f"\nWARNINGS ({len(r['warnings'])})")
    out += [f"  - {w}" for w in r["warnings"]] or ["  none"]
    return "\n".join(out)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    with open(sys.argv[1], encoding="utf-8") as fh:
        spec = json.load(fh)
    r = check(spec)
    if "--json" in sys.argv:
        print(json.dumps(r, ensure_ascii=False, indent=1))
    else:
        print(report(r))
    sys.exit(1 if r["errors"] else 0)


if __name__ == "__main__":
    main()
