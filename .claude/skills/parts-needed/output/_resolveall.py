"""Resolve all 5 test jobs at once: for each ambiguous P/N, pick the
candidate whose basename matches the user's choice; then run --resolve."""
import json, os, subprocess, sys
os.chdir(r"\\npsvr05\FOXFAB_REDIRECT$\lbadong\Desktop\Engineering Tool Hub\EngineeringToolHub")

# User's choices keyed by normalized P/N → desired basename (or IGNORE/SKIP)
BY_BASENAME = {
    # 2610 receptacle → 2610.SLDPRT  (appears in J15803, J16257, J16482)
    "2610":               "2610.SLDPRT",
    # Leviton mini flanged inlet
    "MLTPB":              "Leviton ML2-PB.SLDPRT",
    # SOCOMEC lug kit
    "39543020":           "39543020.SLDASM",
}
# Direct SKIPs (no persistent learning)
SKIP = {
    "EXA005036":  "SKIP",
}

JOBS = ['J15423','J15803','J16257','J16482','J16498']

def norm(s):
    import re
    return re.sub(r"[\s\-_./\\]+","", (s or "").upper())

for J in JOBS:
    rows_path = f".claude/skills/parts-needed/output/{J}/_rows.json"
    plan_out = subprocess.run(
        ['python','.claude/skills/parts-needed/scripts/modelcopy.py',
         '--job',J,'--pns-json',rows_path,'--plan'],
        capture_output=True, text=True).stdout
    d = json.loads(plan_out)
    choices = {}
    for row in d['report']:
        if row['status'] not in ('ambiguous','none'):
            continue
        npn = row['npn']
        # Try basename-keyed choice
        if npn in BY_BASENAME:
            want = BY_BASENAME[npn]
            match = next((c for c in row['candidates']
                          if os.path.basename(c) == want), None)
            if match:
                choices[npn] = match
                continue
        if npn in SKIP:
            choices[npn] = SKIP[npn]
            continue
        # Unhandled → SKIP this job (shouldn't happen after our clauses)
        choices[npn] = "SKIP"
        print(f"  [warn] {J} unhandled: {row['pn']}")
    cpath = f".claude/skills/parts-needed/output/{J}/_choices.json"
    with open(cpath,"w",encoding="utf-8") as f:
        json.dump(choices, f, indent=2)
    # Resolve
    res_out = subprocess.run(
        ['python','.claude/skills/parts-needed/scripts/modelcopy.py',
         '--job',J,'--pns-json',rows_path,'--resolve',cpath],
        capture_output=True, text=True).stdout
    r = json.loads(res_out)
    counts = {}
    for s in r['summary']:
        counts[s['status']] = counts.get(s['status'],0)+1
    copied = sum(1 for s in r['summary']
                  for f in s.get('files',[]) if f.get('state')=='copied')
    print(f"{J}: dest={r['dest'].split(os.sep)[-1]}  copied={copied}  {counts}")
