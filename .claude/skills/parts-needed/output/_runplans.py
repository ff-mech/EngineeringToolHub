import json, subprocess, os
os.chdir(r"\\npsvr05\FOXFAB_REDIRECT$\lbadong\Desktop\Engineering Tool Hub\EngineeringToolHub")
for J in ['J15423','J15803','J16257','J16482','J16498']:
    r = subprocess.run(
        ['python','.claude/skills/parts-needed/scripts/modelcopy.py','--job',J,
         '--pns-json',f'.claude/skills/parts-needed/output/{J}/_rows.json','--plan'],
        capture_output=True, text=True)
    d = json.loads(r.stdout)
    print(f'=== {J} ===')
    counts = {}
    for row in d['report']:
        counts[row['status']] = counts.get(row['status'],0)+1
    print('  counts:', counts)
    for row in d['report']:
        if row['status'] in ('ambiguous','none'):
            cands = [os.path.basename(c) for c in row.get('candidates',[])]
            print(f"  {row['status']:10} {row['pn'][:32]:32} -> {cands[:3]}")
