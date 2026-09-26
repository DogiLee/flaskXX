"""Package source diffs and evidence without reading production data."""
from pathlib import Path
import difflib, hashlib, json
root=Path(__file__).resolve().parents[2]
before=root/'.investigation/audit-fixes-before'
out=root/'outputs/audit-fixes-20260922'
paths=[root/n for n in ['app.py','depo.py','excel_araclari.py']]
paths+=sorted((root/'templates').glob('*.html'))+sorted((root/'static').glob('*'))+sorted((root/'tests').glob('*.py'))+sorted((root/'tests').glob('*.cjs'))
diff=[];changed=[]
for path in paths:
    if not path.is_file(): continue
    relative=path.relative_to(root);old=before/relative
    current=path.read_text(encoding='utf-8');previous=old.read_text(encoding='utf-8') if old.exists() else ''
    if current==previous:continue
    changed.append(str(relative))
    diff.extend(difflib.unified_diff(previous.splitlines(True),current.splitlines(True),fromfile='before/'+relative.as_posix(),tofile='after/'+relative.as_posix()))
(out/'changes.diff').write_text(''.join(diff),encoding='utf-8')
(out/'source-hashes.json').write_text(json.dumps({str(p.relative_to(root)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths if p.is_file()},indent=2),encoding='utf-8')
results=json.loads((out/'results.json').read_text(encoding='utf-8'))
assert results['tests']==24 and not results['failures'] and not results['errors']
print(json.dumps({'changed_files':changed,'audit_tests':24,'audit_failures':0},ensure_ascii=False,indent=2))
