"""Complementary isolated route evidence for the frontend fetch contract."""
import json
from audit import Audit, OUT

case=Audit()
case.setUp()
try:
    client=case.mod.app.test_client()
    response=client.post('/api/not',json={'kart_id':1,'not':'synthetic'},follow_redirects=True)
    report={'unauthenticated_note_post':{
        'history':[{'status':r.status_code,'location':r.location} for r in response.history],
        'final_status':response.status_code,
        'content_type':response.content_type,
        'writes_expected':False,
    }}
    assert response.status_code == 401 and response.is_json and not response.history
    (OUT/'route-probes.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
finally:
    case.doCleanups()
