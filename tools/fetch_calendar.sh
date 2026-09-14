#!/bin/bash
# 나이스 학사일정 → v2/calendar.json 갱신
#
# 언제 돌리나: 학년도가 바뀔 때, 또는 학교가 재량휴업일을 새로 잡았을 때.
# 인증키는 ~/Top_Class/.neis_key 에만 둔다(.gitignore 처리됨).
# ⚠️ 생성되는 calendar.json 에는 키가 들어가지 않는다 — 저장소가 공개라서다.
set -e
cd "$(dirname "$0")/.."
KEY=$(cat .neis_key)
FROM=${1:-20260301}
TO=${2:-20270228}
curl -s "https://open.neis.go.kr/hub/SchoolSchedule?KEY=${KEY}&Type=json&pSize=1000&ATPT_OFCDC_SC_CODE=N10&SD_SCHUL_CODE=8140267&AA_FROM_YMD=${FROM}&AA_TO_YMD=${TO}" -o /tmp/neis.json
python3 - "$FROM" "$TO" <<'PY'
import json, io, datetime, sys
j=json.load(io.open('/tmp/neis.json',encoding='utf-8'))
if 'SchoolSchedule' not in j:
    raise SystemExit('나이스 응답 오류: '+json.dumps(j,ensure_ascii=False))
rows=j['SchoolSchedule'][1]['row']
off={}
for r in rows:
    if r['SBTR_DD_SC_NM']=='해당없음': continue
    d=r['AA_YMD']; nm=r['EVENT_NM'].strip()
    if nm=='토요휴업일': continue
    if datetime.date(int(d[:4]),int(d[4:6]),int(d[6:])).weekday()>=5: continue
    off.setdefault(d[:4]+'-'+d[4:6]+'-'+d[6:], nm)
out={"_주의":"NEIS 학사일정에서 생성. 인증키는 여기 들어있지 않다. 갱신: tools/fetch_calendar.sh",
 "학교":"청양고등학교","교육청코드":"N10","학교코드":"8140267","학년도":"2026",
 "기간":sys.argv[1]+'~'+sys.argv[2],"받은날":datetime.date.today().isoformat(),
 "휴업일":dict(sorted(off.items()))}
io.open('v2/calendar.json','w',encoding='utf-8').write(json.dumps(out,ensure_ascii=False,indent=1))
print('평일 휴업일 %d일 → v2/calendar.json' % len(off))
PY
