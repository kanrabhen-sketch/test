const AT = require('fs').readFileSync('.at','utf8').trim();
const DB = 'ai-studio-2ebc80c3-0c85-4e16-bcac-c4329e982059';
const SRC = `https://firestore.googleapis.com/v1/projects/radiant-badge-xfs6l/databases/${DB}/documents`;
const DST = `https://firestore.googleapis.com/v1/projects/foobierp-vy1v5/databases/${DB}/documents`;
const H = { 'Authorization': `Bearer ${AT}`, 'Content-Type': 'application/json' };
const COLS = ['activity_logs','clients','projects','resources','tasks','teams','users'];

async function listAll(base, col) {
  let out = [], token = '';
  do {
    const r = await fetch(`${base}/${col}?pageSize=300${token?`&pageToken=${token}`:''}`, { headers: H });
    const j = await r.json();
    if (j.error) throw new Error(`${col}: ${j.error.message}`);
    (j.documents||[]).forEach(d => out.push(d));
    token = j.nextPageToken || '';
  } while (token);
  return out;
}

(async () => {
  let total = 0, subFound = [];
  for (const col of COLS) {
    const docs = await listAll(SRC, col);
    for (const d of docs) {
      const id = d.name.split('/').pop();
      const r = await fetch(`${DST}/${col}?documentId=${encodeURIComponent(id)}`, {
        method: 'POST', headers: H, body: JSON.stringify({ fields: d.fields || {} })
      });
      const j = await r.json();
      if (j.error) { console.log(`  FAIL ${col}/${id}: ${j.error.message}`); continue; }
      total++;
      // 하위 컬렉션 확인
      const sr = await fetch(`${SRC}/${col}/${encodeURIComponent(id)}:listCollectionIds`, {
        method: 'POST', headers: H, body: '{}' });
      const sj = await sr.json();
      if (sj.collectionIds && sj.collectionIds.length) subFound.push(`${col}/${id} -> ${sj.collectionIds.join(',')}`);
    }
    console.log(`${col.padEnd(15)} ${docs.length} docs copied`);
  }
  console.log(`\n총 ${total}건 복사 완료`);
  console.log(subFound.length ? `하위 컬렉션 발견:\n  ${subFound.join('\n  ')}` : '하위 컬렉션 없음');
})();
