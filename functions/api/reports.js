// Cloudflare Pages Function — 보고 기록 공유 저장소 API
// 바인딩 이름: DB (Cloudflare 대시보드에서 D1 → 이 Pages 프로젝트에 DB로 바인딩)
// 엔드포인트:
//   GET    /api/reports     — 최신 200건 조회
//   POST   /api/reports     — 신규 저장
//   DELETE /api/reports     — 전체 삭제

const DDL = `
  CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data TEXT NOT NULL,
    created_at TEXT NOT NULL
  );
`;

async function ensureTable(db) {
  await db.prepare(DDL).run();
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}

export async function onRequestGet(context) {
  const db = context.env.DB;
  if (!db) return json({ error: 'D1 DB가 바인딩되지 않았습니다 (Pages Settings → Bindings에서 DB로 연결).' }, 500);
  try {
    await ensureTable(db);
    const { results } = await db
      .prepare(`SELECT id, data FROM reports ORDER BY created_at DESC, id DESC LIMIT 200`)
      .all();
    const items = results.map((r) => ({ id: r.id, ...JSON.parse(r.data) }));
    return json(items);
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}

export async function onRequestPost(context) {
  const db = context.env.DB;
  if (!db) return json({ error: 'D1 DB가 바인딩되지 않았습니다.' }, 500);
  try {
    const body = await context.request.json();
    if (!body || !body.author || !body.startAt || !body.endAt || !body.reportText) {
      return json({ error: '필수 필드 누락' }, 400);
    }
    body.savedAt = body.savedAt || new Date().toISOString();
    body.createdAt = body.createdAt || body.endAt;
    await ensureTable(db);
    await db
      .prepare(`INSERT INTO reports (data, created_at) VALUES (?, ?)`)
      .bind(JSON.stringify(body), body.createdAt)
      .run();
    return json({ ok: true });
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}

export async function onRequestDelete(context) {
  const db = context.env.DB;
  if (!db) return json({ error: 'D1 DB가 바인딩되지 않았습니다.' }, 500);
  try {
    await ensureTable(db);
    await db.prepare(`DELETE FROM reports`).run();
    return json({ ok: true });
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}
