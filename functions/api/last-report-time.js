// GET /api/last-report-time — 비밀번호 없이도 마지막 보고의 createdAt만 반환
// 보고 본문/작성자/카운트는 노출하지 않음. "집계 시작점" 컨텍스트 제공용.

const DDL = `
  CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data TEXT NOT NULL,
    created_at TEXT NOT NULL
  );
`;

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}

export async function onRequestGet(context) {
  const db = context.env.DB;
  if (!db) return json({ error: 'D1 DB가 바인딩되지 않았습니다.' }, 500);
  try {
    await db.prepare(DDL).run();
    const { results } = await db
      .prepare(`SELECT created_at FROM reports ORDER BY created_at DESC, id DESC LIMIT 1`)
      .all();
    const lastCreatedAt = results.length > 0 ? results[0].created_at : null;
    return json({ lastCreatedAt });
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}
