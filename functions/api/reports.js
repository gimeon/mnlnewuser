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

const MAX_ITEMS = 10;

async function ensureTable(db) {
  await db.prepare(DDL).run();
}

async function trimOldRows(db) {
  // 최신 MAX_ITEMS 개만 남기고 나머지 삭제
  await db.prepare(
    `DELETE FROM reports WHERE id NOT IN (
       SELECT id FROM reports ORDER BY created_at DESC, id DESC LIMIT ?
     )`
  ).bind(MAX_ITEMS).run();
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}

// HISTORY_PASSCODE 환경변수와 X-Passcode 헤더를 비교. 일치하면 null, 아니면 401 응답 반환.
function requirePasscode(context) {
  const expected = context.env.HISTORY_PASSCODE;
  if (!expected) {
    // 환경변수 미설정이면 사실상 잠금 해제된 상태 — 운영 실수 방지용 명시 로그
    console.warn('[auth] HISTORY_PASSCODE env var 미설정. 모든 요청 통과.');
    return null;
  }
  const got = context.request.headers.get('X-Passcode');
  if (got !== expected) {
    return json({ error: 'unauthorized' }, 401);
  }
  return null;
}

export async function onRequestGet(context) {
  const denied = requirePasscode(context);
  if (denied) return denied;
  const db = context.env.DB;
  if (!db) return json({ error: 'D1 DB가 바인딩되지 않았습니다 (Pages Settings → Bindings에서 DB로 연결).' }, 500);
  try {
    await ensureTable(db);
    const { results } = await db
      .prepare(`SELECT id, data FROM reports ORDER BY created_at DESC, id DESC LIMIT ?`)
      .bind(MAX_ITEMS)
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
    await trimOldRows(db);
    return json({ ok: true });
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}

export async function onRequestDelete(context) {
  const denied = requirePasscode(context);
  if (denied) return denied;
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
