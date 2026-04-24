// DELETE /api/reports/:id — 개별 보고 기록 삭제

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}

export async function onRequestDelete(context) {
  const db = context.env.DB;
  if (!db) return json({ error: 'D1 DB가 바인딩되지 않았습니다.' }, 500);
  const id = Number(context.params.id);
  if (!Number.isFinite(id)) return json({ error: 'invalid id' }, 400);
  try {
    await db.prepare(`DELETE FROM reports WHERE id = ?`).bind(id).run();
    return json({ ok: true });
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}
