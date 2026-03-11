
/** タスク追加 */
function api_addTask(task) {
  try {
    const sh = getTaskSheet_();
    const lastRow = sh.getLastRow();
    const nextRow = lastRow + 1;

    const taskName = String(task?.task || '').trim();
    const link = String(task?.link || '').trim();
    const dateVal = task?.date ? new Date(task.date) : '';

    if (!taskName) {
      return { ok: false, error: 'タスク名は必須です' };
    }

    const no = Math.max(1, lastRow); // 1行目ヘッダ前提

    sh.getRange(nextRow, 1).setValue(no);         // No
    sh.getRange(nextRow, 2).setValue(taskName);   // タスク
    sh.getRange(nextRow, 3).setValue(link);       // リンク
    sh.getRange(nextRow, 4).setValue(dateVal || ''); // 期日
    sh.getRange(nextRow, 5).setValue('未着手');   // ステータス
    sh.getRange(nextRow, 6).setValue(0);          // 進捗率
    sh.getRange(nextRow, 6).setNumberFormat('0%');

    const map = getTaskHeaderMap_(sh);
    if (map['進捗']) {
      const memberHeaders = getTaskMemberHeaders_(getTaskHeaders_(sh));
      sh.getRange(nextRow, map['進捗']).setValue(`0/${memberHeaders.length}`);
    }

    return {
      ok: true,
      no,
    };
  } catch (err) {
    return {
      ok: false,
      error: err.message,
    };
  }
}