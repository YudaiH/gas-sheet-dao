/**
 * SheetDAOのインスタンスを作成します。
 * 外部プロジェクトからライブラリとして使用する場合、この関数を経由して呼び出します。
 * @param {string} sheetId スプレッドシートのID
 * @param {string} sheetName 対象のシート名
 * @return {SheetDAO} SheetDAOインスタンス
 */
function create(sheetId, sheetName) {
  return new SheetDAO(sheetId, sheetName);
}
