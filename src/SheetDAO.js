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

/**
 * SheetDAO: Google Spread Sheet をDBのように操作するためのDAOクラス
 */
class SheetDAO {
  /**
   * コンストラクター
   * @param {string} sheetId スプレッドシートのID
   * @param {string} sheetName 対象のシート名
   */
  constructor(sheetId, sheetName) {
    const ss = SpreadsheetApp.openById(sheetId);
    this.sheet = ss.getSheetByName(sheetName);
    if (!this.sheet) {
      throw new Error(`Sheet "${sheetName}" not found in file.`);
    }

    this.lastCol = this.sheet.getLastColumn();
    if (this.lastCol === 0) {
      throw new Error(
        `Sheet "${sheetName}" is empty. Please define a header row.`,
      );
    }

    this.headers = this.sheet.getRange(1, 1, 1, this.lastCol).getValues()[0];
  }

  /**
   * すべてのレコードを取得します。
   * @returns {Object[]} オブジェクトの配列
   */
  listAll() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }

    const values = this.sheet
      .getRange(2, 1, lastRow - 1, this.lastCol)
      .getValues();
    return values.map((row) => this._toObject(row));
  }

  /**
   * 指定したIDのレコードを取得します。
   * @param {number|string} id
   * @returns {Object|null}
   */
  find(id) {
    const rowIndex = this._getRowIndexById(id);
    return rowIndex ? this._getObjectByRowIndex(rowIndex) : null;
  }

  /**
   * 新しいレコードを追加します。IDは自動採番されます。
   * @param {Object} data IDを除くデータオブジェクト { key: value, ... }
   * @returns {Object} 追加されたレコード（ID付き）
   */
  add(data) {
    const lock = LockService.getScriptLock();
    try {
      // 最大30秒待機
      lock.waitLock(30000);

      const lastRow = this.sheet.getLastRow();
      const idKey = this.headers[0];

      const lastId =
        lastRow <= 1 ? 0 : this.sheet.getRange(lastRow, 1).getValue();
      const newId = (Number(lastId) || 0) + 1;

      const newRecord = { [idKey]: newId, ...data };
      const rowValue = this._toRowArray(newRecord);

      this.sheet.appendRow(rowValue);
      return newRecord;
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * 指定したIDのレコードを更新します。
   * @param {number|string} id
   * @param {Object} data 更新するプロパティーを含むオブジェクト
   * @returns {boolean} 成功したらtrue
   */
  update(id, data) {
    const lock = LockService.getScriptLock();
    try {
      // 最大30秒待機
      lock.waitLock(30000);

      const rowIndex = this._getRowIndexById(id);
      if (!rowIndex) {
        return false;
      }

      const currentRecord = this._getObjectByRowIndex(rowIndex);
      const updatedRecord = { ...currentRecord, ...data };
      const rowValue = this._toRowArray(updatedRecord);

      this.sheet.getRange(rowIndex, 1, 1, this.lastCol).setValues([rowValue]);
      return true;
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * 指定したIDのレコードを削除します。
   * @param {number|string} id
   * @returns {boolean} 成功したらtrue
   */
  remove(id) {
    const rowIndex = this._getRowIndexById(id);
    if (!rowIndex) {
      return false;
    }

    this.sheet.deleteRow(rowIndex);
    return true;
  }

  // --- 内部補助メソッド ---

  /** @private */
  _toObject(rowArray) {
    return this.headers.reduce((obj, key, index) => {
      obj[key] = rowArray[index];
      return obj;
    }, {});
  }

  /** @private */
  _toRowArray(obj) {
    return this.headers.map((key) => obj[key] ?? "");
  }

  /** @private */
  _getRowIndexById(id) {
    const lastRow = this.sheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }

    const idColumnValues = this.sheet.getRange(1, 1, lastRow, 1).getValues();

    const stringId = String(id);
    for (let i = 1; i < idColumnValues.length; i++) {
      if (String(idColumnValues[i][0]) === stringId) {
        return i + 1;
      }
    }
    return null;
  }

  /** @private */
  _getObjectByRowIndex(rowIndex) {
    const rowValues = this.sheet
      .getRange(rowIndex, 1, 1, this.lastCol)
      .getValues()[0];
    return this._toObject(rowValues);
  }
}
