// ファイル拡張子チェックメソッド
function checkExtension(fileName) {
    // 許可拡張子一覧
    const ALLOW_EXTS = new Array('xls', 'xlsx', 'xlsm');
    
    // 拡張子存在チェック
    let position = fileName.lastIndexOf(".");
    if (position === -1) {
        return false;
    }
    
    // 拡張子が許可されたものかチェック
    let targetExt = fileName.slice(position +1);
    if (ALLOW_EXTS.indexOf(targetExt) === -1) {
        return false;
    }

    return true;
}