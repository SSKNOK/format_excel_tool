/**
 * checkKinoGaiyoSheet：[機能概要]シートのお作法チェックを行います。
 * @param workSheet ワークシートオブジェクト
 * @param workSheetName ワークシート名
 * 
 */
function checkKinoGaiyoSheet(workSheet, workSheetName) {
    let errorList = [];
    for (errorItem of checkNoNumber(workSheet, workSheetName)){
        errorList.push(errorItem);
    }
    console.log("checkKinoGaiyoSheet:L12");
    console.log(errorList);
    return errorList;
}

/**
 * checkNoNumber：No項目番号チェック
 * @param workSheet ワークシートオブジェクト
 * @param workSheetName  ワークシート名
 * 
 * [概要]
 * 「No」とついた項目番号の連番が正しいかチェックします。
 */
function checkNoNumber(workSheet, workSheetName) {

    // 「No」項目名
    const NO_HEADER = 'No';
    // エラーメッセージ
    ERROR_MESSAGE = '連番が不整合です';
    // 連番不整合リスト
    let errorList = [];
    // チェック中フラグ
    let isDoingCheck = false;

    // 使用されているセル範囲を取得
    let range = XLSX.utils.decode_range(workSheet['!ref']);
					
    // 列のループ
    for (var colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
        let currentNo = 0;

        // 行のループ
        for (var rowIdx = range.s.r; rowIdx <= range.e.r; rowIdx++) {
            
            // セルのアドレスを取得する
            let address = XLSX.utils.encode_cell({ r: rowIdx, c:colIdx });
            let cell = workSheet[address];

            if (cell === undefined && currentNo === 0) {
                continue;
            }

            if (cell === undefined && currentNo !== 0) {
                isDoingCheck = false;
                continue;
            }

            // セルの値が「No」であればチェックスタート
            if (cell.v === NO_HEADER) {
                currentNo = 0;
                isDoingCheck = true;
                continue;
            }

            // セルにNoが設定されている場合の値の整合性チェック
            if (isDoingCheck) {
                if (cell.v == currentNo + 1) { // 連番の整合性がとれている場合
                    currentNo = currentNo + 1;
                } else if (cell.v != currentNo + 1) { // 連番の整合性がとれていない場合
                    errorList.push({sheet: workSheetName, cell: address, message: ERROR_MESSAGE});
                    currentNo = currentNo + 1;
                }
            }
        }
    }
    return errorList;
}