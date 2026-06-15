/**
 * CheckDailyTasksAndNotify
 * 
 * 【概要】
 * SharePoint上の「年間スケジュール_通年」シートから本日の担当タスクを抽出します。
 * 他社がマクロで生成する表に対応するため、テーブル化不要で動作します。
 * 
 * 【前提条件】
 * - シート名：年間スケジュール_通年
 * - 3行目（インデックス2）：日付行（mm/dd形式）
 * - A列：担当者名（文字列）
 * - 日付列：BC結合セルなどに「○」または「まる」が入力されている
 * 
 * 【戻り値】
 * Power Automateに以下の配列を返します：
 * [{ person: string, task: string, date: string }, ...]
 */

function main(workbook: ExcelScript.Workbook, targetDate: string): TaskResult[] {

    // ------------------- シート取得 -------------------
    const sheetName = "年間スケジュール_通年";
    const sheet = workbook.getWorksheet(sheetName);

    // 使用範囲全体を取得
    const usedRange = sheet.getUsedRange();
    const values = usedRange.getValues() as any[][];

    const results: TaskResult[] = [];

    // ------------------- 日付行の処理 -------------------
    // 3行目が日付行（mm/dd形式）
    const dateRow = values[2];

    // 今日の日付（yyyy-MM-dd）を mm/dd形式に変換
    const todayMMDD = formatDateToMMDD(targetDate);

    // 日付行を左から順に検索
    for (let col = 0; col < dateRow.length; col++) {
        const cellValue = dateRow[col];
        const cellStr = String(cellValue).trim();

        // mm/dd形式の日付と一致するかチェック
        if (isMatchingDate(cellStr, todayMMDD)) {

            // ------------------- 該当列のタスクを抽出 -------------------
            for (let row = 3; row < values.length; row++) {

                const personCell = values[row][0];   // A列：担当者

                // 担当者名が空の場合はスキップ
                if (!personCell || String(personCell).trim() === "") {
                    continue;
                }

                const taskCell = values[row][col];   // 今日の日付列のセル

                // ○マークがあれば結果に追加
                if (taskCell && isMarked(taskCell)) {
                    results.push({
                        person: String(personCell).trim(),
                        task: String(values[row][col + 1] || taskCell || "").trim(),
                        date: targetDate
                    });
                }
            }
            break;   // 今日の日付列が見つかったら終了
        }
    }

    return results;
}

// ------------------- ヘルパー関数 -------------------

/**
 * mm/dd形式の日付と一致するか判定
 */
function isMatchingDate(cellValue: string, todayMMDD: string): boolean {
    if (!cellValue) return false;
    
    // 不要な文字を除去（例: "6/15" や "06/15" や "6/15（水）" など）
    const cleaned = cellValue.replace(/（.*）/g, '').trim();
    return cleaned === todayMMDD || cleaned === todayMMDD.replace(/^0/, '');
}

/**
 * yyyy-MM-dd を mm/dd形式に変換（例: 2026-06-15 → 6/15）
 */
function formatDateToMMDD(yyyyMMDD: string): string {
    const date = new Date(yyyyMMDD);
    const month = date.getMonth() + 1;   // 0始まりなので+1
    const day = date.getDate();
    return `${month}/${day}`;
}

/**
 * ○マーク判定
 */
function isMarked(value: any): boolean {
    if (!value) return false;
    const str = String(value).trim();
    return str.includes('○') || 
           str === 'まる' || 
           str === '✓';
}

/**
 * 戻り値の型定義
 */
interface TaskResult {
    person: string;
    task: string;
    date: string;
}
