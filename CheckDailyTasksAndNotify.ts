/**
 * CheckDailyTasksAndNotify
 * 
 * 目的: タスク管理表から今日の担当タスク（○がついているもの）を抽出
 * 作成日: 2026/06
 */
function main(workbook: ExcelScript.Workbook, targetDate: string): TaskResult[] {

    const sheet = workbook.getFirstWorksheet();           // 必要ならシート名指定: getWorksheet("タスク管理")
    const usedRange = sheet.getUsedRange();
    const values = usedRange.getValues() as any[][];      // 2次元配列

    const results: TaskResult[] = [];

    // 3行目（インデックス2）を日付行とする
    const dateRow = values[2];

    for (let col = 0; col < dateRow.length; col++) {
        const cell = dateRow[col];

        if (cell instanceof Date) {
            const cellDateStr = cell.toISOString().split('T')[0]; // yyyy-MM-dd形式

            if (cellDateStr === targetDate) {
                // 今日の日付列を発見 → 下の行を走査
                for (let row = 3; row < values.length; row++) {   // 4行目以降
                    const personCell = values[row][0];             // A列：担当者

                    if (!personCell || String(personCell).trim() === "") {
                        continue;
                    }

                    const taskCell = values[row][col];             // BC結合セルがある列

                    if (taskCell && isMarked(taskCell)) {
                        results.push({
                            person: String(personCell).trim(),
                            task: String(values[row][col + 1] || values[row][col] || "").trim(), // 必要に応じて調整
                            date: targetDate,
                            column: col
                        });
                    }
                }
                break; // 今日の日付は1列だけのはずなので終了
            }
        }
    }

    return results;
}

// ○ を検知するヘルパー関数
function isMarked(value: any): boolean {
    if (!value) return false;
    const str = String(value).trim();
    return str === '○' || 
           str === 'まる' || 
           str.includes('○') || 
           str === '✓' ||
           str === 'はい';
}

// 戻り値の型定義（Power Automateで扱いやすくするため）
interface TaskResult {
    person: string;
    task: string;
    date: string;
    column?: number;
}
