/**
 * CheckDailyTasksAndNotify
 * 
 * 【概要】
 * SharePoint上の年間スケジュール_通年シートから本日の担当タスクを抽出します。
 * 当日に○がついている行から「担当者」と「タスク名」を抽出します。
 * 複数ある場合はすべて取得します。
 * 
 * 【前提条件】
 * - シート名：年間スケジュール_通年
 * - 3行目：日付行（mm/dd形式）
 * - A列：担当者名
 * - 日付列（BC結合セルなど）：○が入っている行から担当者とタスク名を抽出
 */

function main(workbook: ExcelScript.Workbook, targetDate: string): TaskResult[] {

    const sheet = workbook.getWorksheet("年間スケジュール_通年");
    const usedRange = sheet.getUsedRange();
    const values = usedRange.getValues() as any[][];

    const results: TaskResult[] = [];

    // 3行目が日付行
    const dateRow = values[2];
    const todayMMDD = formatDateToMMDD(targetDate);

    // 日付列を検索
    for (let col = 0; col < dateRow.length; col++) {
        const cellStr = String(dateRow[col]).trim();
        const cleaned = cellStr.replace(/（.*）/g, '').trim();

        if (cleaned === todayMMDD || cleaned === todayMMDD.replace(/^0/, '')) {

            // 該当日の全行を走査
            for (let row = 3; row < values.length; row++) {

                const person = String(values[row][0]).trim();
                if (!person) continue;

                const taskCell = values[row][col];

                // ○がある行から担当者とタスク名を抽出
                if (taskCell && String(taskCell).trim().includes('○')) {
                    results.push({
                        person: person,
                        task: String(values[row][col + 1] || taskCell || "").trim(),
                        date: targetDate
                    });
                }
            }
            break;
        }
    }

    return results;
}

function formatDateToMMDD(yyyyMMDD: string): string {
    const date = new Date(yyyyMMDD);
    return `${date.getMonth() + 1}/${date.getDate()}`;
}

interface TaskResult {
    person: string;
    task: string;
    date: string;
}
