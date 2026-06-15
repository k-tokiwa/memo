/**
 * CheckDailyTasksAndNotify
 * 
 * 【概要】
 * SharePoint上の年間スケジュール_通年シートから本日の担当タスクを抽出します。
 * 当日に○がついている行から「担当者」と「タスク名」を抽出します。
 */

function main(workbook: ExcelScript.Workbook, targetDate: string): TaskResult[] {

    const sheet = workbook.getWorksheet("年間スケジュール_通年");
    const usedRange = sheet.getUsedRange();
    const values: (string | number | boolean | Date)[][] = usedRange.getValues();

    const results: TaskResult[] = [];

    const dateRow = values[2];
    const todayMMDD = formatDateToMMDD(targetDate);

    for (let col = 0; col < dateRow.length; col++) {
        const cellStr = String(dateRow[col]).trim();
        const cleaned = cellStr.replace(/（.*）/g, '').trim();

        if (cleaned === todayMMDD || cleaned === todayMMDD.replace(/^0/, '')) {

            for (let row = 3; row < values.length; row++) {

                const person = String(values[row][0]).trim();
                if (!person) continue;

                const taskCell = values[row][col];

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
