// Power AutomateのJSON構造を定義
interface FlowAction {
    type?: string;
    inputs?: unknown;
    parameters?: unknown;
    expression?: unknown;
    recurrence?: unknown;
    runAfter?: Record<string, string[]>;
    actions?: Record<string, FlowAction>;
    else?: {
        actions?: Record<string, FlowAction>;
    };
    cases?: Record<string, { case?: unknown; actions?: Record<string, FlowAction> }>;
    default?: { actions?: Record<string, FlowAction> };
}

interface FlowData {
    properties?: {
        definition?: {
            triggers?: Record<string, FlowAction>;
            actions?: Record<string, FlowAction>;
        }
    }
}

function getOrderedActionNames(actionsObj: Record<string, FlowAction>): string[] {
    const allNames = Object.keys(actionsObj);
    const ordered: string[] = [];
    const remaining = new Set(allNames);

    let loopGuard = 0;
    while (remaining.size > 0 && loopGuard < 1000) {
        let addedInThisPass = false;
        const candidates = Array.from(remaining).filter(name => {
            const action = actionsObj[name];
            const runAfter = action.runAfter || {};
            const deps = Object.keys(runAfter);
            return deps.every(dep => !allNames.includes(dep) || ordered.includes(dep));
        });

        candidates.sort();

        for (const name of candidates) {
            ordered.push(name);
            remaining.delete(name);
            addedInThisPass = true;
        }

        if (!addedInThisPass) {
            Array.from(remaining).forEach(n => ordered.push(n));
            break;
        }
        loopGuard++;
    }
    return ordered;
}

function getOperationId(inputsObj: unknown): string {
    if (!inputsObj || typeof inputsObj !== 'object') return "";
    const inputs = inputsObj as Record<string, unknown>;

    if(inputs.host && typeof inputs.host === 'object') {
        const host = inputs.host as Record<string, unknown>;
        if (host.operationId) {
            return String(host.operationId);
        }
    }
    return "";
}

function main(workbook: ExcelScript.Workbook) {
    const jsonSheet = workbook.getWorksheet("JSON");
    if(!jsonSheet) {
        console.log("エラー： 'JSON'シートがありません。");
        return;
    }

    const jsonString = jsonSheet.getRange("A1").getValue() as string;
    if(!jsonString) {
        console.log("エラー： A1セルが空です。");
        return;
    }

    let flowData: FlowData;
    try {
        flowData = JSON.parse(jsonString) as FlowData;
    } catch (e) {
        console.log("エラー： JSONの形式が正しくありません。");
        return;
    }

    const outSheetName = "フロー設計書";
    let outSheet = workbook.getWorksheet(outSheetName);
    if (outSheet) { outSheet.delete(); }
    outSheet = workbook.addWorksheet(outSheetName);

    const headers = ["No", "階層 / 分岐ルート", "アクション名（変更後）", "種類 / デフォルト名", "入力パラメーター（inputs/条件など）", "実行条件（runAfter）"];
    outSheet.getRange("A1:F1").setValues([headers]);

    const rows: string[][] = [];
    const rowColors: string[] = [];
    let counter = 1;

    // 3. トリガーの解析（recurrence対応）
    const triggers: Record<string, FlowAction> = flowData.properties?.definition?.triggers || {};
    for (const tName in triggers) {
        const trigger = triggers[tName];
        const type = trigger.type || "Unknown";
        const opId = getOperationId(trigger.inputs);
        const displayType = opId ? `${type}\n(${opId})` : type;

        // inputsに加えてrecurrenceがあればマージする
        let targetTriggerInputs: Record<string, unknown> = (trigger.inputs as Record<string, unknown>) || {};
        if (trigger.recurrence) {
            targetTriggerInputs = { "【スケジュール設定】": trigger.recurrence, ...targetTriggerInputs };
        }
        const inputs = JSON.stringify(targetTriggerInputs, null, 2);

        rows.push([counter.toString(), "Root（トリガー）", tName, displayType, inputs, "-"]);
        rowColors.push("");
        counter++;
    }

    // 4. アクションの解析関数
    function parseActions(actionsObj: Record<string, FlowAction>, parent: string, currentColor: string) {
        if (!actionsObj || Object.keys(actionsObj).length === 0) return;

        const orderedNames = getOrderedActionNames(actionsObj);

        for (const aName of orderedNames) {
            const action = actionsObj[aName];
            const type = action.type || "Unknown";

            let targetInputs: Record<string, unknown> = (action.inputs as Record<string, unknown>) || (action.parameters as Record<string, unknown>) || {};
            const opId = getOperationId(action.inputs);
            const displayType = opId ? `${type}\n(${opId})` : type;

            if (type === "If" && action.expression) {
                targetInputs = { "【条件式】": action.expression, ...targetInputs };
            }

            const inputsStr = JSON.stringify(targetInputs, null, 2);

            const runAfterObj = action.runAfter || {};
            const runAfterStr = Object.keys(runAfterObj).length > 0 ? Object.keys(runAfterObj).map(k => `${k}\n(${runAfterObj[k].join(", ")})`).join("\n\n") : "（前ステップ実行後）";

            rows.push([counter.toString(), parent, aName, displayType, inputsStr, runAfterStr]);
            rowColors.push(currentColor);
            counter++;

            let nextColor = currentColor;
            if (type === "Scope") {
                nextColor = '#FCE4D6';
            } else if (type === "Foreach" || type === "Until") {
                nextColor = '#F2F2F2';
            }

            if (action.actions) {
                const childParent = (type === "If") ? `${parent} > ${aName} [◯ True]` : `${parent} > ${aName}`;
                const passColor = (type === "If") ? "#E2EFDA" : nextColor;
                parseActions(action.actions, childParent, passColor);
            } 
            if (action.else && action.else.actions) {
                const childParent = `${parent} > ${aName} [✕ False]`;
                const passColor = (type === "If") ? "#F4CCCC" : nextColor;
                parseActions(action.else.actions, childParent, passColor);
            }
            if (action.cases) {
                for (const cName in action.cases) {
                    const cObj = action.cases[cName];
                    if (cObj.actions) {
                        parseActions(cObj.actions, `${parent} > ${aName} [Case: ${cObj.case || cName}]`, nextColor);
                    }
                }
            }
            if (action.default && action.default.actions) {
                parseActions(action.default.actions, `${parent} > ${aName} [既定（Default）]`, nextColor);
            }
        }
    }

    const rootActions: Record<string, FlowAction> = flowData.properties?.definition?.actions || {};
    parseActions(rootActions, "Root", "");

    // 5. シートへの書き込みと書式設定
    if (rows.length > 0) {
        const dataRange = outSheet.getRangeByIndexes(1, 0, rows.length, headers.length);
        dataRange.setValues(rows);

        const headerRange = outSheet.getRange("A1:F1");
        headerRange.getFormat().getFill().setColor("#4F818D");
        headerRange.getFormat().getFont().setColor("White");
        headerRange.getFormat().getFont().setBold(true);

        const fullRange = outSheet.getUsedRange();
        if (fullRange) {
            fullRange.getFormat().setWrapText(true);
            fullRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
        }

        for (let i = 0; i < rowColors.length; i++) {
            const color = rowColors[i];
            if (color) {
                const rowRange = outSheet.getRangeByIndexes(i + 1, 0, 1, headers.length);
                rowRange.getFormat().getFill().setColor(color);
            }
        }

        outSheet.getRange("A:A").getFormat().setColumnWidth(40);
        outSheet.getRange("B:B").getFormat().setColumnWidth(250);
        outSheet.getRange("C:C").getFormat().setColumnWidth(180);
        outSheet.getRange("D:D").getFormat().setColumnWidth(160);
        outSheet.getRange("E:E").getFormat().setColumnWidth(350);
        outSheet.getRange("F:F").getFormat().setColumnWidth(150);
    }

    console.log("フロー設計書の出力（スケジュールトリガー対応版）が完了しました。");
}