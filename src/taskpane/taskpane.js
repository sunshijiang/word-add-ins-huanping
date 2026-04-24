// ============================================
// 全局状态
// ============================================

let isInTable = false;
let isApplying = false;
let debounceTimer = null;

const DEFAULT_CONFIG = {
    fontName: "宋体",
    fontSize: 10.5,
    paragraphAlign: "center",
    lineSpacing: 15,
    spaceBefore: 0,
    spaceAfter: 0,
    tableAlign: "center",
    cellVAlign: "center"
};

const WESTERN_FONT = "Times New Roman";

// ============================================
// 初始化
// ============================================

Office.onReady(() => {
    bindUI();
    startSelectionListener();
});

// ============================================
// UI 绑定
// ============================================

function bindUI() {
    document.getElementById("btnSelectAll").onclick = selectWholeTable;
    document.getElementById("btnApplyCurrent").onclick = () => applyConfig(DEFAULT_CONFIG);
}

// ============================================
// 表格检测（事件驱动）
// ============================================

function startSelectionListener() {
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChanged
    );

    checkIfInTable();
}

function onSelectionChanged() {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(checkIfInTable, 150);
}

async function checkIfInTable() {
    try {
        await Word.run(async (context) => {
            const table = await getCurrentTable(context);

            const statusBox = document.getElementById("statusBox");
            const btn = document.getElementById("btnSelectAll");

            if (table) {
                isInTable = true;
                statusBox.textContent = "✅ 在表格中";
                statusBox.className = "status in-table";
                btn.disabled = false;
            } else {
                isInTable = false;
                statusBox.textContent = "⚠️ 不在表格中";
                statusBox.className = "status out-table";
                btn.disabled = true;
            }
        });
    } catch (e) {
        console.error(e);
    }
}

// ============================================
// 获取当前表格（统一入口）
// ============================================

async function getCurrentTable(context) {
    const selection = context.document.getSelection();
    const tables = selection.tables;

    tables.load("items");
    await context.sync();

    return tables.items.length > 0 ? tables.items[0] : null;
}

// ============================================
// 全选表格
// ============================================

async function selectWholeTable() {
    if (!isInTable) return;

    await Word.run(async (context) => {
        const table = await getCurrentTable(context);

        if (!table) return;

        table.getRange().select();
        await context.sync();
    });
}

// ============================================
// 应用配置（核心优化版）
// ============================================

async function applyConfig(config) {
    if (isApplying) return;
    isApplying = true;

    try {
        await Word.run(async (context) => {

            const table = await getCurrentTable(context);

            if (!table) {
                alert("请先选中表格");
                return;
            }

            // 表格对齐
            const alignMap = {
                left: Word.Alignment.left,
                center: Word.Alignment.center,
                right: Word.Alignment.right
            };
            table.alignment = alignMap[config.tableAlign];

            // 批量加载
            table.rows.load("items");
            await context.sync();

            const rows = table.rows.items;

            rows.forEach(r => r.cells.load("items"));
            await context.sync();

            rows.forEach(r => {
                r.cells.items.forEach(cell => {
                    cell.body.paragraphs.load("items");
                });
            });
            await context.sync();

            // 批量应用
            rows.forEach(row => {
                row.cells.items.forEach(cell => {

                    const vMap = {
                        top: Word.VerticalAlignment.top,
                        center: Word.VerticalAlignment.center,
                        bottom: Word.VerticalAlignment.bottom
                    };

                    cell.verticalAlignment = vMap[config.cellVAlign];

                    cell.body.paragraphs.items.forEach(para => {

                        const pMap = {
                            left: Word.Alignment.left,
                            center: Word.Alignment.center,
                            right: Word.Alignment.right,
                            justify: Word.Alignment.justified
                        };

                        para.alignment = pMap[config.paragraphAlign];

                        para.lineSpacing = config.lineSpacing;
                        para.spaceBefore = config.spaceBefore;
                        para.spaceAfter = config.spaceAfter;

                        // 字体（完整修复）
                        para.font.name = config.fontName;
                        para.font.nameFarEast = config.fontName;
                        para.font.nameAscii = WESTERN_FONT;
                        para.font.nameComplexScript = WESTERN_FONT;

                        para.font.size = config.fontSize;
                    });
                });
            });

            await context.sync();
        });

    } catch (e) {
        console.error("应用失败:", e);
        alert(e.message);
    }

    isApplying = false;
}