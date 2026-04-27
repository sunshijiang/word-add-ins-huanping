/* global Office, Word, document, console, alert, setTimeout, clearTimeout */

let isInTable = false;
let isApplying = false;
let debounceTimer = null;

const TABLE_TEXT_CONFIG = {
  fontName: "宋体",
  westernFontName: "Times New Roman",
  fontSize: 10.5,
  alignment: Word.Alignment.center,
  lineSpacing: 15,
  spaceBefore: 0,
  spaceAfter: 0,
  leftIndent: 0,
  rightIndent: 0,
  firstLineIndent: 0,
};

const BODY_PARAGRAPH_CONFIG = {
  alignment: Word.Alignment.justified,
  leftIndent: 0,
  rightIndent: 0,
  firstLineIndent: 21,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: 22.5,
};

Office.onReady(() => {
  bindUI();
  startSelectionListener();
});

function bindUI() {
  document.getElementById("btnSelectAll").onclick = selectWholeTable;
  document.getElementById("btnCreateTable").onclick = createStandardTable;
  document.getElementById("btnApplyTable").onclick = applyCurrentTableStandard;
  document.getElementById("btnApplyParagraph").onclick = applyBodyParagraphStandard;
}

function startSelectionListener() {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );
  checkIfInTable();
}

function onSelectionChanged() {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(checkIfInTable, 120);
}

async function checkIfInTable() {
  try {
    await Word.run(async (context) => {
      const table = await getCurrentTable(context);
      const statusBox = document.getElementById("statusBox");
      const selectButton = document.getElementById("btnSelectAll");
      const applyTableButton = document.getElementById("btnApplyTable");

      if (table) {
        isInTable = true;
        statusBox.textContent = "✅ 在表格中";
        statusBox.className = "status in-table";
      } else {
        isInTable = false;
        statusBox.textContent = "⚠️ 不在表格中";
        statusBox.className = "status out-table";
      }

      selectButton.disabled = !isInTable;
      applyTableButton.disabled = !isInTable;
    });
  } catch (error) {
    console.error("检测表格状态失败：", error);
  }
}

async function getCurrentTable(context) {
  const selection = context.document.getSelection();

  // 优先使用 parentTableOrNullObject，可覆盖“光标在单元格内但未选中整表”的场景。
  selection.load("parentTableOrNullObject");
  await context.sync();

  if (selection.parentTableOrNullObject) {
    const parentTable = selection.parentTableOrNullObject;
    parentTable.load("isNullObject");
    await context.sync();
    if (!parentTable.isNullObject) {
      return parentTable;
    }
  }

  // 兼容路径：如果上面不可用，再回退到 selection.tables。
  try {
    selection.load("tables");
    await context.sync();
    const tables = selection.tables;
    tables.load("items");
    await context.sync();
    return tables.items.length > 0 ? tables.items[0] : null;
  } catch (error) {
    console.warn("回退读取 selection.tables 失败：", error.message);
    return null;
  }
}

async function selectWholeTable() {
  if (!isInTable) {
    await checkIfInTable();
  }

  await runSafe(async () => {
    await Word.run(async (context) => {
      const table = await getCurrentTable(context);
      if (!table) {
        throw new Error("请先将光标放到表格任意单元格。");
      }

      if (typeof table.select === "function") {
        table.select();
      } else {
        table.getRange().select();
      }

      await context.sync();
    });

    await checkIfInTable();
  });
}

function readDimension(id) {
  const value = parseInt(document.getElementById(id).value, 10);
  return Number.isFinite(value) && value > 0 ? value : null;
}

async function runSafe(action) {
  if (isApplying) {
    return;
  }

  isApplying = true;
  try {
    await action();
  } catch (error) {
    console.error(error);
    alert(`执行失败：${error.message}`);
  }
  isApplying = false;
}

async function safeSync(context, stage) {
  try {
    await context.sync();
    return true;
  } catch (error) {
    console.warn(`${stage} 失败，已自动降级：`, error.message);
    return false;
  }
}

async function createStandardTable() {
  const rows = readDimension("rows");
  const cols = readDimension("cols");

  if (!rows || !cols) {
    alert("请输入合法的行数和列数（大于 0 的整数）。");
    return;
  }

  await runSafe(async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const values = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ""));

      // 优先 Replace，避免部分环境下 After 在光标位置失效。
      const table = selection.insertTable(rows, cols, Word.InsertLocation.replace, values);

      // 先确保“创建动作”提交，后续格式失败也不影响创建结果。
      await context.sync();

      await applyTableShape(context, table);
      await applyTableTextAndParagraph(context, table);

      if (typeof table.select === "function") {
        table.select();
      } else {
        table.getRange().select();
      }
      await safeSync(context, "选中新建表格");
    });

    await checkIfInTable();
  });
}

async function applyCurrentTableStandard() {
  await runSafe(async () => {
    await Word.run(async (context) => {
      const table = await getCurrentTable(context);
      if (!table) {
        throw new Error("请先将光标放到目标表格中。");
      }

      await applyTableShape(context, table);
      await applyTableTextAndParagraph(context, table);
    });
  });
}

async function applyTableShape(context, table) {
  table.alignment = Word.Alignment.center;
  if (typeof table.autoFitWindow === "function") {
    table.autoFitWindow();
  } else if (typeof table.autoFitBehavior === "function") {
    table.autoFitBehavior(Word.AutoFitBehaviorType.window);
  }
  await safeSync(context, "应用表格宽度/对齐");

  // 边框设置按组提交，某一组失败不影响其他组。
  applyBorderSafely(table, Word.BorderLocation.left, 0, false);
  applyBorderSafely(table, Word.BorderLocation.right, 0, false);
  await safeSync(context, "应用左右边框");

  applyBorderSafely(table, Word.BorderLocation.top, 1.5, true);
  applyBorderSafely(table, Word.BorderLocation.bottom, 1.5, true);
  await safeSync(context, "应用上下边框");

  applyBorderSafely(table, Word.BorderLocation.insideHorizontal, 0.5, true);
  applyBorderSafely(table, Word.BorderLocation.insideVertical, 0.5, true);
  await safeSync(context, "应用内部边框");
}

function applyBorderSafely(table, borderLocation, width, visible) {
  try {
    const border = table.getBorder(borderLocation);
    border.type = visible ? Word.BorderType.single : Word.BorderType.none;
    border.color = "#000000";
    if (visible) {
      border.width = width;
    }
  } catch (error) {
    console.warn("边框对象不可用：", borderLocation, error.message);
  }
}

async function applyTableTextAndParagraph(context, table) {
  table.rows.load("items");
  if (!(await safeSync(context, "加载表格行"))) {
    return;
  }

  table.rows.items.forEach((row) => row.cells.load("items"));
  if (!(await safeSync(context, "加载单元格"))) {
    return;
  }

  table.rows.items.forEach((row) => {
    row.cells.items.forEach((cell) => {
      cell.verticalAlignment = Word.VerticalAlignment.center;
      cell.body.paragraphs.load("items");
    });
  });
  if (!(await safeSync(context, "加载单元格段落"))) {
    return;
  }

  table.rows.items.forEach((row) => {
    row.cells.items.forEach((cell) => {
      cell.body.paragraphs.items.forEach((paragraph) => {
        paragraph.alignment = TABLE_TEXT_CONFIG.alignment;
        paragraph.spaceBefore = TABLE_TEXT_CONFIG.spaceBefore;
        paragraph.spaceAfter = TABLE_TEXT_CONFIG.spaceAfter;
        paragraph.lineSpacing = TABLE_TEXT_CONFIG.lineSpacing;
        paragraph.leftIndent = TABLE_TEXT_CONFIG.leftIndent;
        paragraph.rightIndent = TABLE_TEXT_CONFIG.rightIndent;
        paragraph.firstLineIndent = TABLE_TEXT_CONFIG.firstLineIndent;

        paragraph.font.name = TABLE_TEXT_CONFIG.fontName;
        paragraph.font.nameFarEast = TABLE_TEXT_CONFIG.fontName;
        paragraph.font.nameAscii = TABLE_TEXT_CONFIG.westernFontName;
        paragraph.font.nameComplexScript = TABLE_TEXT_CONFIG.westernFontName;
        paragraph.font.size = TABLE_TEXT_CONFIG.fontSize;
      });
    });
  });

  await safeSync(context, "应用表内字体与段落");
}

async function applyBodyParagraphStandard() {
  await runSafe(async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("paragraphs");
      await context.sync();

      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (!paragraphs.items.length) {
        throw new Error("未找到可设置的段落，请先选中正文内容。");
      }

      paragraphs.items.forEach((paragraph) => {
        paragraph.alignment = BODY_PARAGRAPH_CONFIG.alignment;
        paragraph.leftIndent = BODY_PARAGRAPH_CONFIG.leftIndent;
        paragraph.rightIndent = BODY_PARAGRAPH_CONFIG.rightIndent;
        paragraph.firstLineIndent = BODY_PARAGRAPH_CONFIG.firstLineIndent;
        paragraph.spaceBefore = BODY_PARAGRAPH_CONFIG.spaceBefore;
        paragraph.spaceAfter = BODY_PARAGRAPH_CONFIG.spaceAfter;
        paragraph.lineSpacing = BODY_PARAGRAPH_CONFIG.lineSpacing;
      });

      await context.sync();
    });
  });
}
