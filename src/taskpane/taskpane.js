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
  debounceTimer = setTimeout(checkIfInTable, 150);
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
    console.error(error);
  }
}

async function getCurrentTable(context) {
  const selection = context.document.getSelection();
  selection.load("tables");
  await context.sync();

  const tables = selection.tables;
  tables.load("items");
  await context.sync();

  return tables.items.length > 0 ? tables.items[0] : null;
}

async function selectWholeTable() {
  if (!isInTable) {
    return;
  }

  await Word.run(async (context) => {
    const table = await getCurrentTable(context);
    if (!table) {
      return;
    }
    table.getRange().select();
    await context.sync();
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
      const table = selection.insertTable(rows, cols, Word.InsertLocation.after, values);

      applyTableShape(table);
      await applyTableTextAndParagraph(context, table);

      table.getRange().select();
      await context.sync();
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

      applyTableShape(table);
      await applyTableTextAndParagraph(context, table);
      await context.sync();
    });
  });
}

function applyTableShape(table) {
  table.alignment = Word.Alignment.center;

  if (typeof table.autoFitWindow === "function") {
    table.autoFitWindow();
  } else if (typeof table.autoFitBehavior === "function") {
    table.autoFitBehavior(Word.AutoFitBehaviorType.window);
  }

  setBorderIfSupported(table, Word.BorderLocation.top, 1.5, true);
  setBorderIfSupported(table, Word.BorderLocation.bottom, 1.5, true);
  setBorderIfSupported(table, Word.BorderLocation.left, 0, false);
  setBorderIfSupported(table, Word.BorderLocation.right, 0, false);
  setBorderIfSupported(table, Word.BorderLocation.insideHorizontal, 0.5, true);
  setBorderIfSupported(table, Word.BorderLocation.insideVertical, 0.5, true);
}

function setBorderIfSupported(table, borderLocation, width, visible) {
  try {
    const border = table.getBorder(borderLocation);
    border.type = visible ? Word.BorderType.single : Word.BorderType.none;
    border.color = "#000000";
    if (visible) {
      border.width = width;
    }
  } catch (error) {
    console.warn("当前 Word 版本不支持该边框设置：", borderLocation, error.message);
  }
}

async function applyTableTextAndParagraph(context, table) {
  table.rows.load("items");
  await context.sync();

  table.rows.items.forEach((row) => {
    row.height = 0;
    row.cells.load("items");
  });
  await context.sync();

  table.rows.items.forEach((row) => {
    row.cells.items.forEach((cell) => {
      cell.verticalAlignment = Word.VerticalAlignment.center;
      cell.width = 0;
      cell.body.paragraphs.load("items");
    });
  });
  await context.sync();

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
        paragraph.font.nameOther = TABLE_TEXT_CONFIG.westernFontName;
        paragraph.font.nameComplexScript = TABLE_TEXT_CONFIG.westernFontName;
        paragraph.font.size = TABLE_TEXT_CONFIG.fontSize;
      });
    });
  });
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
