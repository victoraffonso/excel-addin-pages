/**
 * Excel Operations — maps MCP tool names to Office.js Excel API calls.
 * Each function receives an Excel.RequestContext and tool arguments.
 * Operations are queued and executed in a single context.sync() by the caller.
 */

/**
 * Main dispatcher — routes tool name to the correct operation function.
 * Validates the active workbook matches the requested file_path.
 */
export async function executeOperation(context, tool, args) {
  const handler = OPERATIONS[tool];
  if (!handler) {
    throw new Error(`Unknown operation: ${tool}`);
  }

  // Validate workbook and sheet if provided
  if (args.file_path || args.sheet_name) {
    try {
      const wb = context.workbook;
      wb.load('name');
      const sheets = wb.worksheets;
      sheets.load('items/name');
      await context.sync();

      // Verify file_path matches active workbook
      if (args.file_path) {
        const expectedName = args.file_path.split('/').pop().replace(/\.[^.]+$/, '');
        const actualName = (wb.name || '').replace(/\.[^.]+$/, '');
        if (expectedName && actualName && !actualName.includes(expectedName) && !expectedName.includes(actualName)) {
          throw new Error(`Active workbook "${wb.name}" does not match expected "${args.file_path}". Open the correct file in Excel.`);
        }
      }

      // Verify sheet exists
      if (args.sheet_name) {
        const names = sheets.items.map(s => s.name);
        if (!names.includes(args.sheet_name)) {
          throw new Error(`Sheet "${args.sheet_name}" not found. Active workbook "${wb.name}" has sheets: [${names.join(', ')}].`);
        }
      }
    } catch (e) {
      if (e.message.includes('not found') || e.message.includes('does not match')) throw e;
    }
  }

  return await handler(context, args);
}

// Helper: get worksheet by name
function getSheet(context, sheetName) {
  return context.workbook.worksheets.getItem(sheetName);
}

// Helper: parse hex color to Office.js format
function parseColor(hex) {
  return hex ? '#' + hex.replace('#', '') : null;
}

// Helper: column letter to 0-based index
function colToIndex(letter) {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result; // 1-based
}

const OPERATIONS = {
  // ─── Workbook ───
  async get_workbook_info(context, args) {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();

    const info = [];
    for (const ws of sheets.items) {
      const used = ws.getUsedRange();
      used.load('rowCount,columnCount,address');
      await context.sync();
      info.push({
        name: ws.name,
        rows: used.rowCount,
        columns: used.columnCount,
        address: used.address,
      });
    }
    return { sheet_count: sheets.items.length, sheets: info };
  },

  async list_sheets(context, args) {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    return sheets.items.map(s => s.name);
  },

  // ─── Sheet ───
  async create_sheet(context, args) {
    const ws = context.workbook.worksheets.add(args.sheet_name);
    if (args.after) {
      const ref = context.workbook.worksheets.getItem(args.after);
      ref.load('position');
      await context.sync();
      ws.position = ref.position + 1;
    }
    return `Created sheet: ${args.sheet_name}`;
  },

  async delete_sheet(context, args) {
    context.workbook.worksheets.getItem(args.sheet_name).delete();
    return `Deleted sheet: ${args.sheet_name}`;
  },

  async rename_sheet(context, args) {
    const ws = context.workbook.worksheets.getItem(args.old_name);
    ws.name = args.new_name;
    return `Renamed: ${args.old_name} → ${args.new_name}`;
  },

  async copy_sheet(context, args) {
    const src = context.workbook.worksheets.getItem(args.source_sheet);
    const copy = src.copy('End');
    copy.load('name');
    await context.sync();
    copy.name = args.new_name;
    return `Copied: ${args.source_sheet} → ${args.new_name}`;
  },

  // ─── Read ───
  async read_cell(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.cell);
    rng.load('values,formulas,numberFormat');
    rng.format.font.load('bold,size,name');
    await context.sync();

    return {
      coordinate: args.cell,
      value: rng.values[0][0],
      formula: rng.formulas[0][0]?.toString().startsWith('=') ? rng.formulas[0][0] : null,
      number_format: rng.numberFormat[0][0],
      font_bold: rng.format.font.bold,
      font_size: rng.format.font.size,
      font_name: rng.format.font.name,
    };
  },

  async read_range(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.range_ref);
    rng.load('values');
    await context.sync();
    return rng.values;
  },

  async read_range_with_formulas(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.range_ref);
    rng.load('formulas');
    await context.sync();
    return rng.formulas;
  },

  async read_sheet_as_table(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const used = ws.getUsedRange();
    used.load('values,rowCount,columnCount');
    await context.sync();

    const allData = used.values;
    const maxRows = args.max_rows || 500;
    const hasHeader = args.has_header !== false;

    if (hasHeader && allData.length > 0) {
      const headers = allData[0].map((h, i) => h != null ? String(h) : `col_${i}`);
      const rows = allData.slice(1, maxRows + 1);
      return { headers, rows, total_rows: allData.length - 1, returned_rows: rows.length };
    }
    const headers = allData[0]?.map((_, i) => `col_${i}`) || [];
    const rows = allData.slice(0, maxRows);
    return { headers, rows, total_rows: allData.length, returned_rows: rows.length };
  },

  async read_cell_details(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.cell);
    rng.load('values,formulas,numberFormat');
    rng.format.font.load('bold,italic,size,name,color');
    rng.format.fill.load('color');
    rng.format.load('columnWidth,rowHeight');
    await context.sync();

    const info = {
      coordinate: args.cell,
      value: rng.values[0][0],
      number_format: rng.numberFormat[0][0],
    };
    const formula = rng.formulas[0][0];
    if (formula && String(formula).startsWith('=')) info.formula = formula;

    info.font = {
      name: rng.format.font.name,
      size: rng.format.font.size,
      bold: rng.format.font.bold,
      italic: rng.format.font.italic,
      color: rng.format.font.color,
    };
    info.fill_color = rng.format.fill.color;
    info.column_width = rng.format.columnWidth;
    info.row_height = rng.format.rowHeight;
    return info;
  },

  // ─── Write ───
  async write_cell(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(args.cell).values = [[args.value]];
    return `${args.cell} = ${args.value}`;
  },

  async write_range(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rows = args.data.length;
    const cols = args.data[0]?.length || 0;
    // Calculate end cell from start_cell + dimensions
    ws.getRange(args.start_cell).getResizedRange(rows - 1, cols - 1).values = args.data;
    return `Wrote ${rows}x${cols} from ${args.start_cell}`;
  },

  async write_formula(context, args) {
    const ws = getSheet(context, args.sheet_name);
    let formula = args.formula;
    if (!formula.startsWith('=')) formula = '=' + formula;
    ws.getRange(args.cell).formulas = [[formula]];
    return `${args.cell} = ${formula.substring(0, 80)}${formula.length > 80 ? '...' : ''}`;
  },

  async write_formulas_range(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rows = args.formulas.length;
    const cols = args.formulas[0]?.length || 0;
    ws.getRange(args.start_cell).getResizedRange(rows - 1, cols - 1).formulas = args.formulas;
    return `Wrote ${rows}x${cols} formulas from ${args.start_cell}`;
  },

  // ─── Format ───
  async format_font(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const font = ws.getRange(args.range_ref).format.font;
    if (args.bold != null) font.bold = args.bold;
    if (args.italic != null) font.italic = args.italic;
    if (args.size != null) font.size = args.size;
    if (args.name != null) font.name = args.name;
    if (args.color != null) font.color = parseColor(args.color);
    return `Font applied to ${args.range_ref}`;
  },

  async format_fill(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(args.range_ref).format.fill.color = parseColor(args.color);
    return `Fill ${args.color} applied to ${args.range_ref}`;
  },

  async format_number(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(args.range_ref).numberFormat = args.number_format;
    return `Format '${args.number_format}' on ${args.range_ref}`;
  },

  async set_column_width(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(`${args.column}:${args.column}`).format.columnWidth = args.width;
    return `Column ${args.column} width = ${args.width}`;
  },

  async set_row_height(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(`${args.row}:${args.row}`).format.rowHeight = args.height;
    return `Row ${args.row} height = ${args.height}`;
  },

  async merge_cells(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(args.range_ref).merge();
    return `Merged ${args.range_ref}`;
  },

  async unmerge_cells(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.getRange(args.range_ref).unmerge();
    return `Unmerged ${args.range_ref}`;
  },

  async autofit(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = args.range_ref ? ws.getRange(args.range_ref) : ws.getUsedRange();
    rng.format.autofitColumns();
    rng.format.autofitRows();
    return `Autofit ${args.range_ref || 'used range'}`;
  },

  // ─── Professional Format ───
  async professional_format(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const used = ws.getUsedRange();
    used.load('rowCount,columnCount');
    await context.sync();

    const lastRow = used.rowCount;
    const lastCol = used.columnCount;
    const headerRow = args.header_row || 1;
    const fontName = args.font_name || 'Aptos Narrow';
    const headerColor = parseColor(args.header_color || '1F3864');
    const headerFontColor = parseColor(args.header_font_color || 'FFFFFF');
    const stripeColor = parseColor(args.stripe_color || 'D6E4F0');

    // 1. Font for entire range
    used.format.font.name = fontName;
    used.format.font.size = 11;

    // 2. Header styling
    const headerRange = ws.getRangeByIndexes(headerRow - 1, 0, 1, lastCol);
    headerRange.format.fill.color = headerColor;
    headerRange.format.font.color = headerFontColor;
    headerRange.format.font.bold = true;

    // 3. Zebra stripes
    for (let r = headerRow; r < lastRow; r++) {
      if ((r - headerRow) % 2 === 1) {
        ws.getRangeByIndexes(r, 0, 1, lastCol).format.fill.color = stripeColor;
      }
    }

    // 4. Number formats
    const formatCols = (cols, fmt) => {
      if (!cols) return;
      for (const col of cols) {
        const ci = colToIndex(col) - 1;
        ws.getRangeByIndexes(headerRow, ci, lastRow - headerRow, 1).numberFormat = fmt;
      }
    };
    formatCols(args.number_columns, '#,##0');
    formatCols(args.currency_columns, 'R$ #,##0.00');
    formatCols(args.percent_columns, '0.0%');
    formatCols(args.date_columns, 'dd/mm/yyyy');

    // 5. Borders on data area
    const dataRange = ws.getRangeByIndexes(headerRow - 1, 0, lastRow, lastCol);
    const borders = dataRange.format.borders;
    ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'].forEach(edge => {
      const b = borders.getItem(edge);
      b.style = 'Continuous';
      b.weight = 'Thin';
    });

    // 6. Autofit
    used.format.autofitColumns();

    // 7. Freeze header
    ws.freezePanes.freezeRows(headerRow);

    return `Professional format applied (${lastRow} rows x ${lastCol} cols)`;
  },

  // ─── Charts ───
  async add_chart(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const typeMap = {
      line: Excel.ChartType.line,
      bar: Excel.ChartType.barClustered,
      column: Excel.ChartType.columnClustered,
      pie: Excel.ChartType.pie,
      area: Excel.ChartType.area,
      scatter: Excel.ChartType.xyscatter,
      xy_scatter: Excel.ChartType.xyscatter,
    };
    const ct = typeMap[args.chart_type] || Excel.ChartType.line;
    const dataRange = ws.getRange(args.data_range);
    const chart = ws.charts.add(ct, dataRange, 'Auto');

    const pos = ws.getRange(args.position || 'E1');
    pos.load('left,top');
    await context.sync();
    chart.left = pos.left;
    chart.top = pos.top;
    chart.width = args.width || 400;
    chart.height = args.height || 250;
    if (args.chart_name) chart.name = args.chart_name;

    return `Chart '${args.chart_name || 'Chart'}' (${args.chart_type}) from ${args.data_range}`;
  },

  async list_charts(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const charts = ws.charts;
    charts.load('items/name,items/chartType,items/width,items/height');
    await context.sync();
    return charts.items.map(c => ({
      name: c.name,
      chart_type: c.chartType,
      width: c.width,
      height: c.height,
    }));
  },

  async delete_chart(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.charts.getItem(args.chart_name).delete();
    return `Deleted chart: ${args.chart_name}`;
  },

  // ─── Conditional Formatting ───
  async add_conditional_formatting(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.range_ref);

    if (args.rule_type === 'data_bar') {
      rng.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    } else if (args.rule_type === 'cell_value') {
      const cf = rng.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      const opMap = {
        greater_than: Excel.ConditionalCellValueOperator.greaterThan,
        less_than: Excel.ConditionalCellValueOperator.lessThan,
        between: Excel.ConditionalCellValueOperator.between,
        equal: Excel.ConditionalCellValueOperator.equalTo,
      };
      cf.cellValue.rule = {
        operator: opMap[args.operator] || Excel.ConditionalCellValueOperator.greaterThan,
        formula1: String(args.value),
        formula2: args.value2 != null ? String(args.value2) : undefined,
      };
      cf.cellValue.format.fill.color = parseColor(args.format_color || 'FF0000');
    }
    return `Conditional formatting (${args.rule_type}) on ${args.range_ref}`;
  },

  // ─── Data Validation ───
  async add_data_validation(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.range_ref);

    rng.dataValidation.clear();
    const rule = {};

    if (args.validation_type === 'list') {
      rule.list = { source: args.formula, inCellDropDown: true };
    } else if (args.validation_type === 'whole_number' || args.validation_type === 'decimal') {
      const parts = (args.formula || '0,100').split(',');
      rule.wholeNumber = {
        formula1: parts[0],
        formula2: parts[1] || parts[0],
        operator: Excel.DataValidationOperator.between,
      };
    }
    rng.dataValidation.rule = rule;
    if (args.error_message) {
      rng.dataValidation.errorAlert = { showAlert: true, message: args.error_message };
    }
    return `Data validation (${args.validation_type}) on ${args.range_ref}`;
  },

  // ─── Sort & Filter ───
  async sort_range(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.range_ref);
    const colIdx = colToIndex(args.sort_column) - 1;
    rng.sort.apply([{
      key: colIdx,
      ascending: args.ascending !== false,
    }], false, args.has_header !== false);
    return `Sorted ${args.range_ref} by ${args.sort_column}`;
  },

  async auto_filter(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.autoFilter.apply(ws.getRange(args.range_ref));
    return `Auto filter on ${args.range_ref}`;
  },

  // ─── Named Ranges ───
  async add_named_range(context, args) {
    const ws = getSheet(context, args.sheet_name);
    context.workbook.names.add(args.name, ws.getRange(args.range_ref));
    return `Named range '${args.name}' = ${args.sheet_name}!${args.range_ref}`;
  },

  async list_named_ranges(context, args) {
    const names = context.workbook.names;
    names.load('items/name,items/value');
    await context.sync();
    return names.items.map(n => ({ name: n.name, refers_to: n.value }));
  },

  async delete_named_range(context, args) {
    context.workbook.names.getItem(args.name).delete();
    return `Deleted named range: ${args.name}`;
  },

  // ─── Notes ───
  async add_note(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.cell);
    // Office.js uses comments API
    context.workbook.comments.add(rng, args.text);
    return `Note added to ${args.cell}`;
  },

  async read_note(context, args) {
    const ws = getSheet(context, args.sheet_name);
    const rng = ws.getRange(args.cell);
    try {
      const comment = rng.getComment();
      comment.load('content');
      await context.sync();
      return comment.content;
    } catch {
      return null;
    }
  },

  // ─── View ───
  async freeze_panes(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.freezePanes.freezeAt(ws.getRange(args.cell));
    return `Panes frozen at ${args.cell}`;
  },

  async set_print_area(context, args) {
    const ws = getSheet(context, args.sheet_name);
    ws.pageLayout.printArea = ws.getRange(args.range_ref);
    return `Print area set to ${args.range_ref}`;
  },

  async toggle_gridlines(context, args) {
    const ws = context.workbook.worksheets.getActiveWorksheet();
    ws.showGridlines = args.show !== false;
    return `Gridlines ${args.show ? 'shown' : 'hidden'}`;
  },
};
