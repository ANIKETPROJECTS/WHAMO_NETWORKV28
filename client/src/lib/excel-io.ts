import ExcelJS from 'exceljs';

// ─── Types ───────────────────────────────────────────────────────────────────

export type FilterKey =
  | 'all' | 'conduit' | 'dummy'
  | 'node' | 'reservoir' | 'junction' | 'surgeTank' | 'flowBoundary' | 'pump' | 'checkValve' | 'turbine';

type ColType = 'text' | 'number' | 'boolean' | 'dropdown';

interface ColDef {
  key: string;          // data field name
  header: string;       // column header shown in Excel
  type: ColType;
  options?: string[];   // for dropdown columns
  readOnly?: boolean;   // won't be imported back
}

// ─── Column Definitions per Tab ───────────────────────────────────────────────

const PIPE_TYPE_OPTIONS = ['conduit', 'dummy'];
const PIPE_MATERIAL_OPTIONS = [
  '-- None --',
  'Aluminum','Aluminum structural plate 32 in CR','Aluminum structural plate 32 in CR Historic',
  'Asbestos Cement','Asphalt ditch','Asphalt pavement (rough)','Asphalt pavement (smooth)',
  'Asphalted cast iron (new)','Bare soil','Best concrete','Brick in mortar','Brick sewer',
  'Cast iron','CMP','Concrete','Concrete (centrif. spun)','Concrete (steel forms)',
  'Concrete (wood forms)','Concrete gutter (broom finish)','Concrete gutter (troweled finish)',
  'Concrete gutter, asphalt pavement (rough)','Concrete gutter, asphalt pavement (smooth)',
  'Concrete pavement (float finish)','Copper','Curled wood mat','Ductile Iron',
  'Fiberglass roving','Flood plain, brush','Flood plain, cultivated','Galvanized iron',
  'Glass','Gravel riprap, 25 mm (1 in) D50','Gravel riprap, 50 mm (2 in) D50',
  'Grouted riprap','Jute net','Natural stream, clean','Natural stream, stony notes',
  'Natural stream, weedy','PVC','Riveted steel (new, rough)','Riveted steel (new, smooth)',
  'Rock cut','Rock riprap, 150 mm (6 in) D50','Rock riprap, 300 mm (12 in) D50',
  'Rough channel, with grass','Rough earth','Rough rocks','Soil cement','Steel',
  'Steel and Aluminum 18 In or less CR 3x1 Corrugations Historic',
  'Steel and Aluminum 5x1 and 3x1 Corrugations','Steel and Aluminum Var CR',
  'Steel and Aluminum Var CR Historic','Steel structural plate 18 In CR',
  'Steel structural plate 31 In CR','Steel structural plate 47 In CR',
  'Stone masonry','Stony bottom','Straw with net','Synthetic mat',
  'Very rough channel, with grass','Wood Stave (new, smooth)','Woven paper net',
];
const BC_MODE_OPTIONS = ['fixed', 'schedule'];
const SURGE_TANK_TYPE_OPTIONS = ['SIMPLE', 'DIFFERENTIAL', 'AIRTANK'];
const PUMP_STATUS_OPTIONS = ['ACTIVE', 'INACTIVE'];
const VALVE_STATUS_OPTIONS = ['OPEN', 'CLOSED'];
const TURBINE_MODE_OPTIONS = ['TURBINE', 'GENERATE', 'TURBGOV', 'EMERGENCY'];
const BOOL_OPTIONS = ['true', 'false'];

export const TAB_COLS: Record<FilterKey, ColDef[]> = {
  all: [
    { key: 'label',      header: 'Label',          type: 'text' },
    { key: '_type',      header: 'Type',            type: 'text', readOnly: true },
    { key: '_unit',      header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'type',       header: 'Pipe Type',       type: 'dropdown', options: PIPE_TYPE_OPTIONS },
    { key: 'nodeNumber', header: 'Node #',          type: 'number' },
    { key: 'diameter',   header: 'Diameter',        type: 'number' },
    { key: 'length',     header: 'Length',          type: 'number' },
    { key: 'celerity',   header: 'Wave Speed',      type: 'number' },
    { key: 'friction',   header: 'Friction',        type: 'number' },
    { key: 'elevation',  header: 'Elevation',       type: 'number' },
    { key: 'comment',    header: 'Comment',         type: 'text' },
  ],
  conduit: [
    { key: 'label',              header: 'Label',           type: 'text' },
    { key: '_unit',              header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'type',               header: 'Pipe Type',       type: 'dropdown', options: PIPE_TYPE_OPTIONS },
    { key: '_materialLabel',     header: 'Pipe Material',   type: 'dropdown', options: PIPE_MATERIAL_OPTIONS },
    { key: 'length',             header: 'Length',          type: 'number' },
    { key: 'diameter',           header: 'Diameter',        type: 'number' },
    { key: 'celerity',           header: 'Wave Speed',      type: 'number' },
    { key: 'friction',           header: 'Friction',        type: 'number' },
    { key: 'manningsN',          header: "Manning's n",     type: 'number' },
    { key: 'numSegments',        header: 'Segments',        type: 'number' },
    { key: 'includeNumSegments', header: 'Incl. in INP',   type: 'dropdown', options: BOOL_OPTIONS },
    { key: 'hasAddedLoss',       header: 'Added Loss',      type: 'dropdown', options: BOOL_OPTIONS },
    { key: 'cplus',              header: 'CPLUS',           type: 'number' },
    { key: 'cminus',             header: 'CMINUS',          type: 'number' },
    { key: 'pipeE',              header: 'E (Modulus)',     type: 'number' },
    { key: 'pipeWT',             header: 'WT (Wall Thick)', type: 'number' },
    { key: 'variable',           header: 'VARIABLE',        type: 'dropdown', options: BOOL_OPTIONS },
    { key: 'distance',           header: 'Distance',        type: 'number' },
    { key: 'area',               header: 'Area',            type: 'number' },
    { key: 'comment',            header: 'Comment',         type: 'text' },
  ],
  dummy: [
    { key: 'label',        header: 'Label',           type: 'text' },
    { key: '_unit',        header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'type',         header: 'Pipe Type',       type: 'dropdown', options: PIPE_TYPE_OPTIONS },
    { key: 'diameter',     header: 'Diameter',        type: 'number' },
    { key: 'hasAddedLoss', header: 'Added Loss',      type: 'dropdown', options: BOOL_OPTIONS },
    { key: 'cplus',        header: 'CPLUS',           type: 'number' },
    { key: 'cminus',       header: 'CMINUS',          type: 'number' },
    { key: 'comment',      header: 'Comment',         type: 'text' },
  ],
  node: [
    { key: 'label',      header: 'Label',           type: 'text' },
    { key: '_type',      header: 'Type',            type: 'text', readOnly: true },
    { key: '_unit',      header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber', header: 'Node #',          type: 'number' },
    { key: 'elevation',  header: 'Elevation',       type: 'number' },
    { key: 'comment',    header: 'Comment',         type: 'text' },
  ],
  reservoir: [
    { key: 'label',              header: 'Label',            type: 'text' },
    { key: '_unit',              header: 'Unit (SI/FPS)',    type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber',         header: 'Node #',           type: 'number' },
    { key: 'elevation',          header: 'Elevation',        type: 'number' },
    { key: 'mode',               header: 'BC Mode',          type: 'dropdown', options: BC_MODE_OPTIONS },
    { key: 'reservoirElevation', header: 'Res. Elevation',   type: 'number' },
    { key: 'hScheduleNumber',    header: 'H Sched #',        type: 'number' },
    { key: 'comment',            header: 'Comment',          type: 'text' },
  ],
  junction: [
    { key: 'label',      header: 'Label',           type: 'text' },
    { key: '_unit',      header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber', header: 'Node #',          type: 'number' },
    { key: 'elevation',  header: 'Elevation',       type: 'number' },
    { key: 'comment',    header: 'Comment',         type: 'text' },
  ],
  surgeTank: [
    { key: 'label',             header: 'Label',             type: 'text' },
    { key: '_unit',             header: 'Unit (SI/FPS)',     type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber',        header: 'Node #',            type: 'number' },
    { key: 'elevation',         header: 'Elevation',         type: 'number' },
    { key: 'type_st',           header: 'Tank Type',         type: 'dropdown', options: SURGE_TANK_TYPE_OPTIONS },
    { key: 'tankTop',           header: 'Top Elevation',     type: 'number' },
    { key: 'tankBottom',        header: 'Bot. Elevation',    type: 'number' },
    { key: 'initialWaterLevel', header: 'HTANK',             type: 'number' },
    { key: 'riserDiameter',     header: 'Riser Diameter',    type: 'number' },
    { key: 'riserTop',          header: 'Riser Top',         type: 'number' },
    { key: 'hasShape',          header: 'Use SHAPE',         type: 'dropdown', options: BOOL_OPTIONS },
    { key: 'diameter',          header: 'Diameter',          type: 'number' },
    { key: 'celerity',          header: 'Wave Speed',        type: 'number' },
    { key: 'friction',          header: 'Friction',          type: 'number' },
    { key: 'hasAddedLoss',      header: 'Added Loss',        type: 'dropdown', options: BOOL_OPTIONS },
    { key: 'cplus',             header: 'CPLUS',             type: 'number' },
    { key: 'cminus',            header: 'CMINUS',            type: 'number' },
    { key: 'comment',           header: 'Comment',           type: 'text' },
  ],
  flowBoundary: [
    { key: 'label',          header: 'Label',           type: 'text' },
    { key: '_unit',          header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber',     header: 'Node #',          type: 'number' },
    { key: 'scheduleNumber', header: 'Q Sched #',       type: 'number' },
    { key: 'comment',        header: 'Comment',         type: 'text' },
  ],
  pump: [
    { key: 'label',      header: 'Label',           type: 'text' },
    { key: '_unit',      header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber', header: 'Node #',          type: 'number' },
    { key: 'elevation',  header: 'Elevation',       type: 'number' },
    { key: 'pumpStatus', header: 'Status',          type: 'dropdown', options: PUMP_STATUS_OPTIONS },
    { key: 'pumpType',   header: 'PCHAR Type #',    type: 'number' },
    { key: 'rq',         header: 'RQ',              type: 'number' },
    { key: 'rhead',      header: 'RHEAD',           type: 'number' },
    { key: 'rspeed',     header: 'RSPEED (RPM)',    type: 'number' },
    { key: 'rtorque',    header: 'RTORQUE',         type: 'number' },
    { key: 'wr2',        header: 'WR²',             type: 'number' },
    { key: 'comment',    header: 'Comment',         type: 'text' },
  ],
  checkValve: [
    { key: 'label',       header: 'Label',           type: 'text' },
    { key: '_unit',       header: 'Unit (SI/FPS)',   type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber',  header: 'Node #',          type: 'number' },
    { key: 'elevation',   header: 'Elevation',       type: 'number' },
    { key: 'valveStatus', header: 'Status',          type: 'dropdown', options: VALVE_STATUS_OPTIONS },
    { key: 'valveDiam',   header: 'Valve Diameter',  type: 'number' },
    { key: 'comment',     header: 'Comment',         type: 'text' },
  ],
  turbine: [
    { key: 'label',           header: 'Label',               type: 'text' },
    { key: '_unit',           header: 'Unit (SI/FPS)',       type: 'dropdown', options: ['SI', 'FPS'] },
    { key: 'nodeNumber',      header: 'Node #',              type: 'number' },
    { key: 'elevation',       header: 'Elevation',           type: 'number' },
    { key: 'turbineType',     header: 'TCHAR Type #',        type: 'number' },
    { key: 'syncSpeed',       header: 'Sync Speed (RPM)',    type: 'number' },
    { key: 'turbineDiameter', header: 'Diameter',            type: 'number' },
    { key: 'wr2',             header: 'WR²',                 type: 'number' },
    { key: 'turbFriction',    header: 'Friction',            type: 'number' },
    { key: 'windage',         header: 'Windage',             type: 'number' },
    { key: 'operationMode',   header: 'Mode',                type: 'dropdown', options: TURBINE_MODE_OPTIONS },
    { key: 'vScheduleNumber', header: 'V Sched #',           type: 'number' },
    { key: 'comment',         header: 'Comment',             type: 'text' },
  ],
};

// ─── Material label ↔ id helpers ─────────────────────────────────────────────

import { PIPE_MATERIALS, PIPE_MATERIALS_BY_ID } from './pipe-materials';

function materialLabelById(id: number | string | undefined): string {
  if (!id) return '-- None --';
  const m = PIPE_MATERIALS_BY_ID[Number(id)];
  return m ? m.label : '-- None --';
}

function materialIdByLabel(label: string): number | undefined {
  if (!label || label === '-- None --') return undefined;
  const m = PIPE_MATERIALS.find(m => m.label === label);
  return m?.id;
}

// ─── Row value extractor ──────────────────────────────────────────────────────

function getRowValue(col: ColDef, data: Record<string, any>, kind: 'edge' | 'node', subType: string, globalUnit: string): string | number {
  if (col.key === '_type') {
    const TYPE_LABELS: Record<string, string> = {
      reservoir: 'Reservoir', node: 'Node', junction: 'Junction',
      surgeTank: 'Surge Tank', flowBoundary: 'Flow BC',
      pump: 'Pump', checkValve: 'Check Valve', turbine: 'Turbine',
      conduit: 'Conduit', dummy: 'Dummy Pipe',
    };
    return TYPE_LABELS[subType] ?? subType;
  }
  if (col.key === '_unit') {
    return (data.unit as string) || globalUnit;
  }
  if (col.key === '_materialLabel') {
    return materialLabelById(data.materialId);
  }
  const val = data[col.key];
  if (val === undefined || val === null) return '';
  if (typeof val === 'boolean') return val ? 'true' : 'false';
  return val as string | number;
}

// ─── Excel colour constants ───────────────────────────────────────────────────

const HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A73E8' },
};
const DROPDOWN_FILL: ExcelJS.Fill = {
  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF4FF' },
};
const HEADER_FONT: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };

// ─── Export ──────────────────────────────────────────────────────────────────

export interface ExportRow {
  id: string;
  kind: 'edge' | 'node';
  subType: string;
  data: Record<string, any>;
}

export async function exportTabToExcel(
  filter: FilterKey,
  rows: ExportRow[],
  globalUnit: string,
  tabLabel: string,
): Promise<void> {
  const cols = TAB_COLS[filter];
  if (!cols) throw new Error(`Unknown filter: ${filter}`);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'WHAMO Network Designer';
  wb.created = new Date();

  const ws = wb.addWorksheet(tabLabel, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  // ── Set column widths and headers ──
  ws.columns = cols.map(col => ({
    header: col.header,
    key: col.key,
    width: Math.max(16, col.header.length + 4),
  }));

  // Style header row
  const headerRow = ws.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = {
      bottom: { style: 'thin', color: { argb: 'FF1557B0' } },
    };
  });
  headerRow.height = 20;

  // ── Add data rows ──
  rows.forEach((row, rowIdx) => {
    const excelRowNum = rowIdx + 2;
    const values: (string | number)[] = cols.map(col =>
      getRowValue(col, row.data, row.kind, row.subType, globalUnit)
    );
    const excelRow = ws.addRow(values);

    // Alternate row shading
    const bgColor = rowIdx % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFF';
    excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const colDef = cols[colNumber - 1];
      if (!colDef) return;
      const isReadOnly = colDef.readOnly;
      cell.fill = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: colDef.type === 'dropdown' && !isReadOnly ? 'FFEEF4FF' : bgColor },
      };
      cell.font = { size: 10, color: { argb: isReadOnly ? 'FF888888' : 'FF1A1A2E' } };
      cell.border = {
        bottom: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        right: { style: 'hair', color: { argb: 'FFD1D5DB' } },
      };
      cell.alignment = { vertical: 'middle' };

      // Add dropdown validation
      if (colDef.type === 'dropdown' && colDef.options && !isReadOnly) {
        const formulaList = '"' + colDef.options.join(',') + '"';
        cell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [formulaList],
          showDropDown: false,
          error: 'Please select a value from the list.',
          errorTitle: 'Invalid Value',
          showErrorMessage: true,
        };
      }
    });
    excelRow.height = 18;
  });

  // ── Add a "Legend" sheet with field descriptions ──
  const legendWs = wb.addWorksheet('Legend (do not edit)');
  legendWs.columns = [
    { header: 'Column', key: 'col', width: 28 },
    { header: 'Description', key: 'desc', width: 45 },
    { header: 'Allowed Values', key: 'vals', width: 50 },
  ];
  const legendHeader = legendWs.getRow(1);
  legendHeader.eachCell(cell => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
  });
  legendHeader.height = 20;
  cols.forEach(col => {
    legendWs.addRow({
      col: col.header,
      desc: col.readOnly ? '(read-only, do not edit)' : col.type === 'dropdown' ? 'Select from dropdown' : col.type,
      vals: col.options ? col.options.join(', ') : col.type === 'number' ? 'Numeric value' : 'Free text',
    });
  });

  // ── Download ──
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `whamo_${filter}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ─── Import ───────────────────────────────────────────────────────────────────

export interface ImportUpdate {
  id: string;
  kind: 'edge' | 'node';
  data: Record<string, any>;
}

export async function importTabFromExcel(
  filter: FilterKey,
  rows: ExportRow[],
  globalUnit: string,
  file: File,
): Promise<{ updates: ImportUpdate[]; skipped: number; matched: number }> {
  const cols = TAB_COLS[filter];
  if (!cols) throw new Error(`Unknown filter: ${filter}`);

  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  // Find the first non-legend sheet
  const ws = wb.worksheets.find(s => !s.name.toLowerCase().includes('legend'));
  if (!ws) throw new Error('No data sheet found in workbook.');

  // Read header row to build column index
  const headerRow = ws.getRow(1);
  const headerMap: Record<string, number> = {};
  headerRow.eachCell((cell, colNum) => {
    const val = String(cell.value ?? '').trim();
    if (val) headerMap[val] = colNum;
  });

  // Build a lookup from label → ExportRow
  const labelLookup = new Map<string, ExportRow>();
  for (const row of rows) {
    const lbl = String(row.data.label ?? '').trim();
    if (lbl) labelLookup.set(lbl, row);
  }

  const updates: ImportUpdate[] = [];
  let skipped = 0;
  let matched = 0;

  // Find label column index
  const labelColDef = cols.find(c => c.key === 'label');
  const labelHeader = labelColDef?.header ?? 'Label';
  const labelColNum = headerMap[labelHeader];

  ws.eachRow((excelRow, rowNum) => {
    if (rowNum === 1) return; // skip header

    const labelCell = labelColNum ? excelRow.getCell(labelColNum) : null;
    const labelVal = String(labelCell?.value ?? '').trim();
    if (!labelVal) { skipped++; return; }

    const existingRow = labelLookup.get(labelVal);
    if (!existingRow) { skipped++; return; }

    matched++;
    const update: Record<string, any> = {};

    cols.forEach(col => {
      if (col.readOnly || col.key === 'label') return;
      const colNum = headerMap[col.header];
      if (!colNum) return;
      const cell = excelRow.getCell(colNum);
      const rawVal = cell.value;
      if (rawVal === null || rawVal === undefined || rawVal === '') return;
      const strVal = String(rawVal).trim();
      if (!strVal) return;

      if (col.key === '_unit') {
        if (strVal === 'SI' || strVal === 'FPS') update.unit = strVal;
        return;
      }
      if (col.key === '_materialLabel') {
        const matId = materialIdByLabel(strVal);
        if (matId !== undefined) {
          update.materialId = matId;
          const m = PIPE_MATERIALS_BY_ID[matId];
          if (m) {
            update.manningsN = m.manningsN;
            const unit: string = (existingRow.data.unit as string) || globalUnit;
            const eVal = unit === 'SI' ? m.youngsModulus_Pa : m.youngsModulus_psi;
            if (eVal > 0) update.pipeE = eVal;
          }
        } else {
          update.materialId = '';
        }
        return;
      }
      if (col.type === 'number') {
        const num = parseFloat(strVal);
        if (!isNaN(num)) update[col.key] = num;
        return;
      }
      if (col.type === 'dropdown' || col.type === 'text') {
        // Handle boolean fields stored as dropdown
        if (col.options && (col.options[0] === 'true' || col.options[0] === 'false')) {
          update[col.key] = strVal === 'true';
          return;
        }
        update[col.key] = strVal;
        return;
      }
    });

    updates.push({ id: existingRow.id, kind: existingRow.kind, data: update });
  });

  return { updates, skipped, matched };
}
