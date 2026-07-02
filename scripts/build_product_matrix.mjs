import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const dataPath = "outputs/product_matrix/product_matrix_data.json";
const outputDir = "outputs/product_matrix";
const outputPath = path.join(outputDir, "海外线下商超产品矩阵_按动物线.xlsx");
const raw = JSON.parse(await fs.readFile(dataPath, "utf8"));

function colName(n) {
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function writeMatrix(sheet, startRow, startCol, matrix) {
  if (!matrix.length || !matrix[0]?.length) return;
  sheet.getRangeByIndexes(startRow, startCol, matrix.length, matrix[0].length).values = matrix;
}

function styleTitle(sheet, title, cols) {
  sheet.showGridLines = false;
  sheet.getRangeByIndexes(0, 0, 1, cols).merge();
  sheet.getCell(0, 0).values = [[title]];
  const titleRange = sheet.getRangeByIndexes(0, 0, 1, cols);
  titleRange.format.fill.color = "#174A5A";
  titleRange.format.font.color = "#FFFFFF";
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 15;
  titleRange.format.rowHeight = 28;
}

function styleHeader(sheet, row, cols, fill = "#DCECEF") {
  const range = sheet.getRangeByIndexes(row, 0, 1, cols);
  range.format.fill.color = fill;
  range.format.font.bold = true;
  range.format.borders = { preset: "all", style: "thin", color: "#B7C9D1" };
}

function styleBody(sheet, startRow, rows, cols) {
  if (rows <= 0) return;
  const range = sheet.getRangeByIndexes(startRow, 0, rows, cols);
  range.format.wrapText = true;
  range.format.verticalAlignment = "top";
  range.format.borders = { preset: "inside", style: "thin", color: "#E6EEF2" };
  range.format.font.size = 10;
}

function setWidths(sheet, widths) {
  widths.forEach((w, i) => {
    sheet.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth = w;
  });
}

function channelSummary(rows) {
  const counts = new Map();
  for (const r of rows) counts.set(r["渠道定位"], (counts.get(r["渠道定位"]) ?? 0) + 1);
  return ["流通走量款", "中端利润款", "高端形象款"].map((c) => `${c} ${counts.get(c) ?? 0}`).join("；");
}

function matrixRowsFor(animal) {
  const subset = raw.detail.filter((r) => r["适用动物线"] === animal);
  const grouped = new Map();
  for (const r of subset) {
    const key = [r["一级大类"], r["二级细分品类"], r["渠道定位"]].join("|");
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(r);
  }
  return [...grouped.entries()]
    .sort((a, b) => a[0].localeCompare(b[0], "zh"))
    .map(([key, items]) => {
      const [main, sub, channel] = key.split("|");
      const skuList = items.map((r) => r["SKU 编码"]).join(", ");
      const names = items.map((r) => r["产品名称"]).join("；");
      const logic =
        channel === "流通走量款"
          ? "基础常备，适合整箱补货和高频补货"
          : channel === "中端利润款"
            ? "控制主题数量，作为利润补充陈列"
            : "少量陈列或样机，按客户订单补货";
      return [main, sub, channel, items.length, skuList, names, logic];
    });
}

function addAnimalSheet(wb, animal, sheetName) {
  const sheet = wb.worksheets.add(sheetName);
  const rows = raw.detail.filter((r) => r["适用动物线"] === animal);
  const summary = raw.animal_summary.find((r) => r["适用动物线"] === animal);
  const topHeaders = ["动物线", "SKU 数量", "流通走量款", "中端利润款", "高端形象款", "主要品类", "矩阵判断"];
  const topValues = [[
    animal,
    rows.length,
    summary?.["流通走量款"] ?? 0,
    summary?.["中端利润款"] ?? 0,
    summary?.["高端形象款"] ?? 0,
    summary?.["主要品类"] ?? "",
    summary?.["矩阵判断"] ?? "",
  ]];
  const headers = ["一级大类", "二级细分品类", "渠道定位", "SKU 数量", "SKU 编码", "产品名称", "线下铺货逻辑"];
  const matrix = matrixRowsFor(animal);
  styleTitle(sheet, `${animal} 产品矩阵`, headers.length);
  writeMatrix(sheet, 2, 0, [topHeaders]);
  writeMatrix(sheet, 3, 0, topValues);
  styleHeader(sheet, 2, topHeaders.length, "#EEF3D8");
  styleBody(sheet, 3, 1, topHeaders.length);
  writeMatrix(sheet, 6, 0, [headers]);
  writeMatrix(sheet, 7, 0, matrix);
  styleHeader(sheet, 6, headers.length);
  styleBody(sheet, 7, matrix.length, headers.length);
  setWidths(sheet, [20, 24, 15, 10, 44, 76, 38]);
  sheet.freezePanes.freezeRows(7);
}

const wb = Workbook.create();

const overview = wb.worksheets.add("总览");
styleTitle(overview, "海外线下商超产品矩阵总览（按适用动物线拆分）", 7);
const animalHeaders = ["适用动物线", "SKU 数量", "流通走量款", "中端利润款", "高端形象款", "主要品类", "矩阵判断"];
const animalRows = raw.animal_summary.map((r) => animalHeaders.map((h) => r[h] ?? ""));
writeMatrix(overview, 2, 0, [animalHeaders]);
writeMatrix(overview, 3, 0, animalRows);
styleHeader(overview, 2, animalHeaders.length);
styleBody(overview, 3, animalRows.length, animalHeaders.length);

const catHeaders = ["适用动物线", "一级大类", "二级细分品类", "SKU 数量", "均衡判断"];
const catRows = raw.cat_counts.map((r) => catHeaders.map((h) => r[h]));
writeMatrix(overview, 8, 0, [catHeaders]);
writeMatrix(overview, 9, 0, catRows);
styleHeader(overview, 8, catHeaders.length, "#EEF3D8");
styleBody(overview, 9, catRows.length, catHeaders.length);
setWidths(overview, [20, 22, 26, 10, 34, 34, 46]);
overview.freezePanes.freezeRows(3);

addAnimalSheet(wb, "仓鼠/小型啮齿类", "仓鼠小宠矩阵");
addAnimalSheet(wb, "兔/豚鼠类", "兔豚鼠矩阵");
addAnimalSheet(wb, "通用小宠耗材", "通用耗材矩阵");

const detail = wb.worksheets.add("标准明细");
const detailHeaders = ["适用动物线", "一级大类", "二级细分品类", "渠道定位", "SKU 编码", "产品名称", "产品核心属性", "库存铺货属性"];
const detailRows = raw.detail.map((r) => detailHeaders.map((h) => r[h] ?? ""));
styleTitle(detail, "标准化明细表格（日常备货 / SKU 管理 / 内部对账）", detailHeaders.length);
writeMatrix(detail, 1, 0, [detailHeaders]);
writeMatrix(detail, 2, 0, detailRows);
styleHeader(detail, 1, detailHeaders.length);
styleBody(detail, 2, detailRows.length, detailHeaders.length);
setWidths(detail, [18, 18, 22, 14, 18, 44, 58, 58]);
detail.freezePanes.freezeRows(2);

const redundancy = wb.worksheets.add("冗余精简");
const redHeaders = ["冗余组", "涉及 SKU", "建议保留", "建议精简/合并", "判断口径"];
const redRows = raw.redundancy.map((r) => redHeaders.map((h) => r[h]));
styleTitle(redundancy, "功能重叠与可精简 SKU", redHeaders.length);
writeMatrix(redundancy, 1, 0, [redHeaders]);
writeMatrix(redundancy, 2, 0, redRows);
styleHeader(redundancy, 1, redHeaders.length);
styleBody(redundancy, 2, redRows.length, redHeaders.length);
setWidths(redundancy, [32, 62, 42, 58, 44]);
redundancy.freezePanes.freezeRows(2);

const gaps = wb.worksheets.add("新品与铺货组合");
styleTitle(gaps, "缺口赛道、新品方向与经销商铺货组合", 5);
const gapHeaders = ["缺口赛道", "新品拓展方向", "线下铺货原因"];
writeMatrix(gaps, 2, 0, [gapHeaders]);
writeMatrix(gaps, 3, 0, raw.gaps);
styleHeader(gaps, 2, gapHeaders.length);
styleBody(gaps, 3, raw.gaps.length, gapHeaders.length);
const setHeaders = ["组合场景", "引流款", "利润款/形象款", "铺货逻辑"];
writeMatrix(gaps, 11, 0, [setHeaders]);
writeMatrix(gaps, 12, 0, raw.sets);
styleHeader(gaps, 11, setHeaders.length, "#EEF3D8");
styleBody(gaps, 12, raw.sets.length, setHeaders.length);
setWidths(gaps, [28, 54, 68, 62, 20]);

await fs.mkdir(outputDir, { recursive: true });
for (const sheetName of ["总览", "仓鼠小宠矩阵", "兔豚鼠矩阵", "通用耗材矩阵", "标准明细", "冗余精简", "新品与铺货组合"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName}_新版.png`), bytes);
}

const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(outputPath);

const inspect = await wb.inspect({ kind: "sheet", include: "id,name", maxChars: 2000 });
console.log(inspect.ndjson);
console.log(outputPath);
