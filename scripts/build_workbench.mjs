#!/usr/bin/env node
/** Build the persistent MATIC detailed due-diligence workbench. */

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const skillRoot = path.resolve(scriptDir, "..");
const defaultOutput = path.join(skillRoot, "assets", "matic-detailed-dd-workbench.xlsx");
const outputPath = process.argv[2] ? path.resolve(process.argv[2]) : defaultOutput;
const previewDir = process.argv[3] ? path.resolve(process.argv[3]) : null;
const execFileAsync = promisify(execFile);

const DATA_START = 5;
const DATA_END = 204;
const FONT = "Microsoft YaHei";
const COLORS = {
  navy: "#1F4E78",
  blue: "#D9EAF7",
  paleBlue: "#EEF5FA",
  border: "#B4C7DC",
  white: "#FFFFFF",
  text: "#1F1F1F",
  muted: "#666666",
  green: "#E2F0D9",
  greenText: "#375623",
  yellow: "#FFF2CC",
  yellowText: "#7F6000",
  red: "#F4CCCC",
  redText: "#9C0006",
  gray: "#E7E6E6",
  grayText: "#595959",
};


function columnName(index) {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    value -= 1;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name;
}


function applyBaseSheet(sheet, title, note, headers, widths, tableName) {
  const lastCol = columnName(headers.length - 1);
  sheet.showGridLines = false;
  sheet.mergeCells(`A1:${lastCol}1`);
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${lastCol}1`).format = {
    fill: COLORS.navy,
    font: { name: FONT, size: 16, bold: true, color: COLORS.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  sheet.getRange("A1").format.rowHeight = 30;

  sheet.mergeCells(`A2:${lastCol}2`);
  sheet.getRange("A2").values = [[note]];
  sheet.getRange(`A2:${lastCol}2`).format = {
    fill: COLORS.paleBlue,
    font: { name: FONT, size: 9, color: COLORS.muted },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("A2").format.rowHeight = 38;

  const values = [headers, ...Array.from({ length: DATA_END - DATA_START + 1 }, () => Array(headers.length).fill(null))];
  sheet.getRange(`A4:${lastCol}${DATA_END}`).values = values;
  sheet.getRange(`A4:${lastCol}4`).format = {
    fill: COLORS.navy,
    font: { name: FONT, size: 10, bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  sheet.getRange(`A5:${lastCol}${DATA_END}`).format = {
    font: { name: FONT, size: 9, color: COLORS.text },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };
  sheet.getRange(`A4:${lastCol}${DATA_END}`).format.rowHeight = 32;
  sheet.getRange(`A4:${lastCol}4`).format.rowHeight = 42;
  widths.forEach((width, index) => {
    const col = columnName(index);
    sheet.getRange(`${col}4:${col}${DATA_END}`).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);

  const table = sheet.tables.add(`A4:${lastCol}${DATA_END}`, true, tableName);
  table.style = "TableStyleMedium2";
  table.showFilterButton = true;
  return sheet;
}


function addListValidation(sheet, column, values) {
  sheet.getRange(`${column}${DATA_START}:${column}${DATA_END}`).dataValidation = {
    rule: { type: "list", values },
  };
}


function addStatusFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  for (const text of ["已确认", "已完成", "已访谈", "通过", "可生成最终报告"]) {
    range.conditionalFormats.add("containsText", {
      text,
      format: { fill: COLORS.green, font: { color: COLORS.greenText, bold: true } },
    });
  }
  for (const text of ["待", "争议", "处理中", "有条件", "建议阶段性报告", "接近最终成文"]) {
    range.conditionalFormats.add("containsText", {
      text,
      format: { fill: COLORS.yellow, font: { color: COLORS.yellowText } },
    });
  }
  for (const text of ["重大", "触发", "阻断", "与原判断矛盾"]) {
    range.conditionalFormats.add("containsText", {
      text,
      format: { fill: COLORS.red, font: { color: COLORS.redText, bold: true } },
    });
  }
  for (const text of ["已合并", "已推翻", "已失效", "不适用", "不纳入"]) {
    range.conditionalFormats.add("containsText", {
      text,
      format: { fill: COLORS.gray, font: { color: COLORS.grayText } },
    });
  }
}


function formatDateColumns(sheet, columns) {
  for (const column of columns) {
    sheet.getRange(`${column}${DATA_START}:${column}${DATA_END}`).format.numberFormat = "yyyy-mm-dd";
  }
}


async function ensureFrozenPanes(xlsxPath) {
  // artifact-tool records freeze intent, but some exporters omit the pane XML.
  // Patch only this unsupported presentation detail; workbook content remains artifact-tool-authored.
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "matic-freeze-"));
  const unpacked = path.join(tempRoot, "xlsx");
  const patched = path.join(tempRoot, "patched.xlsx");
  await fs.mkdir(unpacked, { recursive: true });
  try {
    await execFileAsync("unzip", ["-q", xlsxPath, "-d", unpacked]);
    for (let index = 1; index <= 10; index += 1) {
      const rows = index === 1 ? 2 : 4;
      const xmlPath = path.join(unpacked, "xl", "worksheets", `sheet${index}.xml`);
      let xml = await fs.readFile(xmlPath, "utf8");
      if (xml.includes("<x:pane")) continue;
      const pane = `<x:pane ySplit="${rows}" topLeftCell="A${rows + 1}" activePane="bottomLeft" state="frozen" />`;
      if (/<x:sheetView\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/<x:sheetView\b([^>]*)\/>/, `<x:sheetView$1>${pane}</x:sheetView>`);
      } else {
        xml = xml.replace("</x:sheetView>", `${pane}</x:sheetView>`);
      }
      await fs.writeFile(xmlPath, xml, "utf8");
    }
    await execFileAsync("zip", ["-qr", patched, "."], { cwd: unpacked });
    await fs.copyFile(patched, xlsxPath);
  } finally {
    await fs.rm(tempRoot, { recursive: true, force: true });
  }
}


const workbook = Workbook.create();
const home = workbook.worksheets.add("首页");
const candidate = workbook.worksheets.add("调研对象候选库");
const interview = workbook.worksheets.add("联系与访谈进度");
const material = workbook.worksheets.add("材料及专家报告索引");
const competition = workbook.worksheets.add("竞品与替代方案矩阵");
const issue = workbook.worksheets.add("问题跟踪表");
const conclusion = workbook.worksheets.add("结论登记表");
const dispute = workbook.worksheets.add("专家分歧表");
const selection = workbook.worksheets.add("第一部分入选结论");
const updateLog = workbook.worksheets.add("更新日志");


applyBaseSheet(
  candidate,
  "调研对象候选库",
  "启动时建立30—40名或家以上候选对象，优先江浙沪和长三角地区；实名对象与机构/科室均可，禁止推断私人联系方式。",
  ["对象编号", "主类别", "辅助标签", "对象类型", "姓名或机构", "单位、职务或建议访谈角色", "地区", "长三角优先", "匹配理由", "拟验证问题", "优先级", "与项目方关系或利益冲突", "公开来源", "公开联系路径", "当前状态", "关联问题编号", "关联结论编号", "备注"],
  [11, 22, 18, 13, 22, 26, 13, 12, 30, 30, 9, 28, 34, 28, 14, 17, 18, 24],
  "CandidatePool",
);
addListValidation(candidate, "B", ["行业专家", "下游客户", "上游（设备）厂家", "竞争对手", "核心团队", "其它（如投资专家、知识产权专家等）"]);
addListValidation(candidate, "D", ["个人", "医院科室", "高校/科研机构", "企业", "投资机构", "专业服务机构", "其他"]);
addListValidation(candidate, "H", ["是", "否"]);
addListValidation(candidate, "K", ["A", "B", "C"]);
addListValidation(candidate, "O", ["待筛选", "拟联系", "已联系", "已安排", "已访谈", "暂缓", "不适用"]);
addStatusFormatting(candidate, `O${DATA_START}:O${DATA_END}`);


applyBaseSheet(
  interview,
  "联系与访谈进度",
  "访谈意见须经用户确认后才能正式支撑详细报告第一部分第2—5节；访谈编号使用INT-01起顺序编号。",
  ["访谈编号", "对象编号", "主类别", "姓名或机构", "计划访谈角色", "联系状态", "计划日期", "实际日期", "访谈方式", "访谈负责人", "核心问题", "访谈材料编号", "用户确认状态", "形成或影响的结论编号", "后续行动", "备注"],
  [12, 11, 22, 22, 20, 14, 13, 13, 12, 14, 32, 17, 14, 22, 28, 22],
  "InterviewProgress",
);
addListValidation(interview, "C", ["行业专家", "下游客户", "上游（设备）厂家", "竞争对手", "核心团队", "其它（如投资专家、知识产权专家等）"]);
addListValidation(interview, "F", ["未联系", "拟联系", "已联系", "已安排", "已完成", "取消", "暂缓"]);
addListValidation(interview, "I", ["线下", "电话", "视频", "书面", "其他"]);
addListValidation(interview, "M", ["未提交", "待确认", "已确认", "不纳入"]);
addStatusFormatting(interview, `F${DATA_START}:F${DATA_END}`);
addStatusFormatting(interview, `M${DATA_START}:M${DATA_END}`);
formatDateColumns(interview, ["G", "H"]);


applyBaseSheet(
  material,
  "材料及专家报告索引",
  "先登记来源再分析。录音不在本Skill中处理；专家输入以用户确认的专家意见表为主，也可登记访谈文字稿或笔记。",
  ["材料编号", "材料类型", "文件或材料名称", "版本或日期", "提供方或作者", "对应访谈编号", "收到日期", "用户确认状态", "证据性质", "主要内容", "影响的问题编号", "影响的结论编号", "来源位置或链接", "保密或使用限制", "备注"],
  [12, 20, 30, 16, 22, 16, 13, 14, 16, 32, 18, 20, 36, 22, 22],
  "MaterialIndex",
);
addListValidation(material, "B", ["标准尽调报告", "BP", "项目方回复", "专家意见表", "访谈文字稿", "实验数据", "第三方检测", "知识产权", "股权与工商", "注册与临床", "市场与客户", "公开资料", "其他"]);
addListValidation(material, "H", ["未确认", "待确认", "已确认", "不纳入"]);
addListValidation(material, "I", ["项目方陈述", "专家判断", "原始数据", "第三方证据", "公开事实", "AI推断"]);
addStatusFormatting(material, `H${DATA_START}:H${DATA_END}`);
formatDateColumns(material, ["G"]);


applyBaseSheet(
  competition,
  "竞品与替代方案矩阵",
  "每个“比较对象×对比维度”单独一行并沿用稳定CMP编号；启动时继承标准尽调结果，新增材料须主动识别新对象、新路线和新优劣势。",
  ["对比编号", "对比对象", "研发或供应主体", "对象分类", "产品/技术阶段", "目标适应证或场景", "对比维度", "本项目相对结论", "具体判断", "数据比较条件", "可比性", "证据性质", "证据强度", "专家支持", "专家反对或限制", "对竞争地位的影响", "持续性或可弥补性", "下一步验证动作", "判定标准", "关联问题编号", "关联材料或访谈编号", "关联结论编号", "当前状态", "最后更新日期", "备注"],
  [12, 24, 22, 18, 18, 24, 18, 16, 34, 30, 15, 22, 12, 30, 30, 18, 20, 30, 28, 18, 24, 20, 16, 15, 24],
  "CompetitionMatrix",
);
addListValidation(competition, "D", ["直接竞品", "间接竞品", "医院现行方案", "潜在替代技术"]);
addListValidation(competition, "H", ["相对优势", "相对劣势", "基本相当", "暂无法判断"]);
addListValidation(competition, "K", ["直接可比", "有限可比", "不可直接比较"]);
addListValidation(competition, "L", ["直接头对头数据", "可比公开数据", "跨研究间接比较", "项目方主张", "专家判断", "暂无可靠证据"]);
addListValidation(competition, "M", ["强", "中", "弱"]);
addListValidation(competition, "P", ["重大正面", "正面", "中性", "负面", "重大负面", "待判断"]);
addListValidation(competition, "Q", ["结构性优势", "阶段性领先", "容易追赶", "可弥补劣势", "难弥补劣势", "待判断", "不适用"]);
addListValidation(competition, "W", ["草拟", "待补证", "基本确认", "已确认", "仍有争议", "已替代", "已失效"]);
addStatusFormatting(competition, `W${DATA_START}:W${DATA_END}`);
formatDateColumns(competition, ["X"]);


applyBaseSheet(
  issue,
  "问题跟踪表",
  "问题可随尽调过程持续新增；风险事项在本表同时记录证据、严重性、发生可能性、触发等级和解除条件。",
  ["问题编号", "问题分类", "核心问题或风险", "标准尽调判断", "新材料或专家输入", "主要证据", "当前判断", "当前状态", "严重性", "发生可能性", "触发等级", "责任对象", "下一步验证动作", "所需材料或访谈对象", "解除或验证条件", "截止或复核日期", "关联材料编号", "关联结论编号", "最后更新日期", "备注"],
  [11, 20, 32, 28, 32, 30, 32, 18, 12, 13, 16, 18, 30, 28, 28, 15, 18, 20, 15, 22],
  "IssueTracker",
);
addListValidation(issue, "B", ["技术", "市场", "团队", "注册与临床", "知识产权与权属", "股权与商业关联", "产业链与产业化", "政策与适应证", "其他"]);
addListValidation(issue, "H", ["已确认", "基本确认", "仍有争议", "待补证", "与原判断矛盾", "已解除风险"]);
addListValidation(issue, "I", ["重大", "高", "中", "低"]);
addListValidation(issue, "J", ["高", "中", "低", "未知"]);
addListValidation(issue, "K", ["未触发", "疑似触发", "重大前置条件", "确认触发", "不适用"]);
addStatusFormatting(issue, `H${DATA_START}:K${DATA_END}`);
formatDateColumns(issue, ["P", "S"]);


applyBaseSheet(
  conclusion,
  "结论登记表",
  "结论编号永久保留；新增材料既要映射既有结论，也要主动发现新结论。合并、推翻或失效时更新状态并保留历史。",
  ["结论编号", "所属板块", "结论类型", "当前结论", "适用范围或前提", "专家支持", "专家反对或限制", "客观证据", "证据强度", "当前状态", "决策重要性", "第一部分候选", "关联问题编号", "关联材料或访谈编号", "首次形成日期", "最后更新日期", "变更说明", "合并至或替代编号"],
  [14, 14, 11, 34, 26, 30, 30, 30, 12, 14, 12, 14, 18, 24, 15, 15, 28, 20],
  "ConclusionRegister",
);
addListValidation(conclusion, "B", ["技术", "市场", "核心团队", "其他"]);
addListValidation(conclusion, "C", ["优势", "风险"]);
addListValidation(conclusion, "I", ["强", "中", "弱"]);
addListValidation(conclusion, "J", ["草拟", "待补证", "基本确认", "已确认", "仍有争议", "已合并", "已推翻", "已失效"]);
addListValidation(conclusion, "K", ["高", "中", "低"]);
addListValidation(conclusion, "L", ["是", "否", "待定"]);
addStatusFormatting(conclusion, `J${DATA_START}:J${DATA_END}`);
formatDateColumns(conclusion, ["O", "P"]);


applyBaseSheet(
  dispute,
  "专家分歧表",
  "分歧不得按简单多数票关闭；需结合专业范围、信息基础、利益关系、适用条件和客观证据判断。",
  ["分歧编号", "关联问题或结论编号", "分歧主题", "支持观点及专家", "反对观点及专家", "分歧来源编号", "可能原因或适用条件", "对项目判断的影响", "补充调研方案", "当前状态", "解决依据", "最后更新日期"],
  [12, 22, 28, 32, 32, 20, 30, 30, 28, 14, 30, 15],
  "ExpertDisputes",
);
addListValidation(dispute, "J", ["新建", "待补访", "部分澄清", "已解决", "保留分歧"]);
addStatusFormatting(dispute, `J${DATA_START}:J${DATA_END}`);
formatDateColumns(dispute, ["L"]);


applyBaseSheet(
  selection,
  "第一部分入选结论",
  "只有专家证据门槛和跨板块去重检查均通过的结论，才能进入详细报告第一部分第2—5节；成文时不显示结论编号。",
  ["显示板块", "显示顺序", "结论编号", "优势或风险", "第一部分表述", "专家支持括注", "决策重要性", "入选理由", "证据门槛检查", "去重检查", "当前状态", "最后更新日期"],
  [16, 11, 14, 13, 38, 34, 13, 28, 18, 14, 14, 15],
  "PartOneSelection",
);
addListValidation(selection, "A", ["技术方面", "市场方面", "核心团队", "其他方面"]);
addListValidation(selection, "D", ["优势", "风险"]);
addListValidation(selection, "G", ["高", "中", "低"]);
addListValidation(selection, "I", ["通过", "不通过", "待补证"]);
addListValidation(selection, "J", ["通过", "不通过", "待调整"]);
addListValidation(selection, "K", ["草拟", "拟入选", "已入选", "暂不入选", "已替换"]);
addStatusFormatting(selection, `I${DATA_START}:K${DATA_END}`);
formatDateColumns(selection, ["L"]);


applyBaseSheet(
  updateLog,
  "更新日志",
  "每次新增材料、专家意见、公开信息刷新或成文均追加记录。Skill可建议成文，但只有用户明确指令才能生成报告。",
  ["更新编号", "更新日期", "更新类型", "新增材料编号", "影响工作表", "新增内容", "修正、合并或推翻内容", "对当前判断的影响", "是否建议成文", "操作人或说明"],
  [12, 15, 18, 20, 22, 34, 34, 32, 20, 24],
  "UpdateLog",
);
addListValidation(updateLog, "C", ["启动", "材料更新", "专家更新", "公开信息刷新", "阶段性成文", "最终成文", "其他"]);
addListValidation(updateLog, "I", ["否", "建议阶段性报告", "接近最终成文", "可生成最终报告"]);
addStatusFormatting(updateLog, `I${DATA_START}:I${DATA_END}`);
formatDateColumns(updateLog, ["B"]);


home.showGridLines = false;
home.mergeCells("A1:H1");
home.getRange("A1").values = [["MATIC 项目详细尽调工作台"]];
home.getRange("A1:H1").format = {
  fill: COLORS.navy,
  font: { name: FONT, size: 18, bold: true, color: COLORS.white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
home.getRange("A1").format.rowHeight = 36;
home.mergeCells("A2:H2");
home.getRange("A2").values = [["持续维护问题、证据、访谈、结论与分歧；日常更新不改文件名，阶段性或最终成文时另存快照。"]];
home.getRange("A2:H2").format = {
  fill: COLORS.paleBlue,
  font: { name: FONT, size: 10, color: COLORS.muted },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
home.getRange("A4:B4").values = [["项目信息", "内容"]];
home.getRange("A5:B11").values = [
  ["项目名称", "[项目名称]"],
  ["项目简称", "[项目简称]"],
  ["当前阶段", "启动详细尽调"],
  ["工作台版本", "V0.1"],
  ["创建日期", "[YYYY-MM-DD]"],
  ["最后更新日期", "[YYYY-MM-DD]"],
  ["当前尽调建议", "[待填写]"],
];
home.getRange("D4:E4").values = [["进展指标", "当前数量"]];
home.getRange("D5:D11").values = [["候选对象"], ["已完成访谈"], ["已登记材料"], ["竞品对比记录"], ["开放问题"], ["有效结论"], ["重大风险"]];
home.getRange("E5:E11").formulas = [
  ["=COUNTA('调研对象候选库'!$A$5:$A$204)"],
  ["=COUNTIF('联系与访谈进度'!$F$5:$F$204,\"已完成\")"],
  ["=COUNTA('材料及专家报告索引'!$A$5:$A$204)"],
  ["=COUNTA('竞品与替代方案矩阵'!$A$5:$A$204)"],
  ["=COUNTIF('问题跟踪表'!$H$5:$H$204,\"<>\")-COUNTIF('问题跟踪表'!$H$5:$H$204,\"已确认\")-COUNTIF('问题跟踪表'!$H$5:$H$204,\"已解除风险\")"],
  ["=COUNTIFS('结论登记表'!$A$5:$A$204,\"<>\",'结论登记表'!$J$5:$J$204,\"<>已合并\",'结论登记表'!$J$5:$J$204,\"<>已推翻\",'结论登记表'!$J$5:$J$204,\"<>已失效\")"],
  ["=COUNTIF('问题跟踪表'!$I$5:$I$204,\"重大\")"],
];
home.getRange("A4:B11").format = {
  font: { name: FONT, size: 10, color: COLORS.text },
  wrapText: true,
  borders: { preset: "all", style: "thin", color: COLORS.border },
  verticalAlignment: "center",
};
home.getRange("D4:E11").format = {
  font: { name: FONT, size: 10, color: COLORS.text },
  wrapText: true,
  borders: { preset: "all", style: "thin", color: COLORS.border },
  verticalAlignment: "center",
};
home.getRange("A4:B4").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: COLORS.white }, horizontalAlignment: "center" };
home.getRange("D4:E4").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: COLORS.white }, horizontalAlignment: "center" };
home.getRange("A5:A11").format = { fill: COLORS.blue, font: { name: FONT, size: 10, bold: true, color: COLORS.text } };
home.getRange("D5:D11").format = { fill: COLORS.blue, font: { name: FONT, size: 10, bold: true, color: COLORS.text } };
home.getRange("E5:E11").format.numberFormat = "0";

home.getRange("A13:D13").merge();
home.getRange("A13").values = [["状态色说明"]];
home.getRange("A13:D13").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: COLORS.white } };
home.getRange("A14:D17").values = [
  ["绿色", "已确认/已完成", "可进入报告", "继续维护证据链"],
  ["黄色", "待补证/仍有争议", "暂不定论", "安排补访或材料核验"],
  ["红色", "重大风险/确认触发", "可能阻断推进", "明确事实、规则和解除条件"],
  ["灰色", "已合并/已推翻/不适用", "保留历史", "不得删除或复用编号"],
];
home.getRange("A14:D17").format = { font: { name: FONT, size: 10, color: COLORS.text }, borders: { preset: "all", style: "thin", color: COLORS.border }, wrapText: true };
home.getRange("A14:D14").format.fill = COLORS.green;
home.getRange("A15:D15").format.fill = COLORS.yellow;
home.getRange("A16:D16").format.fill = COLORS.red;
home.getRange("A17:D17").format.fill = COLORS.gray;

home.getRange("A19:H19").merge();
home.getRange("A19").values = [["使用提示"]];
home.getRange("A19:H19").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: COLORS.white } };
home.getRange("A20:H24").merge(true);
home.getRange("A20:A24").values = [
  ["1. 启动详细尽调时先建立问题、候选对象和材料计划，不得在无专家证据时强行形成稳定结论。"],
  ["2. 每份新材料先登记索引，再更新问题、结论、分歧和第一部分入选结论。"],
  ["3. 竞品矩阵每行记录一个对象与一个维度；每次更新均检查新对象、新路线、新优劣势及比较条件变化。"],
  ["4. 阶段性或最终报告只在明确指令后生成；成文前检查专家证据门槛、去重和结论历史。"],
  ["5. 本工作台不处理录音转写，不设置固定市场规模工作表；市场重大变化使用现有清单记录。"],
];
home.getRange("A20:H24").format = { fill: COLORS.paleBlue, font: { name: FONT, size: 10, color: COLORS.text }, wrapText: true, borders: { preset: "all", style: "thin", color: COLORS.border } };
home.getRange("A:A").format.columnWidth = 18;
home.getRange("B:B").format.columnWidth = 30;
home.getRange("C:C").format.columnWidth = 8;
home.getRange("D:D").format.columnWidth = 20;
home.getRange("E:E").format.columnWidth = 15;
home.getRange("F:H").format.columnWidth = 14;
home.getRange("A4:H24").format.rowHeight = 28;
home.freezePanes.freezeRows(2);


await fs.mkdir(path.dirname(outputPath), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
await ensureFrozenPanes(outputPath);

if (previewDir) {
  await fs.mkdir(previewDir, { recursive: true });
  for (const sheetName of ["首页", "调研对象候选库", "联系与访谈进度", "材料及专家报告索引", "竞品与替代方案矩阵", "问题跟踪表", "结论登记表", "专家分歧表", "第一部分入选结论", "更新日志"]) {
    const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
    const safeName = sheetName.replaceAll("/", "_");
    await fs.writeFile(path.join(previewDir, `${safeName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
}

// The renderer may emit an inspection sidecar next to the workbook; it is QA output, not a Skill asset.
await fs.rm(`${outputPath}.inspect.ndjson`, { force: true });
process.stdout.write(`${outputPath}\n`);
