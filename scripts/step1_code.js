const WORD_FILE_URL = atob("aHR0cHM6Ly93d3cua2RvY3MuY24vbC9jZDhxcjhXeFJRVFY=");
const LOG_FILE_URL  = atob("aHR0cHM6Ly93d3cua2RvY3MuY24vbC9jbWJ0c3ZveXh2dFM=");

let violationHitCount   = 0;
let brandDeleteInTitle  = 0;
let brandDeleteInDesc   = 0;
let vioHitInBrand       = 0;
let vioHitInDesc        = 0;
let vioHitInTitle       = 0;
let duplicateImageCount = 0;
let priceAdjustedCount  = 0;

function safeString(v) { return v == null ? "" : String(v); }
function updateProgress(pct, msg) { Application.StatusBar = pct + "% - " + msg; }

// ============================================================
// ★ 修复版 removeEmojiAndFixSpaces（不会误删参数）
// ============================================================
function removeEmojiAndFixSpaces(text) {
    if (!text) return text;

    text = text.replace(/[\p{Extended_Pictographic}]/gu, "");
    text = text.replace(/\s{3,}/g, " ");

    return text.trim();
}

// ============================================================
// ★ 修复版 cleanTitlePrefix（不破坏尺寸/型号）
// ============================================================
function cleanTitlePrefix(text) {
    if (!text) return text;
    return text.replace(/^\s+/, "");
}

const sheet    = Application.ActiveSheet;
const used     = sheet.UsedRange;
const colCount = used.Columns.Count;
const rowCount = used.Rows.Count;

let titleCol  = null;
let brandCol  = null;
let priceCol  = null;
let buyboxCol = null;
let eanCol    = null;
let descCols  = [];

function matchHeader(text, keywords) { return keywords.some(k => text.includes(k)); }

for (let j = 1; j <= colCount; j++) {
  const h = (sheet.Cells(1, j).Text || "").trim();
  if (!titleCol  && matchHeader(h, ["产品名称","商品名称","标题","Title","Name"])) titleCol  = j;
  if (!brandCol  && matchHeader(h, ["Brand","品牌"]))                               brandCol  = j;
  if (!priceCol  && matchHeader(h, ["价格","Price"]))                              priceCol  = j;
  if (!buyboxCol && matchHeader(h, ["BuyBox价格","BuyBox","Buy Box Price","BB价格","BuyBoxPrice"])) buyboxCol = j;
  if (!eanCol    && matchHeader(h, ["EAN","ean","条码","商品条码","EAN码","商品EAN","国际条码","条形码"])) eanCol = j;
  if (matchHeader(h, ["产品短描述","短描述","描述","Description","Desc"])) descCols.push(j);
}

updateProgress(10, "列名识别完成");

let violations = [];
try {
  const wf    = KSDrive.openFile(WORD_FILE_URL);
  const ws    = wf.Application.ActiveSheet;
  const wRows = ws.UsedRange.Rows.Count;
  for (let i = 1; i <= wRows; i++) {
    ["B","C","H"].forEach(col => {
      const v = safeString(ws.Range(col + i).Value2).trim();
      if (v) violations.push(v);
    });
  }
  wf.close();
} catch (e) {}

violations = Array.from(new Set(violations));
updateProgress(20, "违规词库加载完成");

function normalizeBrandWord(w) { return w ? String(w).trim().toLowerCase() : ""; }
let brandSet = new Set();
if (brandCol) {
  for (let i = 2; i <= rowCount; i++) {
    const raw = sheet.Cells(i, brandCol).Value2;
    if (!raw) continue;
    let b = normalizeBrandWord(raw);
    if (b.length >= 1) brandSet.add(b);
  }
}
const brandList = Array.from(brandSet);

function normalizeVioWord(w) { return w ? String(w).trim().toLowerCase() : ""; }
let vioSet = new Set();
for (let v of violations) {
    let n = normalizeVioWord(v);
    if (n.length >= 1) vioSet.add(n);
}
const vioList = Array.from(vioSet);

// ============================================================
// ★ 方式 1：前后字符判断法（最稳，不漏词）
// ============================================================
function isBoundaryChar(ch) {
    // 字母、数字、中文 → 非边界
    return !(/[a-z0-9\u4E00-\u9FFF]/i.test(ch));
}

// ============================================================
// ★ checkBrandHasViolation（方式 1 改写版）
// ============================================================
function checkBrandHasViolation(text) {
    if (!text) return 0;
    const t = text.toLowerCase();
    let count = 0;

    for (let w of vioList) {
        const word = w.toLowerCase();
        let idx = t.indexOf(word);

        while (idx !== -1) {
            const before = t[idx - 1] || "";
            const after  = t[idx + word.length] || "";

            if (isBoundaryChar(before) && isBoundaryChar(after)) {
                count++;
                break;
            }

            idx = t.indexOf(word, idx + 1);
        }
    }
    return count;
}

// ============================================================
// ★ cleanBrandV2（方式 1 改写版 + 参数保护）
// ============================================================
function cleanBrandV2(text, isTitle) {
    if (!text) return text;

    text = text.replace(/[\u200B-\u200F\u202A-\u202E\u2060\uFEFF]/g, "");
    text = text.replace(/&nbsp;|&#160;/gi, " ");
    text = text.replace(/&amp;|&#38;/gi, "&");
    text = text.replace(/&#(\d+);/g, (_, code) => String.fromCharCode(code));

    const full2half = s => s.replace(/[\uFF01-\uFF5E]/g, ch =>
        String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
    ).replace(/\u3000/g, " ");
    text = full2half(text);

    let out = text;
    let removed = 0;

    for (let bw of brandList) {
        const word = bw.toLowerCase();
        let idx = out.toLowerCase().indexOf(word);

        while (idx !== -1) {
            const before = out[idx - 1] || "";
            const after  = out[idx + word.length] || "";

            const isParam =
                (/[\d]/.test(before) && /[\d]/.test(after)) ||
                (/[\d]/.test(before) && /[a-z]/i.test(after)) ||
                (/[a-z]/i.test(before) && /[\d]/.test(after)) ||
                (/[x×]/i.test(before) && /[\d]/.test(after)) ||
                (/[\d]/.test(before) && /[mkc]g?/i.test(after));

            if (!isParam && isBoundaryChar(before) && isBoundaryChar(after)) {
                out = out.slice(0, idx) + out.slice(idx + word.length);
                removed++;
                idx = out.toLowerCase().indexOf(word, idx);
            } else {
                idx = out.toLowerCase().indexOf(word, idx + 1);
            }
        }
    }

    if (removed > 0) {
        if (isTitle) brandDeleteInTitle += removed;
        else         brandDeleteInDesc  += removed;
    }

    return out.trim();
}

// ============================================================
// ★ cleanVioV2（方式 1 改写版 + 参数保护）
// ============================================================
function cleanVioV2(text) {
    if (!text) return { text, count: 0 };

    text = text.replace(/[\u200B-\u200F\u202A-\u202E\u2060\uFEFF]/g, "");
    text = text.replace(/&nbsp;|&#160;/gi, " ");
    text = text.replace(/&amp;|&#38;/gi, "&");
    text = text.replace(/&#(\d+);/g, (_, code) => String.fromCharCode(code));

    const full2half = s => s.replace(/[\uFF01-\uFF5E]/g, ch =>
        String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
    ).replace(/\u3000/g, " ");
    text = full2half(text);

    let out = text;
    let removed = 0;

    for (let vw of vioList) {
        const word = vw.toLowerCase();
        let idx = out.toLowerCase().indexOf(word);

        while (idx !== -1) {
            const before = out[idx - 1] || "";
            const after  = out[idx + word.length] || "";

            const isParam =
                (/[\d]/.test(before) && /[\d]/.test(after)) ||
                (/[\d]/.test(before) && /[a-z]/i.test(after)) ||
                (/[a-z]/i.test(before) && /[\d]/.test(after)) ||
                (/[x×]/i.test(before) && /[\d]/.test(after)) ||
                (/[\d]/.test(before) && /[mkc]g?/i.test(after));

            if (!isParam && isBoundaryChar(before) && isBoundaryChar(after)) {
                out = out.slice(0, idx) + out.slice(idx + word.length);
                removed++;
                idx = out.toLowerCase().indexOf(word, idx);
            } else {
                idx = out.toLowerCase().indexOf(word, idx + 1);
            }
        }
    }

    return { text: out.trim(), count: removed };
}
// ======================
let titleArr      = [];
let brandArr      = [];
let descArrs      = [];
let rowHighlight  = new Array(rowCount + 1).fill(false);

if (titleCol) {
  for (let i = 2; i <= rowCount; i++) {
    titleArr[i] = cleanTitlePrefix(removeEmojiAndFixSpaces(safeString(sheet.Cells(i, titleCol).Value2)));
  }
}

if (brandCol) {
  for (let i = 2; i <= rowCount; i++) {
    brandArr[i] = removeEmojiAndFixSpaces(safeString(sheet.Cells(i, brandCol).Value2));
  }
}

for (let col of descCols) {
  const arr = [];
  for (let i = 2; i <= rowCount; i++) {
    arr[i] = removeEmojiAndFixSpaces(safeString(sheet.Cells(i, col).Value2));
  }
  descArrs.push({ col, arr });
}

updateProgress(40, "数据已批量读取");

for (let i = 2; i <= rowCount; i++) {

    // 产品名称违规词清洗
    let r = cleanVioV2(titleArr[i]);
    if (r.count > 0) {
        titleArr[i] = r.text;
        vioHitInTitle     += r.count;
        violationHitCount += r.count;
        rowHighlight[i]    = true;
    }

    // 品牌列违规词检查
    const brandClean = brandArr[i];
    const hitBrand = checkBrandHasViolation(brandClean);

    if (hitBrand > 0) {
        vioHitInBrand     += hitBrand;
        violationHitCount += hitBrand;
        sheet.Cells(i, brandCol).FormatConditions.Delete();
        sheet.Cells(i, brandCol).Interior.Color = 255;
    }

    // 描述列违规词清洗
    for (let d of descArrs) {
        let r2 = cleanVioV2(d.arr[i]);
        if (r2.count > 0) {
            d.arr[i] = r2.text;
            vioHitInDesc      += r2.count;
            violationHitCount += r2.count;
        }
    }

    // 产品名称品牌词清洗
    titleArr[i] = cleanBrandV2(titleArr[i], true);

    // 描述品牌词清洗
    for (let d of descArrs) {
        d.arr[i] = cleanBrandV2(d.arr[i], false);
    }
}

updateProgress(80, "清洗完成，准备写回");

// 写回产品名称
if (titleCol) {
  for (let i = 2; i <= rowCount; i++) sheet.Cells(i, titleCol).Value2 = titleArr[i];
}

// 写回品牌列
if (brandCol) {
  for (let i = 2; i <= rowCount; i++) sheet.Cells(i, brandCol).Value2 = brandArr[i];
}

// 写回描述列
for (let d of descArrs) {
  for (let i = 2; i <= rowCount; i++) sheet.Cells(i, d.col).Value2 = d.arr[i];
}

// 高亮违规行
for (let i = 2; i <= rowCount; i++) {
  if (rowHighlight[i]) sheet.Rows(i).Interior.Color = 65535;
}

// ============================================================
// ★ 价格计算部分
// ============================================================

function fmt2(v) { return Number(v.toFixed(2)); }

// 插入“价格取最大值”列
sheet.Columns(priceCol + 1).Insert();
const maxPriceCol = priceCol + 1;
sheet.Cells(1, maxPriceCol).Value2 = "价格取最大值";

// 插入“售价（价格×2.2）”
sheet.Columns(maxPriceCol + 1).Insert();
const salePriceCol = maxPriceCol + 1;
sheet.Cells(1, salePriceCol).Value2 = "售价（价格×2.2）";

// 插入“原价（售价×1.68）”
sheet.Columns(salePriceCol + 1).Insert();
const originalPriceCol = salePriceCol + 1;
sheet.Cells(1, originalPriceCol).Value2 = "原价（售价×1.68）";

// 插入“最低价（售价×0.95）”
sheet.Columns(originalPriceCol + 1).Insert();
const lowestPriceCol = originalPriceCol + 1;
sheet.Cells(1, lowestPriceCol).Value2 = "最低价（售价×0.95）";

for (let i = 2; i <= rowCount; i++) {
    const p1 = parseFloat(sheet.Cells(i, priceCol).Value2) || 0;
    const p2 = parseFloat(sheet.Cells(i, buyboxCol).Value2) || 0;

    const maxPrice      = fmt2(Math.max(p1, p2));
    const salePrice     = fmt2(maxPrice * 2.2);
    const originalPrice = fmt2(salePrice * 1.68);
    const lowestPrice   = fmt2(salePrice * 0.95);

    sheet.Cells(i, maxPriceCol).Value2      = maxPrice;
    sheet.Cells(i, salePriceCol).Value2     = salePrice;
    sheet.Cells(i, originalPriceCol).Value2 = originalPrice;
    sheet.Cells(i, lowestPriceCol).Value2   = lowestPrice;

    priceAdjustedCount++;
}

// ============================================================
// ★ 生成产品ID列（ASIN + 自定义前缀3位 + 随机10位）
// ============================================================

let asinCol = null;
for (let j = 1; j <= colCount; j++) {
    const h = (sheet.Cells(1, j).Text || "").trim();
    if (!asinCol && matchHeader(h, ["ASIN","asin","Asin"])) asinCol = j;
}

if (asinCol) {

    sheet.Columns(asinCol + 1).Insert();
    const productIdCol = asinCol + 1;
    sheet.Cells(1, productIdCol).Value2 = "产品ID";

    const CUSTOM_PREFIX = "HXJ";

    function random10() {
        const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        let out = "";
        for (let i = 0; i < 10; i++) {
            out += chars[Math.floor(Math.random() * chars.length)];
        }
        return out;
    }

    for (let i = 2; i <= rowCount; i++) {
        const asin = safeString(sheet.Cells(i, asinCol).Value2).trim();
        const rand = CUSTOM_PREFIX + random10();
        sheet.Cells(i, productIdCol).Value2 = asin + rand;
    }
}
// ======================
// ★ 统计信息写入
// ======================
let statStart = sheet.UsedRange.Rows.Count + 2;

function writeStat(label, value) {
  sheet.Cells(statStart, 1).Value2 = label;
  sheet.Cells(statStart, 2).Value2 = value;
  statStart++;
}

writeStat("品牌词替换（产品名称）", brandDeleteInTitle);
writeStat("品牌词替换（描述）",     brandDeleteInDesc);
writeStat("违规词命中（品牌列）",   vioHitInBrand);
writeStat("违规词命中（描述列）",   vioHitInDesc);
writeStat("违规词命中（产品名称）", vioHitInTitle);
writeStat("违规词命中总数",         violationHitCount);
writeStat("图片重复数",             duplicateImageCount);
writeStat("价格调整行数",           priceAdjustedCount);

Application.StatusBar = "处理完成";
