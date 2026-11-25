// يبني 3 فواتير Word (واحدة لكل شركة) — تشغيل: node build_invoices.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Header, Footer,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TableLayoutType, HeightRule,
} = require("docx");

const A = (f) => fs.readFileSync(path.join(__dirname, "assets", f));
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 800;
const CW = PAGE_W - 2 * MARGIN; // عرض المحتوى
const FONT = "Arial";
const EMU = 9525; // px -> EMU

// ---------- أدوات مساعدة ----------
function t(text, o = {}) {
  const size = (o.size || 10) * 2;
  return new TextRun({
    text: String(text), rightToLeft: o.ltr ? false : true,
    font: { ascii: FONT, hAnsi: FONT, cs: FONT, eastAsia: FONT },
    size, sizeComplexScript: size,
    bold: !!o.bold, boldComplexScript: !!o.bold,
    color: o.color || "222222", characterSpacing: o.spacing,
  });
}
// align: "start" (يمين) | "end" (يسار) | "center"
function p(runs, o = {}) {
  const map = { center: AlignmentType.CENTER, end: AlignmentType.RIGHT, start: undefined };
  return new Paragraph({
    bidirectional: !o.ltrPara,
    alignment: o.ltrPara ? ({ start: AlignmentType.LEFT, end: AlignmentType.LEFT, center: AlignmentType.CENTER })[o.align || "start"] : map[o.align || "start"],
    spacing: { before: o.before || 0, after: o.after || 0, line: o.line },
    border: o.border,
    children: Array.isArray(runs) ? runs : [runs],
  });
}
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const line = (color, size = 4, style = BorderStyle.SINGLE) => ({ style, size, color });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });

function cell(children, o = {}) {
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    columnSpan: o.span,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    borders: o.borders || noBorders,
    verticalAlign: o.v || VerticalAlign.CENTER,
    margins: o.m || { top: 60, bottom: 60, left: 100, right: 100 },
    children: Array.isArray(children) ? children : [children],
  });
}
function table(widths, rows, o = {}) {
  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    borders: o.borders || { ...noBorders, insideHorizontal: NONE, insideVertical: NONE },
    rows,
  });
}
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const spacer = (after = 120) => p(t("", { size: 6 }), { after });
// أسطر كتابة (جدول بخط سفلي لكل سطر)
const writeLines = (n, w, color = "9E9E9E") => table([w], Array.from({ length: n }, () => row([cell(p(t("")), { w, borders: { ...noBorders, bottom: line(color, 4) } })], 380)));

function img(file, w, h, type = "png") {
  return new ImageRun({ type, data: A(file), transformation: { width: w, height: h } });
}
function floatImg(file, w, h, x, y) {
  return new ImageRun({
    type: "png", data: A(file), transformation: { width: w, height: h },
    floating: {
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: Math.round(x * EMU) },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: Math.round(y * EMU) },
      behindDocument: true, allowOverlap: true,
    },
  });
}
// حقل: عنوان + خط منقّط للتعبئة
function fieldRow(label, wl, wv, o = {}) {
  return [
    cell(p(t(label, { bold: true, size: o.size || 10, color: o.lc })), { w: wl, m: { top: 50, bottom: 50, left: 60, right: 60 } }),
    cell(p(t(o.value || "", { size: o.size || 10 })), {
      w: wv, m: { top: 50, bottom: 50, left: 60, right: 60 },
      borders: { ...noBorders, bottom: line(o.lineColor || "9E9E9E", 4, BorderStyle.DOTTED) },
    }),
  ];
}
function save(doc, name) {
  return Packer.toBuffer(doc).then((b) => { fs.writeFileSync(path.join(__dirname, name), b); console.log("wrote", name); });
}
function makeDoc(children, headerKids, footerKids, o = {}) {
  const page = { size: { width: PAGE_W, height: PAGE_H }, margin: { top: o.top || MARGIN, bottom: o.bottom || MARGIN, left: o.side || MARGIN, right: o.side || MARGIN, header: 0, footer: 0 } };
  if (o.pageBorder) page.borders = { pageBorderTop: o.pageBorder, pageBorderBottom: o.pageBorder, pageBorderLeft: o.pageBorder, pageBorderRight: o.pageBorder, pageBorders: { offsetFrom: "page" } };
  return new Document({
    creator: "Invoice",
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    sections: [{
      properties: { page },
      headers: { default: new Header({ children: [new Paragraph({ children: headerKids })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ children: footerKids })] }) },
      children,
    }],
  });
}
const pad0 = { top: 0, bottom: 0, left: 0, right: 0 };
const blank = (w) => cell(p(t("")), { w });

// =====================================================================
// التصميم الأول (مثلثات ملونة) — الرواد للبلاستيك
// =====================================================================
function design1() {
  const HD = "D7D7D7", BD = line("4A4A4A", 6), kids = [];
  const W = [1030, 1065, 2235, 4021, 1855]; // رقم، الكمية، السعر، وصف المنتج، الإجمالي
  // العنوان + الزخرفة
  kids.push(table([5600, 4606], [row([
    cell([
      p(t("فاتورة مبيعات", { bold: true, size: 26, color: "111111" }), { align: "center" }),
      p(t("I N V O I C E", { bold: true, size: 12, color: "111111", spacing: 160, ltr: true }), { align: "center" }),
    ], { w: 5600, m: pad0 }),
    cell(p(img("d1_motif.png", 120, 31), { align: "center", before: 260 }), { w: 4606, m: pad0 }),
  ])]));
  kids.push(p(t("الرواد للبلاستيك ومواد التعبئة والتغليف", { bold: true, size: 10, color: "555555" }), { align: "center", before: 80, after: 160 }));
  // الحقول
  const lw = 1500, vw = 4800;
  kids.push(table([lw, vw, CW - lw - vw], ["اسم العميل :", "التـاريـخ :", "رقم الفاتـورة :"].map((l) => row([...fieldRow(l, lw, vw, { size: 11 }), blank(CW - lw - vw)], 420))));
  kids.push(spacer(260));
  // رأس الجدول منفصل
  const hd = ["رقم", "الكمية", "السعـــر", "وصـف المنتـج", "الإجمالـي"];
  kids.push(table(W, [row(hd.map((h, i) => cell(p(t(h, { bold: true, size: 11 }), { align: "center" }), { w: W[i], fill: HD, borders: allBorders(BD) })), 520)]));
  kids.push(spacer(200));
  const rows = [];
  for (let i = 1; i <= 10; i++) rows.push(row(W.map((w, c) => cell(p(t(c === 0 ? i : "", { bold: true, size: 11 }), { align: "center" }), { w, borders: allBorders(BD) })), 580));
  kids.push(table(W, rows));
  // الإجمالي المطلوب سداده (أسفل يسار)
  const tw = W[3] + W[4] - 400;
  kids.push(table([CW - tw, tw], [row([blank(CW - tw), cell(p(t("الإجمالي المطلوب سداده :", { bold: true, size: 11 })), { w: tw, fill: HD, borders: allBorders(BD) })], 480)]));
  kids.push(spacer(200));
  // ملاحظات + أيقونات تواصل
  const nw = 5300;
  const lines = table([nw], [0, 1, 2].map(() => row([cell(p(t("")), { w: nw, borders: { ...noBorders, bottom: line("333333", 6) } })], 380)));
  const icon = (f, label) => row([cell(p(img(f, 22, 22)), { w: 500, m: { top: 30, bottom: 30, left: 0, right: 0 } }), cell(p(t(label, { size: 10, color: "333333" })), { w: 4800 })]);
  kids.push(table([nw, CW - nw], [row([
    cell([p(t("ملاحظات :", { bold: true, size: 13 }), { after: 40 }), lines, spacer(200),
      table([500, 4800], [icon("d1_phone.png", "......................................"), icon("d1_mail.png", "......................................"), icon("d1_pin.png", "......................................")]),
    ], { w: nw, v: VerticalAlign.TOP, m: pad0 }),
    blank(CW - nw),
  ])]));
  return save(makeDoc(kids,
    [floatImg("d1_tr.png", 300, 80, 794 - 300, 0)],
    [floatImg("d1_bl.png", 260, 135, 0, 1123 - 135)],
    { top: 1400, bottom: 600, side: 1050 }), "فاتورة_الرواد.docx");
}

// =====================================================================
// التصميم الثاني (أسود/أصفر) — اللؤلؤة لصناعة البلاستيك
// =====================================================================
function design2() {
  const BK = "0A0A0A", YL = "FFE016", GR = "F4F4F4", kids = [];
  // داخل المنطقة الرمادية: العنوان يمين
  kids.push(table([5200, CW - 5200], [row([
    cell([
      p(t("فاتورة مبيعات", { bold: true, size: 28, color: "333333" }), { after: 60 }),
      p(t("اللؤلؤة لصناعة البلاستيك ومواد التعبئة والتغليف", { bold: true, size: 10.5, color: "555555" }), { after: 40 }),
      p([t("رقم الفاتورة: ", { size: 9.5, color: "555555" }), t("...............   ", { size: 9.5, color: "999999" }), t("التاريخ: ", { size: 9.5, color: "555555" }), t("...............", { size: 9.5, color: "999999" })]),
    ], { w: 5200, m: pad0 }),
    blank(CW - 5200),
  ])]));
  kids.push(spacer(1250));
  // الحقول (عمودان كما في الصورة)
  const lw = 1900, vw = 3000, gap = CW - 2 * (lw + vw);
  const r2 = (a, b, h) => row([...fieldRow(a, lw, vw, { size: 10.5 }), blank(gap), ...fieldRow(b, lw, vw, { size: 10.5 })], h);
  kids.push(table([lw, vw, gap, lw, vw], [r2("اسم العميل :", "العنوان :", 520), r2("اسم البنك :", "رقم الهاتف الاول :"), r2("رقم الحساب :", "رقم الهاتف الثاني :")]));
  kids.push(spacer(380));
  // الجدول: خانات رأس سوداء منفصلة + عمود الرقم أسود
  const W = [1000, 160, 4386, 160, 1500, 1500, 1500]; // الرقم، فاصل، اسم المنتج، فاصل، السعر، الكمية، اجمالي
  const sep = { ...noBorders, left: line("FFFFFF", 12), right: line("FFFFFF", 12) };
  const hd = ["الرقم", null, "أسم المنتج", null, "السعر", "الكمية", "اجمالي"];
  const rows = [row(hd.map((h, i) => h === null ? blank(W[i]) : cell(p(t(h, { bold: true, color: "FFFFFF", size: 10.5 }), { align: "center" }), { w: W[i], fill: BK, borders: sep })), 470)];
  rows.push(row(W.map((w) => cell(p(t("", { size: 4 })), { w, m: pad0 })), 120));
  for (let i = 1; i <= 8; i++) {
    const f = i % 2 ? GR : "FFFFFF";
    rows.push(row(W.map((w, c) => {
      if (c === 0) return cell(p(t(i, { color: "FFFFFF", size: 11 }), { align: "center" }), { w, fill: BK });
      if (c === 1) return blank(w);
      return cell(p(t("", { size: 10.5 }), { align: c === 2 ? "start" : "center" }), { w, fill: c === 3 ? f : f });
    }), 470));
  }
  kids.push(table(W, rows));
  kids.push(spacer(360));
  // الإجماليات يسار + سياسة الاسترجاع يمين
  const tw = [1900, 2200];
  const tot = (l, fill, h) => row([cell(p(t(l, { bold: true, size: 11 })), { w: tw[0], fill }), cell(p(t("", { size: 11 }), { align: "end" }), { w: tw[1], fill })], h || 440);
  const totals = table(tw, [tot("الاجمالي"), tot("الضريبة"), row([blank(tw[0]), blank(tw[1])], 120), tot("الاجمالي", YL, 480)]);
  const rw = CW - 4100 - 500;
  kids.push(table([rw, 500, 4100], [row([
    cell([p(t("الاستبدال و الاسترجاع :", { bold: true, size: 11 }), { before: 300, after: 120 }),
      p(t("الإستبدال والإسترجاع خلال فترة الـ 14 يوم من تاريخ أستلام السلعة .", { size: 10.5, color: "333333" }), { line: 320 })],
      { w: rw, v: VerticalAlign.TOP, m: pad0 }),
    blank(500),
    cell(totals, { w: 4100, v: VerticalAlign.TOP, m: pad0 }),
  ])]));
  kids.push(spacer(200));
  kids.push(table([rw, 500, 4100], [row([
    cell([p(t("توقيع البائع", { bold: true, size: 14 }), { align: "center", before: 300, after: 420 }), p(t("____________________"), { align: "center" })], { w: rw, v: VerticalAlign.TOP, m: pad0 }),
    blank(500),
    cell([p(t("Amount Due :", { size: 14, color: "222222", ltr: true }), { ltrPara: true }), p(t("..................", { bold: true, size: 26, ltr: true }), { ltrPara: true })], { w: 4100, v: VerticalAlign.TOP, m: pad0 }),
  ])]));
  return save(makeDoc(kids,
    [floatImg("d2_top.png", 794, 258, 0, 0)],
    [floatImg("d2_bottom.png", 794, 48, 0, 1123 - 48)],
    { top: 1650, side: 1100 }), "فاتورة_اللؤلؤة.docx");
}

// =====================================================================
// التصميم الثالث (جدول بني على طريقة Excel) — النخبة للصناعات البلاستيكية
// =====================================================================
function design3() {
  const BK = "000000", OR = "E0963F", TAN = "C9B48A", BR = "B05A0A", DBR = "68332D", LB = "DCF0FB", GRY = "A8A9AD", LG = "F3F3F3";
  const kids = [];
  const k = (l, w, col = OR) => cell(p(t(l, { size: 9, color: col, bold: true })), { w, m: { top: 30, bottom: 30, left: 40, right: 40 } });
  const meta = table([1400, 1700], ["رقم الفاتورة:", "تاريخ الفاتورة:", "تاريخ الاستحقاق:"].map((l) => row([k(l, 1400), cell(p(t("..................", { size: 9, color: OR, ltr: true }), { align: "end" }), { w: 1700, m: { top: 30, bottom: 30, left: 40, right: 40 } })])));
  // الرأس: يمين أسود (اسم مختصر) | وسط ذهبي (بيانات الشركة) | يسار أسود (العنوان + بيانات الفاتورة)
  const H = [1700, 4106, 4300];
  const wHead = H.reduce((a, b) => a + b);
  kids.push(table(H, [row([
    cell([p(t("النخبة", { bold: true, size: 16, color: TAN }), { align: "center" })], { w: H[0], fill: BK }),
    cell([
      p(t("................@gmail.com", { bold: true, size: 10, color: "111111", ltr: true }), { align: "center", ltrPara: true }),
      p(t("هاتف: .....................", { bold: true, size: 10, color: "111111" }), { align: "center" }),
      p(t("العنوان: ...........................", { bold: true, size: 9.5, color: "111111" }), { align: "center", after: 60 }),
      p(t("النخبة للصناعات البلاستيكية", { bold: true, size: 12, color: "111111" }), { align: "center" }),
      p(t("والعبوات ومواد التغليف", { bold: true, size: 10.5, color: "111111" }), { align: "center" }),
    ], { w: H[1], fill: TAN, m: { top: 100, bottom: 100, left: 100, right: 100 } }),
    cell([p(t("فاتورة مبيعات", { bold: true, size: 22, color: OR }), { align: "center", after: 120 }), meta], { w: H[2], fill: BK, m: { top: 120, bottom: 120, left: 120, right: 120 } }),
  ], 2300)]));
  // فاتورة إلى (صندوق رمادي)
  const cw = [1400, 4106, 1800, 2800];
  const ci = (l, v, bold) => cell(p(t(l, { bold: !!bold, size: 11 })), { w: v, fill: LG });
  kids.push(table(cw, [
    row([ci("فاتورة إلى:", cw[0], true), cell(p(t("....................................", { color: "999999" })), { w: cw[1], fill: LG }), cell(p(t("")), { w: cw[2], fill: LG }), cell(p(t("الموقع", { bold: true }), { align: "center" }), { w: cw[3], fill: LG })], 520),
    row([ci("العنوان", cw[0], true), cell(p(t("....................................", { color: "999999" })), { w: cw[1], fill: LG }), cell(p(t("")), { w: cw[2], fill: LG }), cell(p(t("....................", { color: "999999" }), { align: "center" }), { w: cw[3], fill: LG })], 460),
    row([ci("جوال", cw[0], true), cell(p(t("....................................", { color: "999999" })), { w: cw[1], fill: LG }), cell(p(t("")), { w: cw[2], fill: LG }), cell(p(t("")), { w: cw[3], fill: LG })], 460),
  ], { borders: { ...allBorders(line("555555", 6)), insideHorizontal: NONE, insideVertical: NONE } }));
  // جدول الأصناف
  const W = [600, 1000, 2500, 2806, 1600, 1600]; // بند، الكمية، الصنف، النوع، سعر الوحدة، المبلغ
  const b = allBorders(line("555555", 6));
  const hd = ["بند", "الكمية", "الصنف", "النوع", "سعر الوحدة", "المبلغ"];
  const rows = [row(hd.map((h, i) => cell(p(t(h, { bold: true, color: "FFFFFF", size: 11 }), { align: "center" }), { w: W[i], fill: BR, borders: b })), 460)];
  for (let i = 1; i <= 13; i++) {
    const f = i % 2 ? "FFFFFF" : LG;
    rows.push(row(W.map((w, c) => cell(p(t(c === 0 ? i : "", { size: 11, bold: c === 0 }), { align: "center" }), { w, borders: b, fill: c === 2 && i % 2 === 0 ? LB : f })), 600));
  }
  kids.push(table(W, rows));
  // الإجماليات
  const sw = [W[0] + W[1] + W[2] + W[3], W[4], W[5]];
  const sr = (l, fill, col, h) => row([blank(sw[0]), cell(p(t(l, { bold: true, size: 11, color: col })), { w: sw[1], fill }), cell(p(t("0.00 ج.م", { bold: true, size: 10.5, color: col }), { align: "center" }), { w: sw[2], fill })], h);
  kids.push(table(sw, [
    row([blank(sw[0]), cell(p(t("الإجمالي :", { bold: true, size: 11 })), { w: sw[1], fill: GRY }), cell(p(t("0.00 ج.م", { bold: true, size: 10.5 }), { align: "center" }), { w: sw[2], fill: GRY })], 460),
    row([blank(sw[0]), cell(p(t("قيمة الخصم", { bold: true, size: 11 })), { w: sw[1], fill: GRY }), cell(p(t("0.00 ج.م", { bold: true, size: 10.5 }), { align: "center" }), { w: sw[2], fill: GRY })], 460),
    row([blank(sw[0]), cell(p(t("")), { w: sw[1], fill: GRY }), cell(p(t("")), { w: sw[2], fill: GRY })], 300),
    sr("المجموع", DBR, "FFFFFF", 460),
  ]));
  kids.push(table([wHead], [row([cell(p(t("شكراً لتعاملك معنا!", { bold: true, size: 13, color: "FFFFFF" }), { align: "center" }), { w: wHead, fill: DBR })], 480)]));
  return save(makeDoc(kids, [], [], { top: 800, bottom: 600, side: 900, pageBorder: { style: BorderStyle.SINGLE, size: 12, color: "000000", space: 24 } }), "فاتورة_النخبة.docx");
}

(async () => { await design1(); await design2(); await design3(); })();
