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
    bidirectional: true,
    alignment: map[o.align || "start"],
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
function makeDoc(children, headerKids, footerKids) {
  return new Document({
    creator: "Invoice",
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN + 200, bottom: MARGIN, left: MARGIN, right: MARGIN, header: 0, footer: 0 } } },
      headers: { default: new Header({ children: [new Paragraph({ children: headerKids })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ children: footerKids })] }) },
      children,
    }],
  });
}

// =====================================================================
// 1) ويتكو — مستوحاة من التصميم الثالث (رأس داكن + جدول برتقالي/بني)
// =====================================================================
function wytc() {
  const OR = "E8532B", DK = "2D2D2D", TINT = "FDEDE6", GR = "F3F3F3";
  const kids = [];

  // رأس: [شعار] [بيانات الشركة] [عنوان الفاتورة + بياناتها]
  const metaRows = [["رقم الفاتورة:", ""], ["تاريخ الفاتورة:", ""], ["تاريخ الاستحقاق:", ""]];
  const meta = table([1500, 1700], metaRows.map(([l]) => row([
    cell(p(t(l, { size: 9, color: "F3B49E" })), { w: 1500, m: { top: 30, bottom: 30, left: 40, right: 40 } }),
    cell(p(t("", { size: 9, color: "FFFFFF" })), { w: 1700, m: { top: 30, bottom: 30, left: 40, right: 40 }, borders: { ...noBorders, bottom: line("8A8A8A", 4, BorderStyle.DOTTED) } }),
  ])));
  kids.push(table([2600, 4006, 3600], [row([
    cell([p(img("wytc_logo.png", 150, 59), { align: "center" })], { w: 2600, fill: "FFFFFF", borders: { ...noBorders, top: line(OR, 24), bottom: line(OR, 24) } }),
    cell([
      p(t("ويتكو للاستثمار الصناعي", { bold: true, size: 14, color: DK }), { align: "center" }),
      p(t("وتصنيع العبوات البلاستيكية", { bold: true, size: 12, color: OR }), { align: "center", after: 80 }),
      p(t("بلوك 3 - القطعة 18، شرق المنطقة الصناعية الثانية", { size: 8.5, color: "555555" }), { align: "center" }),
      p(t("برج العرب، الإسكندرية", { size: 8.5, color: "555555" }), { align: "center", after: 40 }),
      p(t("هاتف: ................   البريد: ................", { size: 8.5, color: "555555" }), { align: "center" }),
    ], { w: 4006, fill: TINT, borders: { ...noBorders, top: line(OR, 24), bottom: line(OR, 24) } }),
    cell([
      p(t("فاتورة مبيعات", { bold: true, size: 22, color: OR }), { align: "center", after: 0 }),
      p(t("SALES INVOICE", { size: 8, color: "BBBBBB", spacing: 60, ltr: true }), { align: "center", after: 100 }),
      meta,
    ], { w: 3600, fill: DK, m: { top: 140, bottom: 140, left: 160, right: 160 } }),
  ], 2100)]));
  kids.push(spacer(160));

  // بيانات العميل
  const lw = 1400, vw = 3600, gap = 206;
  kids.push(table([lw, vw, gap, lw, vw], [
    row([cell(p(t("فاتورة إلى", { bold: true, size: 11, color: "FFFFFF" })), { w: lw + vw, span: 2, fill: OR }), cell(p(t("")), { w: gap }), cell(p(t("بيانات إضافية", { bold: true, size: 11, color: "FFFFFF" })), { w: lw + vw, span: 2, fill: DK })]),
    row([...fieldRow("اسم العميل:", lw, vw), cell(p(t("")), { w: gap }), ...fieldRow("الرقم الضريبي:", lw, vw)]),
    row([...fieldRow("العنوان:", lw, vw), cell(p(t("")), { w: gap }), ...fieldRow("طريقة الدفع:", lw, vw)]),
    row([...fieldRow("الجوال:", lw, vw), cell(p(t("")), { w: gap }), ...fieldRow("المندوب:", lw, vw)]),
  ]));
  kids.push(spacer(180));

  // جدول الأصناف
  const W = [700, 900, 2400, 2706, 1700, 1800];
  const hd = ["بند", "الكمية", "الصنف", "الوصف / النوع", "سعر الوحدة", "المبلغ"];
  const hb = allBorders(line("FFFFFF", 6));
  const rows = [row(hd.map((h, i) => cell(p(t(h, { bold: true, color: "FFFFFF", size: 10.5 }), { align: "center" }), { w: W[i], fill: i % 2 ? "B5401F" : OR, borders: hb })), 460)];
  for (let i = 1; i <= 15; i++) {
    const fill = i % 2 ? "FFFFFF" : "FFF4EF";
    const b = { ...allBorders(line("E4C9BE", 4)) };
    rows.push(row(W.map((w, c) => cell(p(t(c === 0 ? i : "", { bold: c === 0, color: c === 0 ? OR : "222222" }), { align: "center" }), { w, fill: c === 2 ? (i % 2 ? "FDF8F5" : "FCEDE6") : fill, borders: b })), 400));
  }
  kids.push(table(W, rows));
  kids.push(spacer(200));

  // الإجماليات + ملاحظات
  const tw = [2300, 1400];
  const tot = (l, fillL, fillV, colL, colV, bold) => row([
    cell(p(t(l, { bold: true, color: colL })), { w: tw[0], fill: fillL, borders: { ...noBorders, bottom: line("FFFFFF", 6) } }),
    cell(p(t("ج.م", { size: 9, color: colV, bold })), { w: tw[1], fill: fillV, borders: { ...noBorders, bottom: line("FFFFFF", 6) } }),
  ], 420);
  const totals = table(tw, [
    tot("الإجمالي:", "D9D9D9", GR, DK, "777777"),
    tot("قيمة الخصم:", "D9D9D9", GR, DK, "777777"),
    tot("ضريبة القيمة المضافة 14%:", "D9D9D9", GR, DK, "777777"),
    tot("المجموع المستحق:", "6B2E17", DK, "FFFFFF", "F3B49E", true),
  ]);
  const notes = [
    p(t("الشروط والملاحظات", { bold: true, size: 11, color: OR }), { after: 60, border: { bottom: line(OR, 8) } }),
    p(t("• المبلغ كتابةً: ..............................................................", { size: 9.5 }), { before: 80, after: 60 }),
    p(t("• البضاعة المباعة لا ترد ولا تستبدل إلا وفقاً للسياسة المعتمدة.", { size: 9.5, color: "555555" }), { after: 40 }),
    p(t("• يرجى ذكر رقم الفاتورة عند السداد.", { size: 9.5, color: "555555" }), { after: 180 }),
    table([3100, 3100], [row([
      cell([p(t("توقيع المستلم", { bold: true }), { align: "center", after: 360 }), p(t("....................................", { color: "999999" }), { align: "center" })], { w: 3100 }),
      cell([p(t("الختم والتوقيع", { bold: true }), { align: "center", after: 360 }), p(t("....................................", { color: "999999" }), { align: "center" })], { w: 3100 }),
    ])]),
  ];
  kids.push(table([6306, 200, 3700], [row([
    cell(notes, { w: 6306, v: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
    cell(p(t("")), { w: 200 }),
    cell(totals, { w: 3700, v: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
  ])]));
  kids.push(spacer(240));
  kids.push(table([CW], [row([cell(p(t("شكراً لتعاملكم معنا!", { bold: true, size: 12, color: "FFFFFF" }), { align: "center" }), { w: CW, fill: "6B2E17", borders: { ...noBorders, top: line(OR, 18) } })], 480)]));

  return save(makeDoc(kids, [floatImg("wytc_top.png", 794, 14, 0, 0)], [floatImg("wytc_bottom.png", 794, 14, 0, 1123 - 14)]), "فاتورة_ويتكو_WYTC.docx");
}

// =====================================================================
// 2) الجوهرة — مستوحاة من التصميم الثاني (أشرطة مائلة + أعمدة داكنة)
// =====================================================================
function gawhara() {
  const BL = "2F6DB5", DK = "1F2A3A", LG = "F3F5F8";
  const kids = [];
  // رأس رمادي فاتح: العنوان يمين والشعار يسار
  kids.push(table([5103, 5103], [row([
    cell([
      p(t("فاتورة مبيعات", { bold: true, size: 26, color: DK }), { after: 40 }),
      p(t("SALES INVOICE", { size: 9, color: BL, spacing: 80, ltr: true }), { after: 120 }),
      table([1500, 2600], [
        row(fieldRow("رقم الفاتورة:", 1500, 2600, { size: 9.5 })),
        row(fieldRow("التاريخ:", 1500, 2600, { size: 9.5 })),
      ]),
    ], { w: 5103, fill: LG, m: { top: 260, bottom: 260, left: 200, right: 300 } }),
    cell([p(img("gawhara_logo.png", 300, 96), { align: "center" })], { w: 5103, fill: LG, m: { top: 200, bottom: 200, left: 200, right: 100 } }),
  ])]));
  kids.push(spacer(260));

  // بيانات العميل (عمودان)
  const lw = 1550, vw = 3350, gap = 406;
  const r2 = (a, b) => row([...fieldRow(a, lw, vw), cell(p(t("")), { w: gap }), ...fieldRow(b, lw, vw)]);
  kids.push(table([lw, vw, gap, lw, vw], [
    r2("اسم العميل:", "العنوان:"),
    r2("اسم البنك:", "رقم الهاتف الأول:"),
    r2("رقم الحساب:", "رقم الهاتف الثاني:"),
  ]));
  kids.push(spacer(260));

  // جدول الأصناف: رؤوس داكنة منفصلة بفواصل بيضاء + عمود الرقم داكن
  const W = [800, 4206, 1700, 1500, 2000];
  const hd = ["الرقم", "اسم المنتج", "السعر", "الكمية", "الإجمالي"];
  const sep = allBorders(line("FFFFFF", 18));
  const rows = [row(hd.map((h, i) => cell(p(t(h, { bold: true, color: "FFFFFF", size: 10.5 }), { align: "center" }), { w: W[i], fill: i === 0 ? BL : DK, borders: sep })), 480)];
  for (let i = 1; i <= 13; i++) {
    const fill = i % 2 ? "F4F4F4" : "FFFFFF";
    rows.push(row(W.map((w, c) => cell(p(t(c === 0 ? i : "", { color: c === 0 ? "FFFFFF" : "222222", bold: c === 0 }), { align: "center" }), {
      w, fill: c === 0 ? DK : fill,
      borders: c === 0 ? { ...noBorders, left: line("FFFFFF", 18), right: line("FFFFFF", 18) } : { ...noBorders, bottom: line("E3E3E3", 4) },
    })), 400));
  }
  kids.push(table(W, rows));
  kids.push(spacer(220));

  // سفلي: سياسة + توقيع (يمين) / إجماليات (يسار)
  const tw = [1900, 2100];
  const tr = (l, fill, col, v, big) => row([
    cell(p(t(l, { bold: true, color: col, size: big ? 11 : 10 })), { w: tw[0], fill }),
    cell(p(t(v, { color: col, size: 9, bold: big }), { align: "end" }), { w: tw[1], fill }),
  ], 420);
  const totals = table(tw, [
    tr("الإجمالي", "FFFFFF", DK, "ج.م"),
    tr("الخصم", "FFFFFF", DK, "ج.م"),
    tr("الضريبة 14%", "FFFFFF", DK, "ج.م"),
    tr("الإجمالي النهائي", BL, "FFFFFF", "ج.م", true),
  ], { borders: { ...noBorders, insideHorizontal: line("E3E3E3", 4), insideVertical: NONE } });
  const left = [
    totals,
    p(t("المبلغ المستحق  Amount Due", { size: 10, color: "666666" }), { before: 160 }),
    p(t("............................ ج.م", { bold: true, size: 16, color: DK }), { before: 60 }),
  ];
  const right = [
    p(t("الاستبدال والاسترجاع:", { bold: true, size: 11, color: DK }), { after: 60 }),
    p(t("الاستبدال والاسترجاع خلال فترة 14 يوماً من تاريخ استلام السلعة، بشرط سلامة العبوات وعدم استخدامها.", { size: 9.5, color: "555555" }), { after: 140, line: 300 }),
    p(t("ملاحظات:", { bold: true, size: 11, color: DK }), { after: 60 }),
    writeLines(2, 5306, "BDBDBD"),
    p(t(""), { after: 200 }),
    p(t("توقيع البائع", { bold: true, size: 12, color: DK }), { align: "center", after: 360 }),
    p(t("______________________", { color: "888888" }), { align: "center" }),
  ];
  kids.push(table([5306, 500, 4400], [row([
    cell(right, { w: 5306, v: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
    cell(p(t("")), { w: 500 }),
    cell(left, { w: 4400, v: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
  ])]));
  kids.push(spacer(200));
  kids.push(p(t("الجوهرة لصناعة البلاستيك ومواد التعبئة والتغليف — تأسست سنة 1990", { size: 9, color: BL, bold: true }), { align: "center" }));
  kids.push(p(t("العنوان: ........................   هاتف: ........................", { size: 8.5, color: "777777" }), { align: "center" }));

  return save(makeDoc(kids, [floatImg("gawhara_top.png", 794, 40, 0, 0)], [floatImg("gawhara_bottom.png", 794, 40, 0, 1123 - 40)]), "فاتورة_الجوهرة.docx");
}

// =====================================================================
// 3) الأصيل — مستوحاة من التصميم الأول (مثلثات ملونة + جدول نظيف)
// =====================================================================
function aseel() {
  const BL = "2E3A9E", RD = "E53935", GN = "2E9E4F", HD = "E3E6F3";
  const kids = [];
  // رأس: العنوان + الحقول يمين / الشعار يسار
  const lw = 1500, vw = 3800;
  kids.push(table([5400, 4806], [row([
    cell([
      p(t("فاتورة مبيعات", { bold: true, size: 26, color: BL }), { align: "center" }),
      p(t("I N V O I C E", { bold: true, size: 11, color: "333333", spacing: 120, ltr: true }), { align: "center", after: 220 }),
      table([lw, vw], [
        row(fieldRow("اسم العميل :", lw, vw)),
        row(fieldRow("التاريخ :", lw, vw)),
        row(fieldRow("رقم الفاتورة :", lw, vw)),
      ]),
    ], { w: 5400, v: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
    cell([p(img("aseel_logo.png", 104, 185), { align: "center" }), p(img("aseel_motif.png", 96, 24), { align: "center", before: 60 })], { w: 4806 }),
  ])]));
  kids.push(spacer(220));

  // جدول الأصناف
  const W = [900, 1100, 1700, 4006, 2500];
  const hd = ["رقم", "الكمية", "السعر", "وصف المنتج", "الإجمالي"];
  const b = allBorders(line("6B6B6B", 6));
  const rows = [row(hd.map((h, i) => cell(p(t(h, { bold: true, size: 10.5, color: BL }), { align: "center" }), { w: W[i], fill: HD, borders: b })), 460)];
  for (let i = 1; i <= 14; i++) {
    rows.push(row(W.map((w, c) => cell(p(t(c === 0 ? i : "", { bold: c === 0 }), { align: "center" }), { w, borders: b, fill: c === 0 ? "F7F8FC" : undefined })), 400));
  }
  kids.push(table(W, rows));

  // صفوف الإجمالي
  const sw = [6206, 4000];
  const sum = (l, fill, col, h) => row([
    cell(p(t(l, { bold: true, size: 10.5, color: col })), { w: sw[0], fill, borders: allBorders(line("6B6B6B", 6)) }),
    cell(p(t("ج.م", { size: 9, color: col === "FFFFFF" ? "FFFFFF" : "777777", bold: true }), { align: "end" }), { w: sw[1], fill: fill === BL ? BL : "FFFFFF", borders: allBorders(line("6B6B6B", 6)) }),
  ], h);
  kids.push(table(sw, [
    sum("الإجمالي قبل الضريبة :", "F0F0F0", "222222", 400),
    sum("الخصم :", "F0F0F0", "222222", 400),
    sum("ضريبة القيمة المضافة (14%) :", "F0F0F0", "222222", 400),
    sum("الإجمالي المطلوب سداده :", BL, "FFFFFF", 460),
  ]));
  kids.push(spacer(220));

  // ملاحظات + تواصل
  const notes = [
    p(t("ملاحظات :", { bold: true, size: 12, color: BL }), { after: 80 }),
    writeLines(4, 5600, "777777"),
  ];
  const contact = [
    p([t("هاتف: ", { bold: true, color: RD }), t("..............................")], { after: 80 }),
    p([t("البريد: ", { bold: true, color: GN }), t("..............................")], { after: 80 }),
    p([t("العنوان: ", { bold: true, color: BL }), t("..............................")], { after: 160 }),
    p(t("الختم والتوقيع", { bold: true }), { align: "center", after: 320 }),
    p(t("..............................", { color: "999999" }), { align: "center" }),
  ];
  kids.push(table([5600, 406, 4200], [row([
    cell(notes, { w: 5600, v: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
    cell(p(t("")), { w: 406 }),
    cell(contact, { w: 4200, v: VerticalAlign.TOP, fill: "F7F8FC", m: { top: 140, bottom: 140, left: 180, right: 180 }, borders: { ...noBorders, right: line(GN, 24) } }),
  ])]));
  kids.push(spacer(160));
  kids.push(p(t("الأصيل للبلاستيك ومواد التعبئة والتغليف", { bold: true, size: 10, color: BL }), { align: "center" }));

  return save(makeDoc(kids,
    [floatImg("aseel_tr.png", 230, 120, 794 - 230, 0)],
    [floatImg("aseel_bl.png", 230, 120, 0, 1123 - 120)]), "فاتورة_الأصيل.docx");
}

(async () => { await wytc(); await gawhara(); await aseel(); })();
