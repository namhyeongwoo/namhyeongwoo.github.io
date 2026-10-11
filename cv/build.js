// Build files/CV_Hyeongwoo_Nam.docx from cv.json.
//   cd cv && node build.js
// Then make the PDF that the website links to:
//   soffice --headless --convert-to pdf --outdir ../files ../files/CV_Hyeongwoo_Nam.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, AlignmentType,
  TabStopType, BorderStyle,
} = require("docx");

const cv = JSON.parse(fs.readFileSync(path.join(__dirname, "cv.json"), "utf8"));
const SELF = cv.name;
const BLUE = "0A2F6B";
const MUTED = "5D6A80";
const FONT = "Arial";
// A4 with 0.8" margins: 11906 - 2 * 1152 = 9602 DXA of text width.
const TEXT_W = 9602;

const run = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const link = (text, url, o = {}) =>
  new ExternalHyperlink({ link: url, children: [run(text, { color: BLUE, ...o })] });

function heading(text) {
  return new Paragraph({
    spacing: { before: 220, after: 90 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "C9D3E3", space: 2 } },
    children: [run(text.toUpperCase(), { bold: true, size: 20, color: BLUE, characterSpacing: 20 })],
  });
}

// Bold left text with a right-aligned date on the same line.
function headLine(left, date, o = {}) {
  return new Paragraph({
    spacing: { before: o.before ?? 120, after: 0 },
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }],
    children: [run(left, { bold: true, size: 21 }), run("\t" + (date || ""), { size: 20, color: MUTED })],
  });
}

function line(children, o = {}) {
  return new Paragraph({ spacing: { before: o.before ?? 0, after: o.after ?? 0 }, indent: o.indent, children });
}

function authorRuns(authors) {
  const out = [];
  authors.forEach((a, i) => {
    const m = a.match(/^(.*?)([*†]*)$/);
    const name = m[1];
    out.push(run(name, { size: 20, bold: name === SELF }));
    if (m[2]) out.push(run(m[2], { size: 20, superScript: true }));
    if (i < authors.length - 1) out.push(run(", ", { size: 20 }));
  });
  return out;
}

function pubBlock(p, label) {
  const kids = [];
  kids.push(new Paragraph({
    spacing: { before: 120, after: 0 },
    indent: { left: 520, hanging: 520 },
    children: [run(label + "\t", { bold: true, size: 20, color: BLUE }), run(p.title, { bold: true, size: 20 })],
    tabStops: [{ type: TabStopType.LEFT, position: 520 }],
  }));
  kids.push(line(authorRuns(p.authors), { indent: { left: 520 } }));
  const venue = [run(p.venue, { size: 20, italics: true, color: "3C4A60" })];
  kids.push(line(venue, { indent: { left: 520 } }));
  if (p.links && p.links.length) {
    const l = [];
    p.links.forEach((x, i) => {
      l.push(link(x.text, x.url, { size: 19 }));
      if (i < p.links.length - 1) l.push(run("  ·  ", { size: 19, color: MUTED }));
    });
    kids.push(line(l, { indent: { left: 520 } }));
  }
  if (p.note) kids.push(line([run(p.note, { size: 17, color: MUTED })], { indent: { left: 520 } }));
  return kids;
}

const body = [];

// Header
body.push(new Paragraph({ spacing: { after: 40 }, children: [run(cv.name, { bold: true, size: 40, color: BLUE })] }));
body.push(line([run(cv.title, { size: 21, color: "3C4A60" })]));
body.push(line([run(cv.lab, { size: 21, color: "3C4A60" })], { after: 60 }));
cv.contact.forEach((row, r) => {
  const kids = [];
  row.forEach((c, i) => {
    kids.push(c.url ? link(c.text, c.url, { size: 19 }) : run(c.text, { size: 19 }));
    if (i < row.length - 1) kids.push(run("  ·  ", { size: 19, color: MUTED }));
  });
  body.push(line(kids, { after: r === cv.contact.length - 1 ? 40 : 0 }));
});

// Research interests
body.push(heading("Research Interests"));
body.push(line([run(cv.interests, { size: 20 })], { before: 40 }));

// Education
body.push(heading("Education"));
cv.education.forEach((e, i) => {
  body.push(headLine(`${e.school}, ${e.location}`, e.date, { before: i ? 140 : 60 }));
  const sub = [run(e.degree, { size: 20 })];
  if (e.detail) sub.push(run(`  ·  ${e.detail}`, { size: 20, color: MUTED }));
  body.push(line(sub));
});

// Research experience
body.push(heading("Research Experience"));
cv.experience.forEach((x, i) => {
  body.push(headLine(x.role, x.date, { before: i ? 140 : 60 }));
  if (x.project) body.push(line([run(x.project, { size: 20, italics: true })]));
  if (x.org) body.push(line([run(x.org, { size: 19, color: MUTED })]));
});

// Publications
body.push(heading("Publications"));
body.push(line([run("Conference Papers", { bold: true, size: 20, color: MUTED })], { before: 60 }));
cv.publications.conference.forEach((p, i) => body.push(...pubBlock(p, `[C${i + 1}]`)));
body.push(line([run("Journal Articles", { bold: true, size: 20, color: MUTED })], { before: 200 }));
cv.publications.journal.forEach((p, i) => body.push(...pubBlock(p, `[J${i + 1}]`)));

// Skills
body.push(heading("Skills"));
cv.skills.forEach((s, i) => {
  body.push(new Paragraph({
    spacing: { before: i ? 40 : 60, after: 0 },
    tabStops: [{ type: TabStopType.LEFT, position: 2600 }],
    children: [run(s.label, { bold: true, size: 20 }), run("\t" + s.items, { size: 20 })],
  }));
});

// Honors
body.push(heading("Honors & Awards"));
cv.honors.forEach((h, i) => {
  body.push(new Paragraph({
    spacing: { before: i ? 80 : 60, after: 0 },
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }],
    children: [run(h.name, { bold: true, size: 20 }), run(`, ${h.org}`, { size: 20 }), run("\t" + h.date, { size: 20, color: MUTED })],
  }));
});


const doc = new Document({
  creator: cv.name,
  title: `CV — ${cv.name}`,
  styles: { default: { document: { run: { font: FONT, size: 20 }, paragraph: { spacing: { line: 264 } } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 760, bottom: 700, left: 1152, right: 1152 } } },
    children: body,
  }],
});

const out = path.join(__dirname, "..", "files", "CV_Hyeongwoo_Nam.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
