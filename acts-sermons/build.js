// 사도행전 주제별 연속설교 — JSON → docx 변환기
// 사용법: NODE_PATH=$(npm root -g) node build.js [json파일...]  (인자 없으면 json/ 전체)
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } = require('docx');
const fs = require('fs');
const path = require('path');

const F = "맑은 고딕";
const p = (text) => new Paragraph({ children: [new TextRun({ text, font: F, size: 24 })], spacing: { after: 160, line: 360 } });
const heading1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font: F, size: 32, bold: true, color: "1F3864" })],
  spacing: { before: 400, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 4 } }
});
const heading2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font: F, size: 26, bold: true, color: "2E4057" })],
  spacing: { before: 300, after: 160 }
});
const subLabel = (text) => new Paragraph({ children: [new TextRun({ text, font: F, size: 24, bold: true, color: "C00000" })], spacing: { before: 240, after: 100 } });
const quoteBlock = (text) => new Paragraph({
  children: [new TextRun({ text, font: F, size: 24, italics: true, color: "404040" })],
  spacing: { after: 160, line: 360 },
  indent: { left: 600, right: 600 },
  border: { left: { style: BorderStyle.SINGLE, size: 16, color: "C00000", space: 8 } }
});
const connector = (text) => new Paragraph({
  children: [new TextRun({ text: "→ " + text, font: F, size: 24, bold: true, color: "1F6A8C" })],
  spacing: { before: 200, after: 200 }, alignment: AlignmentType.RIGHT
});
const empty = () => new Paragraph({ children: [new TextRun("")], spacing: { after: 80 } });
const paras = (arr) => (arr || []).map(p);

function fmtDate(d) {
  const [y, m, day] = d.split('-').map(Number);
  return `${y}년 ${m}월 ${day}일 주일 오전 예배`;
}

function build(s) {
  const ch = [];
  ch.push(new Paragraph({
    children: [new TextRun({ text: `사도행전 주제별 연속설교 (${s.no}/70) · ${s.part}`, font: F, size: 20, color: "888888" })],
    alignment: AlignmentType.CENTER, spacing: { after: 80 }
  }));
  ch.push(new Paragraph({
    children: [new TextRun({ text: s.title, font: F, size: 40, bold: true, color: "1F3864" })],
    alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 }
  }));
  ch.push(new Paragraph({
    children: [new TextRun({ text: `${s.passage} | ${fmtDate(s.date)}`, font: F, size: 22, color: "888888" })],
    alignment: AlignmentType.CENTER, spacing: { after: 80 }
  }));
  ch.push(new Paragraph({ children: [new TextRun({ text: "■", font: F, size: 22, color: "C00000" })], alignment: AlignmentType.CENTER, spacing: { after: 320 } }));
  ch.push(quoteBlock(`${s.passageText} (${s.passage}, 개역개정)`), empty());

  ch.push(heading1("서론"), ...paras(s.intro));
  ch.push(subLabel("오늘의 빅 아이디어"), quoteBlock(s.bigIdea), empty());

  s.points.forEach((pt) => {
    ch.push(heading1(pt.heading), subLabel(pt.keyVerse), empty());
    ch.push(heading2("본문 해설"), ...paras(pt.exposition), empty());
    ch.push(heading2("성경 예화"), ...paras(pt.bibleIllustration), empty());
    ch.push(heading2("일반 예화"), ...paras(pt.generalIllustration), empty());
    ch.push(heading2("적용"), ...paras(pt.application));
    if (pt.connector) ch.push(connector(pt.connector));
    ch.push(empty());
  });

  ch.push(heading1("결론"), ...paras(s.conclusion.before), quoteBlock(s.conclusion.bigIdea), ...paras(s.conclusion.after), empty());
  ch.push(heading1("마무리 기도"), ...paras(s.prayer));

  return new Document({
    creator: "최영완",
    title: s.title,
    styles: {
      default: { document: { run: { font: F, size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: F }, paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: F }, paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 } },
      ]
    },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: ch }]
  });
}

function fileName(s) {
  const t = s.title.replace(/[\s\\/:*?"<>|.,!'“”‘’]/g, '');
  const ref = s.passage.replace(/^사도행전\s*/, '행').replace(/:/g, '장').replace(/\s/g, '') + '절';
  return `${String(s.no).padStart(2, '0')}_${s.date}_${t}_${ref}.docx`;
}

const dir = __dirname;
const files = process.argv.slice(2).length ? process.argv.slice(2) : fs.readdirSync(path.join(dir, 'json')).filter(f => f.endsWith('.json')).sort().map(f => path.join(dir, 'json', f));
(async () => {
  for (const f of files) {
    const s = JSON.parse(fs.readFileSync(f, 'utf8'));
    const buf = await Packer.toBuffer(build(s));
    const out = path.join(dir, 'docx', fileName(s));
    fs.writeFileSync(out, buf);
    console.log('OK', path.basename(out));
  }
})().catch(e => { console.error(e); process.exit(1); });
