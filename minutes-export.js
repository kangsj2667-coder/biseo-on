/* 주간 회의록 초안 → HWP(.hwpx) / Word(.docx) 파일 만들기
   blocks: [{t:'title'|'h'|'p'|'li'|'note', text}, {t:'table', rows:[[...],...]}] */
(function (root) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // ---------- HWPX (한글) ----------
  const NS = 'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"';
  const CH = { title: 7, h: 8, note: 9, th: 10, p: 0, li: 0, td: 0 };
  let pid = 1;
  const hpP = (text, cp, first) => `<hp:p id="${pid++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${first || ''}<hp:run charPrIDRef="${cp}">${text ? `<hp:t>${esc(text)}</hp:t>` : '<hp:t/>'}</hp:run></hp:p>`;
  function hpTable(rows) {
    const nc = rows[0].length, W = 41954;
    const weights = nc === 5 ? [3, 6, 4, 2, 2] : nc === 4 ? [3, 8, 3, 2] : nc === 3 ? [3, 9, 3] : Array(nc).fill(1);
    const sum = weights.reduce((a, b) => a + b, 0);
    const ws = weights.map(w => Math.floor(W * w / sum));
    const trs = rows.map((r, ri) => '<hp:tr>' + r.map((c, ci) =>
      `<hp:tc name="" header="${ri === 0 ? 1 : 0}" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${hpP(c, ri === 0 ? CH.th : CH.td)}</hp:subList><hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="${ws[ci]}" height="282"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>`).join('') + '</hp:tr>').join('');
    const tw = ws.reduce((a, b) => a + b, 0);
    return `<hp:p id="${pid++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:tbl id="${1000000 + pid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rows.length}" colCnt="${nc}" cellSpacing="0" borderFillIDRef="3" noAdjust="0"><hp:sz width="${tw}" widthRelTo="ABSOLUTE" height="${282 * rows.length}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="283" right="283" top="283" bottom="283"/><hp:inMargin left="510" right="510" top="141" bottom="141"/>${trs}</hp:tbl><hp:t/></hp:run></hp:p>`;
  }
  async function buildHwpx(blocks, headerXml, secPrXml, JSZip, title) {
    pid = 1;
    let body = hpP('', 0, `<hp:run charPrIDRef="0">${secPrXml}</hp:run>`);
    blocks.forEach(b => {
      if (b.t === 'table') { body += hpTable(b.rows); return; }
      if (b.t === 'h') body += hpP('', 0);
      body += hpP(b.t === 'li' ? '• ' + b.text : b.text, CH[b.t] ?? 0);
    });
    const sec = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${body}</hs:sec>`;
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const hpf = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package ${NS} version="" unique-identifier="" id=""><opf:metadata><opf:title>${esc(title)}</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">브리피</opf:meta><opf:meta name="CreatedDate" content="text">${now}</opf:meta><opf:meta name="ModifiedDate" content="text">${now}</opf:meta></opf:metadata><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/><opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`;
    const z = new JSZip();
    z.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });
    z.file('version.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1" buildNumber="0" os="1" xmlVersion="1.5" application="Hancom Office Hangul" appVersion="12, 0, 0, 535 WIN32LEWindows_10"/>');
    z.file('Contents/header.xml', headerXml);
    z.file('Contents/section0.xml', sec);
    z.file('Contents/content.hpf', hpf);
    z.file('settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>');
    z.file('META-INF/container.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>');
    z.file('META-INF/manifest.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>');
    return z.generateAsync({ type: typeof window !== 'undefined' ? 'blob' : 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/hwp+zip' });
  }

  // ---------- DOCX (Word) ----------
  const F = '<w:rFonts w:ascii="Malgun Gothic" w:eastAsia="맑은 고딕" w:hAnsi="Malgun Gothic"/>';
  const wR = (text, o = {}) => `<w:r><w:rPr>${F}${o.b ? '<w:b/>' : ''}${o.color ? `<w:color w:val="${o.color}"/>` : ''}<w:sz w:val="${o.sz || 20}"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  const wP = (text, o = {}) => `<w:p><w:pPr><w:spacing w:before="${o.before || 0}" w:after="${o.after ?? 80}"/>${o.ind ? '<w:ind w:left="300" w:hanging="200"/>' : ''}</w:pPr>${wR(text, o)}</w:p>`;
  function wTable(rows) {
    const nc = rows[0].length;
    const weights = nc === 5 ? [3, 6, 4, 2, 2] : nc === 4 ? [3, 8, 3, 2] : nc === 3 ? [3, 9, 3] : Array(nc).fill(1);
    const sum = weights.reduce((a, b) => a + b, 0), W = 9000;
    const ws = weights.map(w => Math.floor(W * w / sum));
    const bd = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:color="999999"/>`).join('');
    return `<w:tbl><w:tblPr><w:tblW w:w="${W}" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders></w:tblPr><w:tblGrid>${ws.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
      rows.map((r, ri) => `<w:tr>${ri === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${r.map((c, ci) => `<w:tc><w:tcPr><w:tcW w:w="${ws[ci]}" w:type="dxa"/>${ri === 0 ? '<w:shd w:val="clear" w:color="auto" w:fill="EEF2F7"/>' : ''}</w:tcPr>${wP(c, { b: ri === 0, after: 0 })}</w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>' + wP('', { after: 0 });
  }
  async function buildDocx(blocks, JSZip, title) {
    const body = blocks.map(b => {
      if (b.t === 'table') return wTable(b.rows);
      if (b.t === 'title') return wP(b.text, { b: true, sz: 32, after: 200 });
      if (b.t === 'h') return wP(b.text, { b: true, sz: 24, color: '1F3A5F', before: 240, after: 100 });
      if (b.t === 'note') return wP(b.text, { sz: 18, color: '666666' });
      if (b.t === 'li') return wP('• ' + b.text, { ind: true });
      return wP(b.text);
    }).join('');
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1300" w:bottom="1440" w:left="1300" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const z = new JSZip();
    z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
    z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
    z.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(title)}</dc:title><dc:creator>브리피</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created></cp:coreProperties>`);
    z.file('word/document.xml', doc);
    return z.generateAsync({ type: typeof window !== 'undefined' ? 'blob' : 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  root.MinutesExport = { buildHwpx, buildDocx };
  if (typeof module !== 'undefined') module.exports = root.MinutesExport;
})(typeof window !== 'undefined' ? window : globalThis);
