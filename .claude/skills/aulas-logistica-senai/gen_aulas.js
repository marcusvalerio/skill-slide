// Gerador de aulas no estilo SENAI (marinho + off-white, títulos Georgia): node gen_aulas.js dados.json saida.pptx
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const { applyTheme } = require(process.env.APPLY_THEME);

const THEME = {
  name: 'Aulas Logistica', headFontFace: 'Georgia', bodyFontFace: 'Calibri',
  colors: { dk1: '1B2333', lt1: 'FFFFFF', dk2: '0B2D5B', lt2: 'F7F7F4', accent1: '2563C9', accent2: 'B45309',
    accent3: '4F8BEA', accent4: '9FC1F5', accent5: 'E7EEF9', accent6: '5E6675', hlink: '2563C9', folHlink: '0B2D5B' }
};
const HX = { line: 'D9DCE2', zebra: 'F1F2F4', warnBg: 'FDF3E3', lightBlue: 'E7EEF9', paleBlue: '9FC1F5' };

async function main() {
  const [, , dataFile, outFile] = process.argv;
  const D = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
  const CURSO = D.curso || 'Assistente de Operações em Logística';
  const PROF = D.professor || 'Marcus Valério';
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.author = PROF; pres.title = `UC${D.uc} · ${D.nome}`;
  const C = pres.SchemeColor;
  const W = 13.333, H = 7.5, M = 0.55, TW = W - 2 * M;
  const footL = `${CURSO} | UC${D.uc} · ${D.nome}`;
  const n2 = n => String(n).padStart(2, '0');

  const footer = (dark) => [
    { text: { text: footL, options: { x: M, y: H - 0.42, w: 8, h: 0.25, fontSize: 9, color: dark ? HX.paleBlue : C.accent6, margin: 0 } } },
    { text: { text: `Professor ${PROF}   ·`, options: { x: W - M - 4.4, y: H - 0.42, w: 4.0, h: 0.25, fontSize: 9, color: dark ? HX.paleBlue : C.accent6, align: 'right', margin: 0 } } },
  ];
  pres.defineSlideMaster({ title: 'MARINHO', background: { color: C.text2 }, objects: footer(true),
    slideNumber: { x: W - M - 0.38, y: H - 0.42, w: 0.38, h: 0.25, fontSize: 9, color: HX.paleBlue, align: 'right' } });
  pres.defineSlideMaster({ title: 'CLARO', background: { color: C.background2 },
    objects: [{ placeholder: { options: { name: 'title', type: 'title', x: M, y: 0.62, w: TW, h: 0.75, fontFace: 'Georgia', fontSize: 30, color: C.text1, align: 'left', valign: 'middle', margin: 0 }, text: '' } }].concat(footer(false)),
    slideNumber: { x: W - M - 0.38, y: H - 0.42, w: 0.38, h: 0.25, fontSize: 9, color: C.accent6, align: 'right' } });

  const tx = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 15, valign: 'top', fontFace: 'Calibri' }, o));
  const kicker = (s, t, dark) => tx(s, t.toUpperCase(), { x: M, y: 0.32, w: TW, h: 0.25, fontSize: 10, bold: true, color: dark ? HX.paleBlue : C.accent1, charSpacing: 2.5 });
  const T = (s, t) => s.addText(t, { placeholder: 'title' });
  const box = (s, x, y, w, h, fill = 'FFFFFF', name = 'cartao') => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: HX.line, width: 0.75 }, objectName: name });
  const stripe = (s, x, y, h, col) => s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: col }, line: { color: col, width: 0 }, objectName: 'faixa' });
  const card = (s, x, y, w, h, col, title, body, o = {}) => {
    box(s, x, y, w, h, o.fill || 'FFFFFF', o.name || 'cartao'); stripe(s, x, y, h, col);
    tx(s, title, { x: x + 0.25, y: y + 0.18, w: w - 0.4, h: 0.4, fontSize: o.tfs || 16, bold: true, color: C.text1 });
    if (body) tx(s, body, { x: x + 0.25, y: y + 0.62, w: w - 0.4, h: h - 0.75, fontSize: o.bfs || 14, color: C.accent6 });
  };
  const STR = [C.text2, C.accent1, C.text2, C.accent1];
  const pergunta = (s, t, y = 6.35) => tx(s, t, { x: M, y, w: TW, h: 0.45, fontSize: 16, italic: true, color: C.accent1 });
  const atencao = (s, t, y) => { s.addShape(pres.shapes.RECTANGLE, { x: M, y, w: TW, h: 0.55, fill: { color: HX.warnBg }, line: { color: HX.warnBg }, objectName: 'atencao' });
    tx(s, [{ text: 'Atenção  ', options: { bold: true, color: C.accent2 } }, { text: t, options: { color: C.accent2 } }], { x: M + 0.25, y, w: TW - 0.5, h: 0.55, fontSize: 14, valign: 'middle' }); };
  const circle = (s, x, y, d, fill, label, fsz = 13) => { s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: 'numero' });
    tx(s, String(label), { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: fsz, bold: true, color: C.background1 }); };
  const notes = (s, n) => { if (n) s.addNotes(n); };

  // ---------- capa ----------
  pres.addSection({ title: 'Abertura da UC' });
  let s = pres.addSlide({ masterName: 'MARINHO', sectionTitle: 'Abertura da UC' });
  tx(s, `${CURSO} · SENAI`.toUpperCase(), { x: M + 0.3, y: 1.15, w: 10, h: 0.3, fontSize: 11, bold: true, color: HX.paleBlue, charSpacing: 3 });
  tx(s, `Unidade Curricular ${D.uc}`, { x: M + 0.3, y: 2.15, w: 10, h: 0.4, fontSize: 18, color: C.background1 });
  tx(s, D.nome, { x: M + 0.3, y: 2.6, w: 11.5, h: 1.5, fontSize: 44, fontFace: 'Georgia', color: C.background1 });
  tx(s, `${D.horas} horas  ·  ${D.aulas.length} aulas de 4 horas`, { x: M + 0.3, y: 4.15, w: 10, h: 0.35, fontSize: 15, color: HX.paleBlue });
  s.addShape(pres.shapes.RECTANGLE, { x: M + 0.3, y: 4.75, w: 1.6, h: 0.03, fill: { color: C.accent3 }, line: { color: C.accent3 }, objectName: 'divisor' });
  tx(s, `Professor ${PROF}`, { x: M + 0.3, y: 4.95, w: 8, h: 0.4, fontSize: 17, bold: true, color: C.background1 });
  notes(s, D.notas_capa || 'Apresente-se, apresente a unidade curricular e combine os combinados da turma.');

  // ---------- mapa ----------
  s = pres.addSlide({ masterName: 'CLARO', sectionTitle: 'Abertura da UC' });
  kicker(s, `Visão geral · ${D.horas} horas`); T(s, 'Mapa da UC');
  {
    const n = D.aulas.length, cols = n <= 5 ? n : 5, rows = Math.ceil(n / cols), gx = 0.15, gy = 0.15, top = 1.65;
    const avail = (n <= 5 ? 4.3 : 5.2), cw = (TW - gx * (cols - 1)) / cols, ch = (avail - gy * (rows - 1)) / rows;
    D.aulas.forEach((a, i) => {
      const r = Math.floor(i / cols), c = i % cols, x = M + c * (cw + gx), y = top + r * (ch + gy), first = i === 0;
      box(s, x, y, cw, ch, first ? C.text2 : 'FFFFFF', `aula-${i + 1}`);
      tx(s, `Aula ${n2(i + 1)}`, { x: x + 0.18, y: y + 0.15, w: cw - 0.36, h: 0.3, fontSize: 11, bold: true, color: first ? HX.paleBlue : C.accent1 });
      tx(s, a.titulo, { x: x + 0.18, y: y + 0.45, w: cw - 0.36, h: n <= 5 ? 1.5 : ch - 0.55, fontSize: n <= 5 ? 18 : 13, fontFace: 'Georgia', color: first ? C.background1 : C.text1 });
      if (n <= 5) tx(s, a.slides.map(x => x.titulo).join(' · '), { x: x + 0.18, y: y + 2.1, w: cw - 0.36, h: ch - 2.25, fontSize: 11, color: first ? HX.paleBlue : C.accent6 });
    });
  }
  notes(s, 'Mostre o caminho da UC: o que vem em cada aula e como as aulas se conectam.');

  // ---------- aulas ----------
  for (let ai = 0; ai < D.aulas.length; ai++) {
    const A = D.aulas[ai], num = ai + 1, sec = `Aula ${n2(num)} · ${A.titulo}`;
    const ct = A.slides.map(x => x.titulo), half = Math.ceil(ct.length / 2);
    const AT = A.atividade, mod = (AT.modo || 'grupo') === 'individual' ? 'individual' : 'em grupo';
    const bnames = A.blocos || ['Bloco 1', 'Bloco 2'];
    const rot = ['Abertura e revisão', `${bnames[0]}`, 'Intervalo', `${bnames[1]}`, `Fixação ${mod}: ${AT.titulo}`, 'Questões de fixação e síntese'];
    pres.addSection({ title: sec });

    // abertura (marinho)
    s = pres.addSlide({ masterName: 'MARINHO', sectionTitle: sec });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 0.75, w: 0.08, h: 1.75, fill: { color: C.accent3 }, line: { color: C.accent3 }, objectName: 'marca' });
    tx(s, `AULA ${n2(num)}`, { x: M + 0.3, y: 0.78, w: 6, h: 0.3, fontSize: 11, bold: true, color: HX.paleBlue, charSpacing: 3 });
    tx(s, A.titulo, { x: M + 0.3, y: 1.1, w: 11.8, h: 1.4, fontSize: 36, fontFace: 'Georgia', color: C.background1, valign: 'middle' });
    tx(s, 'Objetivos de aprendizagem', { x: M + 0.3, y: 2.95, w: 6, h: 0.35, fontSize: 16, bold: true, color: HX.paleBlue });
    tx(s, A.objetivos.map((o, i) => ({ text: o, options: { bullet: true, breakLine: i < A.objetivos.length - 1 } })), { x: M + 0.3, y: 3.4, w: 6.4, h: 3.2, fontSize: 15, color: C.background1, paraSpaceAfter: 6 });
    tx(s, 'Roteiro da aula', { x: 7.6, y: 2.95, w: 5, h: 0.35, fontSize: 16, bold: true, color: HX.paleBlue });
    rot.forEach((r, i) => {
      const y = 3.4 + i * 0.5;
      s.addShape(pres.shapes.OVAL, { x: 7.6, y: y + 0.15, w: 0.14, h: 0.14, fill: { color: i === 2 ? HX.paleBlue : C.accent3 }, line: { color: C.accent3 }, objectName: 'marcador' });
      tx(s, r, { x: 7.95, y, w: W - M - 7.95, h: 0.45, fontSize: 14, color: i === 2 ? HX.paleBlue : C.background1, valign: 'middle', fit: 'shrink' });
    });
    notes(s, (ai === 0 ? 'Abertura: apresente a UC e faça uma sondagem do que a turma já sabe. ' : 'Abertura: retome a aula anterior em 2 ou 3 perguntas. ') + 'Apresente os objetivos e o roteiro do dia.');

    // conteúdo
    A.slides.forEach((sl, si) => {
      const bloco = sl.bloco || (si < half ? 1 : 2);
      const prev = si === 0 ? 0 : (A.slides[si - 1].bloco || (si - 1 < half ? 1 : 2));
      if (bloco !== prev) {
        const lst = A.slides.filter((x, xi) => (x.bloco || (xi < half ? 1 : 2)) === bloco).map(x => x.titulo);
        s = pres.addSlide({ masterName: 'MARINHO', sectionTitle: sec });
        s.addShape(pres.shapes.RECTANGLE, { x: M, y: 1.4, w: 0.08, h: 1.3, fill: { color: C.accent3 }, line: { color: C.accent3 }, objectName: 'marca' });
        tx(s, `AULA ${n2(num)} · BLOCO ${bloco}`, { x: M + 0.3, y: 1.42, w: 8, h: 0.3, fontSize: 11, bold: true, color: HX.paleBlue, charSpacing: 3 });
        tx(s, bnames[bloco - 1], { x: M + 0.3, y: 1.75, w: 11.8, h: 0.95, fontSize: 34, fontFace: 'Georgia', color: C.background1, valign: 'middle' });
        tx(s, lst.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < lst.length - 1 } })), { x: M + 0.3, y: 3.2, w: 11.5, h: 3.4, fontSize: 16, color: C.background1, paraSpaceAfter: 4 });
        notes(s, bloco === 1 ? 'Apresente o que será visto neste bloco.' : 'Retorno do intervalo: retome em uma frase o Bloco 1 e apresente o Bloco 2.');
      }
      s = pres.addSlide({ masterName: 'CLARO', sectionTitle: sec });
      kicker(s, `Aula ${n2(num)} · Bloco ${bloco}`); T(s, sl.titulo);
      const top = 1.65;
      if (sl.tipo === 'conceito') {
        tx(s, sl.definicao, { x: M, y: top, w: TW, h: 1.05, fontSize: 18, color: C.text1 });
        const n = sl.pontos.length, gx = 0.15, cw = (TW - gx * (n - 1)) / n;
        sl.pontos.forEach((p, i) => card(s, M + i * (cw + gx), top + 1.35, cw, 2.5, STR[i % 4], p[0], p[1], { bfs: 15 }));
      } else if (sl.tipo === 'cards') {
        const n = sl.itens.length, gx = 0.15, cw = (TW - gx * (n - 1)) / n;
        sl.itens.forEach((p, i) => card(s, M + i * (cw + gx), top, cw, 3.4, STR[i % 4], p[0], p[1], { tfs: 17, bfs: 15 }));
      } else if (sl.tipo === 'grade') {
        const n = sl.itens.length, cols = n > 6 ? 4 : (n === 4 ? 2 : Math.min(3, n)), rows = Math.ceil(n / cols), gx = 0.15, gy = 0.15;
        const cw = (TW - gx * (cols - 1)) / cols, ch = Math.min(2.0, (4.1 - gy * (rows - 1)) / rows);
        sl.itens.forEach((p, i) => { const r = Math.floor(i / cols), c = i % cols;
          card(s, M + c * (cw + gx), top + r * (ch + gy), cw, ch, r % 2 ? C.accent1 : C.accent2, p[0], p[1], { bfs: 15 }); });
      } else if (sl.tipo === 'processo') {
        const n = sl.passos.length, gap = 0.32, bw = (TW - gap * (n - 1)) / n, bh = 0.95, y = top + 0.3;
        sl.passos.forEach((p, i) => {
          const x = M + i * (bw + gap), hi = i === n - 1;
          s.addShape(pres.shapes.RECTANGLE, { x, y, w: bw, h: bh, fill: { color: hi ? C.text2 : 'FFFFFF' }, line: { color: C.accent1, width: 1.25 }, objectName: `etapa-${i + 1}` });
          tx(s, p[0], { x: x + 0.1, y, w: bw - 0.2, h: bh, fontSize: n > 5 ? 14 : 16, bold: true, align: 'center', valign: 'middle', color: hi ? C.background1 : C.text1 });
          if (i < n - 1) s.addShape(pres.shapes.LINE, { x: x + bw + 0.04, y: y + bh / 2, w: gap - 0.08, h: 0, line: { color: C.accent1, width: 1.5, endArrowType: 'triangle' }, objectName: `seta-${i + 1}` });
          tx(s, p[1], { x, y: y + bh + 0.2, w: bw, h: 1.0, fontSize: n > 5 ? 13 : 15, align: 'center', color: C.accent6 });
        });
        s.addShape(pres.shapes.RECTANGLE, { x: M, y: top + 2.75, w: TW, h: 0.55, fill: { color: HX.lightBlue }, line: { color: HX.lightBlue }, objectName: 'faixa-info' });
        tx(s, `${n} etapas em sequência: cada uma depende da anterior estar correta`, { x: M, y: top + 2.75, w: TW, h: 0.55, fontSize: 14, bold: true, color: C.text2, align: 'center', valign: 'middle' });
      } else if (sl.tipo === 'tabela') {
        const nl = sl.linhas.length, fsz = nl > 6 ? 13 : (nl > 3 ? 14 : 16);
        const rows = [sl.cab.map(c => ({ text: c, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: fsz } }))]
          .concat(sl.linhas.map((l, ri) => l.map((c, ci) => ({ text: c, options: { fontSize: fsz, color: C.text1, bold: ci === 0, fill: { color: ri % 2 ? HX.zebra : 'FFFFFF' } } }))));
        s.addTable(rows, { x: M, y: top, w: TW, colW: sl.colW || Array(sl.cab.length).fill(TW / sl.cab.length), border: { type: 'solid', pt: 0.75, color: HX.line }, margin: [0.06, 0.1, 0.06, 0.1], valign: 'middle', rowH: Math.min(0.62, 4.4 / (nl + 1)), autoPage: false, fontFace: 'Calibri' });
      } else if (sl.tipo === 'comparar') {
        const cw = (TW - 0.3) / 2;
        [[sl.a, C.text2], [sl.b, C.accent1]].forEach(([side, col], k) => {
          const x = M + k * (cw + 0.3);
          box(s, x, top, cw, 4.2, 'FFFFFF', `lado-${k + 1}`); stripe(s, x, top, 4.2, col);
          tx(s, side[0], { x: x + 0.3, y: top + 0.25, w: cw - 0.5, h: 0.5, fontSize: 20, fontFace: 'Georgia', color: C.text1 });
          tx(s, side[1].map((o, i) => ({ text: o, options: { bullet: true, breakLine: i < side[1].length - 1 } })), { x: x + 0.3, y: top + 0.95, w: cw - 0.6, h: 3.0, fontSize: 16, color: C.text1, paraSpaceAfter: 8 });
        });
      } else if (sl.tipo === 'destaque') {
        s.addShape(pres.shapes.RECTANGLE, { x: M, y: top, w: 4.4, h: 2.3, fill: { color: C.text2 }, line: { color: C.text2 }, objectName: 'numero-destaque' });
        tx(s, sl.numero, { x: M, y: top, w: 4.4, h: 2.3, fontSize: 48, fontFace: 'Georgia', color: C.background1, align: 'center', valign: 'middle' });
        tx(s, sl.rotulo, { x: M, y: top + 2.45, w: 4.4, h: 0.6, fontSize: 14, color: C.accent6, align: 'center' });
        tx(s, sl.texto, { x: M + 4.9, y: top, w: TW - 4.9, h: 2.3, fontSize: 18, color: C.text1, valign: 'middle' });
      } else if (sl.tipo === 'pratica') {
        box(s, M, top, TW, 2.6, 'FFFFFF', 'situacao'); stripe(s, M, top, 2.6, C.accent2);
        tx(s, 'CASO PARA DISCUSSÃO', { x: M + 0.3, y: top + 0.22, w: 6, h: 0.3, fontSize: 11, bold: true, color: C.accent2, charSpacing: 2 });
        tx(s, sl.caso, { x: M + 0.3, y: top + 0.6, w: TW - 0.6, h: 1.85, fontSize: 19, color: C.text1 });
        tx(s, [{ text: 'Pergunta para a turma: ', options: { bold: true } }, { text: sl.pergunta }], { x: M, y: top + 3.0, w: TW, h: 0.9, fontSize: 18, italic: true, color: C.accent1 });
      } else if (sl.tipo === 'exemplo') {
        const lw = 7.0, nl = sl.linhas.length, fsz = nl > 7 ? 12 : 14;
        const rows = [sl.cab.map(c => ({ text: c, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: fsz } }))]
          .concat(sl.linhas.map((l, ri) => l.map(c => ({ text: c, options: { fontSize: fsz, color: C.text1, fill: { color: ri % 2 ? HX.zebra : 'FFFFFF' } } }))));
        s.addTable(rows, { x: M, y: top, w: lw, colW: Array(sl.cab.length).fill(lw / sl.cab.length), border: { type: 'solid', pt: 0.75, color: HX.line }, margin: [0.05, 0.08, 0.05, 0.08], valign: 'middle', rowH: Math.min(0.45, 4.1 / (nl + 1)), autoPage: false, fontFace: 'Calibri' });
        const rx = M + lw + 0.35, rw = TW - lw - 0.35;
        tx(s, (sl.passos || []).map((p, i) => ({ text: `${i + 1}. ${p}`, options: { breakLine: i < sl.passos.length - 1 } })), { x: rx, y: top, w: rw, h: 2.6, fontSize: 15, paraSpaceAfter: 6 });
        if (sl.resultado) { s.addShape(pres.shapes.RECTANGLE, { x: rx, y: top + 2.8, w: rw, h: 1.2, fill: { color: HX.lightBlue }, line: { color: C.accent1, width: 1 }, objectName: 'resultado' });
          tx(s, sl.resultado, { x: rx + 0.2, y: top + 2.8, w: rw - 0.4, h: 1.2, fontSize: 16, bold: true, color: C.text2, valign: 'middle' }); }
      } else if (sl.tipo === 'visual') {
        const n = sl.blocos.length, bw = Math.min(1.9, (TW - 0.2 * (n - 1)) / n);
        sl.blocos.forEach((b, i) => { const x = M + i * (bw + 0.2);
          s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: bw, h: 1.6, fill: { color: C.text2 }, line: { color: C.text2 }, objectName: `bloco-${i + 1}` });
          tx(s, b[0], { x, y: top, w: bw, h: 1.6, fontSize: 44, fontFace: 'Georgia', color: C.background1, align: 'center', valign: 'middle' });
          tx(s, b[1], { x, y: top + 1.7, w: bw, h: 0.4, fontSize: 13, color: C.accent6, align: 'center' }); });
        tx(s, sl.texto, { x: M, y: top + 2.45, w: TW, h: 1.5, fontSize: 18, color: C.text1 });
      }
      if (sl.atencao) atencao(s, sl.atencao, sl.pergunta ? 5.75 : 6.3);
      if (sl.pergunta && sl.tipo !== 'pratica') pergunta(s, (sl.pergunta.startsWith('Pergunta') ? '' : 'Pergunta para a turma: ') + sl.pergunta, 6.35);
      notes(s, sl.notas);
    });

    // atividade
    s = pres.addSlide({ masterName: 'CLARO', sectionTitle: sec });
    kicker(s, `Aula ${n2(num)} · Fixação ${mod}${AT.formato ? ' · ' + AT.formato.toLowerCase() : ''}`);
    T(s, `Atividade de fixação: ${AT.titulo}`);
    const sp = AT.tabela ? 0.5 : 0.68;
    AT.passos.forEach((p, i) => { const y = 1.7 + i * sp;
      circle(s, M, y, 0.38, C.text2, i + 1, 12); tx(s, p, { x: M + 0.6, y: y - 0.04, w: TW - 0.7, h: 0.46, fontSize: AT.tabela ? 14 : 16, valign: 'middle' }); });
    if (AT.tabela) { const ty = 1.75 + AT.passos.length * sp, nl = AT.tabela.linhas.length;
      const rows = [AT.tabela.cab.map(c => ({ text: c, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 13 } }))]
        .concat(AT.tabela.linhas.map((l, ri) => AT.tabela.cab.map((c, ci) => ({ text: l[ci] || '', options: { fontSize: 13, color: C.text1, fill: { color: ri % 2 ? HX.zebra : 'FFFFFF' } } }))));
      s.addTable(rows, { x: M, y: ty, w: TW, border: { type: 'solid', pt: 0.75, color: HX.line }, margin: [0.04, 0.08, 0.04, 0.08], rowH: Math.min(0.36, (5.85 - ty) / (nl + 1)), autoPage: false, fontFace: 'Calibri' }); }
    tx(s, [{ text: 'Entrega: ', options: { bold: true, color: C.text2 } }, { text: AT.entrega, options: { color: C.text2, bold: true } }], { x: M, y: 6.3, w: TW, h: 0.45, fontSize: 15 });
    notes(s, AT.notas || 'Circule entre os grupos, faça perguntas e anote dúvidas para as questões de fixação.');

    // questões
    s = pres.addSlide({ masterName: 'CLARO', sectionTitle: sec });
    kicker(s, `Aula ${n2(num)}`); T(s, 'Questões de fixação');
    A.quiz.forEach((q, i) => { const y = 1.75 + i * 0.95;
      circle(s, M, y, 0.48, C.text2, i + 1, 14); tx(s, q, { x: M + 0.75, y: y - 0.05, w: TW - 0.8, h: 0.6, fontSize: 18, valign: 'middle' }); });
    notes(s, 'Peça respostas em voz alta ou num papel e use as dúvidas para retomar pontos.' + (A.respostas ? ' Respostas esperadas: ' + A.respostas : ''));

    // síntese
    s = pres.addSlide({ masterName: 'CLARO', sectionTitle: sec });
    kicker(s, `Aula ${n2(num)} · O que levamos`); T(s, `Síntese da Aula ${n2(num)}`);
    {
      const items = ct.slice(0, 5).map((t, i) => [n2(i + 1), t, false]);
      const nx = D.aulas[ai + 1];
      items.push(['→', nx ? `Próxima aula: ${nx.titulo}` : 'Fim da UC: revise o material e traga suas dúvidas', true]);
      const cols = 3, gx = 0.15, gy = 0.15, cw = (TW - gx * 2) / cols, ch = 1.65;
      items.forEach((it, i) => { const r = Math.floor(i / cols), c = i % cols, x = M + c * (cw + gx), y = 1.7 + r * (ch + gy);
        box(s, x, y, cw, ch, it[2] ? HX.lightBlue : 'FFFFFF', `sintese-${i + 1}`);
        tx(s, it[0], { x: x + 0.25, y: y + 0.18, w: 1, h: 0.4, fontSize: 18, fontFace: 'Georgia', color: C.accent1 });
        tx(s, it[1], { x: x + 0.25, y: y + 0.65, w: cw - 0.5, h: ch - 0.8, fontSize: 15, color: C.text1 }); });
    }
    notes(s, 'Feche retomando os pontos principais e antecipe o tema da próxima aula.');
  }

  await pres.writeFile({ fileName: outFile });
  await applyTheme(outFile, THEME);
  console.log('ok', outFile);
}
main().catch(e => { console.error(e); process.exit(1); });
