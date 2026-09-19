const pptxgen = require('pptxgenjs');
const { iconData } = require('./icons');
const fa = require('react-icons/fa');
const gi = require('react-icons/gi');

// ---------- palette ----------
const INK        = '0F2B3D';
const INK_SOFT   = '1B3E52';
const PAPER      = 'F4F7F9';
const WHITE      = 'FFFFFF';
const TEAL       = '1C6E8C';
const TEAL_DEEP  = '14536B';
const TEAL_TINT  = 'DCEAF1';
const ORANGE     = 'D9601F';
const ORANGE_TINT= 'FBE6D9';
const SLATE      = '55707F';
const SLATE_LT   = '8BA3AF';
const BORDER     = 'DCE6EC';
const GREEN      = '2E7D63';
const GREEN_TINT = 'DFF0E8';
const AMBER_TINT = 'FBF1DC';
const AMBER      = '9C6B16';
const GREY_BAR   = 'B9C6CD';

const K = '맑은 고딕';

const W = 13.333, H = 7.5;
const M = 0.55;                  // side margin

// ---------- helpers (fresh objects every call) ----------
const cardShadow = () => ({ type: 'outer', color: INK, opacity: 0.10, blur: 10, offset: 2, angle: 90 });

function card(slide, pres, x, y, w, h, opts) {
  opts = opts || {};
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: opts.fill || WHITE },
    line: { color: opts.line || BORDER, width: 1 },
    rectRadius: 0.09,
    shadow: opts.shadow === false ? undefined : cardShadow(),
  });
}

function slideTitle(slide, text, opts) {
  opts = opts || {};
  slide.addText([{ text, options: {} }], {
    x: M, y: opts.y === undefined ? 0.42 : opts.y, w: W - 2 * M, h: 0.85,
    isTextBox: true, margin: 0,
    fontFace: K, fontSize: opts.size || 33, bold: true,
    color: opts.color || INK, align: 'left', valign: 'middle',
  });
}

function kicker(slide, text, opts) {
  opts = opts || {};
  slide.addText([{ text, options: {} }], {
    x: M, y: opts.y === undefined ? 1.24 : opts.y, w: W - 2 * M, h: 0.34,
    isTextBox: true, margin: 0,
    fontFace: K, fontSize: 14.5, color: opts.color || SLATE, align: 'left', valign: 'middle',
  });
}

function footnote(slide, text, opts) {
  opts = opts || {};
  slide.addText([{ text, options: {} }], {
    x: M, y: opts.y === undefined ? 6.92 : opts.y, w: W - 2 * M, h: 0.34,
    isTextBox: true, margin: 0,
    fontFace: K, fontSize: 9.5, italic: true,
    color: opts.color || SLATE_LT, align: 'left', valign: 'middle',
  });
}

// icon sitting inside a tinted circle
function iconBadge(slide, pres, cx, cy, d, data, opts) {
  opts = opts || {};
  slide.addShape(pres.ShapeType.ellipse, {
    x: cx - d / 2, y: cy - d / 2, w: d, h: d,
    fill: { color: opts.bg || TEAL_TINT }, line: { color: opts.bg || TEAL_TINT, width: 0 },
  });
  const s = d * (opts.scale || 0.52);
  slide.addImage({ data, x: cx - s / 2, y: cy - s / 2, w: s, h: s });
}

function numDot(slide, pres, x, y, d, label, opts) {
  opts = opts || {};
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d,
    fill: { color: opts.fill || ORANGE }, line: { color: opts.fill || ORANGE, width: 0 },
  });
  slide.addText([{ text: label, options: {} }], {
    x, y, w: d, h: d, isTextBox: true, margin: 0,
    fontFace: K, fontSize: opts.size || 13, bold: true, color: WHITE,
    align: 'center', valign: 'middle',
  });
}

// ================================================================
async function build() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'GLP-1 환자 교육 자료';
  pres.title = 'GLP-1 작용제는 우리 몸에서 무슨 일을 할까?';

  // ---------- icons ----------
  const I = {};
  const jobs = [
    ['brain',   fa.FaBrain,      TEAL],
    ['stomach', gi.GiStomach,    TEAL],
    ['drop',    fa.FaTint,       TEAL],
    ['liver',   gi.GiLiver,      TEAL],
    ['food',    fa.FaUtensils,   TEAL],
    ['chem',    gi.GiChemicalDrop, ORANGE],
    ['clock',   fa.FaClock,      ORANGE],
    ['pills',   gi.GiMedicinePills, TEAL],
    ['syringe', fa.FaSyringe,    TEAL],
    ['muscle',  gi.GiMuscleUp,   TEAL],
    ['weight',  fa.FaWeight,     TEAL],
    ['heart',   fa.FaHeartbeat,  WHITE],
    ['warn',    fa.FaExclamationTriangle, AMBER],
    ['check',   fa.FaCheck,      GREEN],
    ['quest',   fa.FaQuestion,   AMBER],
    ['glass',   gi.GiHourglass,  TEAL],
    ['sugar',   gi.GiSugarCane,  TEAL],
    ['up',      fa.FaArrowUp,    ORANGE],
    ['down',    fa.FaArrowDown,  TEAL],
    ['brainW',  fa.FaBrain,      WHITE],
  ];
  for (const [k, C, col] of jobs) I[k] = await iconData(C, col);

  // ============================================================
  // SLIDE 1 — title (dark)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: INK };

    // motif: the GLP-1 "molecule" — a big soft orange circle + solid dot
    s.addShape(pres.ShapeType.ellipse, {
      x: 9.55, y: 0.75, w: 5.6, h: 5.6,
      fill: { color: ORANGE, transparency: 86 }, line: { color: 'A4522A', width: 1.25 },
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: 11.15, y: 2.35, w: 2.4, h: 2.4,
      fill: { color: ORANGE, transparency: 72 }, line: { color: 'A4522A', width: 1.25 },
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: 12.0, y: 3.2, w: 0.7, h: 0.7,
      fill: { color: ORANGE }, line: { type: 'none' },
    });

    s.addText([{ text: '일반인을 위한 그림 설명', options: {} }], {
      x: M, y: 1.65, w: 8.6, h: 0.4, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 15, bold: true, color: 'E9A87C', charSpacing: 1.5,
      align: 'left', valign: 'middle',
    });
    s.addText([
      { text: 'GLP-1 작용제는', options: { breakLine: true } },
      { text: '우리 몸에서 무슨 일을 할까?', options: {} },
    ], {
      x: M, y: 2.15, w: 8.8, h: 2.0, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 40, bold: true, color: WHITE, lineSpacingMultiple: 1.2,
      align: 'left', valign: 'top',
    });
    s.addText([{ text: '식사 후 몸이 보내는 “이제 그만 먹어도 돼” 신호를,\n약이 어떻게 오래 켜 두는지 그림으로 따라갑니다.', options: {} }], {
      x: M, y: 4.35, w: 8.3, h: 1.0, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 15, color: 'AFC6D2', lineSpacingMultiple: 1.45,
      align: 'left', valign: 'top',
    });
    s.addText([{ text: '교육용 자료 · 2026년 9월 · 근거: 주요 무작위배정 임상시험 및 종설', options: {} }], {
      x: M, y: 6.45, w: 9.0, h: 0.35, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 11, color: '6E8B99', align: 'left', valign: 'middle',
    });
    s.addNotes('도입: GLP-1 작용제를 "살 빼는 신약"으로만 아는 청중이 많습니다. 이 발표의 목표는 약이 무엇을 흉내내는지(우리 몸의 식후 호르몬), 어디에 작용하는지(뇌·위·췌장·간), 그리고 효과와 부작용이 그 기전에서 어떻게 따라 나오는지를 연결해 보여주는 것입니다.');
  }

  // ============================================================
  // SLIDE 2 — GLP-1 is your own hormone (flow)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '먼저, GLP-1은 원래 내 몸이 만드는 호르몬입니다');
    kicker(s, '약 이름이 아니라, 식사를 하면 창자가 뿜어내는 “식후 신호 호르몬”의 이름입니다.');

    const steps = [
      { n: '1', ic: I.food,  h: '음식을 먹는다',       b: '음식이 위를 지나 작은창자에 닿습니다.' },
      { n: '2', ic: I.stomach, h: '창자가 알아챈다',   b: '작은창자·큰창자 벽의 L세포가 영양분을 감지합니다.' },
      { n: '3', ic: I.chem,  h: 'GLP-1이 나온다',      b: 'L세포가 GLP-1 호르몬을 혈액으로 내보냅니다.', accent: true },
      { n: '4', ic: I.brain, h: '뇌·췌장에 전달된다',  b: '“충분히 먹었다”는 신호가 온몸으로 퍼집니다.' },
    ];
    const gap = 0.52;
    const cw = (W - 2 * M - gap * 3) / 4;
    const cy = 2.05, ch = 2.95;

    steps.forEach((st, i) => {
      const x = M + i * (cw + gap);
      card(s, pres, x, cy, cw, ch, { fill: st.accent ? ORANGE_TINT : WHITE, line: st.accent ? 'F0CDB6' : BORDER });
      numDot(s, pres, x + 0.2, cy + 0.2, 0.4, st.n, { fill: st.accent ? ORANGE : TEAL });
      iconBadge(s, pres, x + cw / 2, cy + 1.02, 0.94, st.ic, { bg: st.accent ? WHITE : TEAL_TINT });
      s.addText([{ text: st.h, options: {} }], {
        x: x + 0.14, y: cy + 1.62, w: cw - 0.28, h: 0.4, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 16, bold: true, color: st.accent ? ORANGE : TEAL_DEEP,
        align: 'center', valign: 'middle',
      });
      s.addText([{ text: st.b, options: {} }], {
        x: x + 0.18, y: cy + 2.04, w: cw - 0.36, h: 0.78, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 11.5, color: SLATE, lineSpacingMultiple: 1.3,
        align: 'center', valign: 'top',
      });
      if (i < 3) {
        s.addShape(pres.ShapeType.rightArrow, {
          x: x + cw + 0.10, y: cy + 1.30, w: 0.32, h: 0.30,
          fill: { color: SLATE_LT }, line: { color: SLATE_LT, width: 0 },
        });
      }
    });

    // the catch
    const by = 5.30;
    card(s, pres, M, by, W - 2 * M, 1.20, { fill: AMBER_TINT, line: 'EBD9B4' });
    iconBadge(s, pres, M + 0.62, by + 0.60, 0.68, I.clock, { bg: WHITE, scale: 0.5 });
    s.addText([
      { text: '그런데 — 천연 GLP-1은 1~2분이면 사라집니다.  ', options: { bold: true, color: AMBER } },
      { text: '혈액 속 DPP-4라는 효소가 즉시 잘라 버리기 때문입니다. 그래서 내 몸이 만드는 GLP-1을 그대로 약으로 쓸 수는 없습니다.', options: { color: INK_SOFT } },
    ], {
      x: M + 1.10, y: by + 0.20, w: W - 2 * M - 1.35, h: 0.80, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 13, lineSpacingMultiple: 1.35, align: 'left', valign: 'middle',
    });

    footnote(s, '그림은 개념 설명을 위한 단순화 도식입니다. 출처: Drucker DJ. Cell Metab 2018; Müller TD 외. Mol Metab 2019.');
    s.addNotes('핵심 메시지: GLP-1은 외부 물질이 아니라 우리 몸의 정상 식후 호르몬입니다. 인크레틴(incretin)이라 부릅니다. 반감기가 1~2분에 불과해(DPP-4에 의한 분해) 천연 호르몬 자체는 치료제가 될 수 없다는 점이 다음 장의 출발점입니다.');
  }

  // ============================================================
  // SLIDE 3 — the drug is a long-lasting copy
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '약은 “오래 버티도록 고친 복제품”입니다');
    kicker(s, '새로운 작용을 더한 것이 아니라, 같은 신호가 며칠 동안 유지되도록 구조만 바꿔 놓은 것입니다.');

    const pw = (W - 2 * M - 0.40) / 2;
    const py = 1.80, ph = 2.35;

    // left: natural
    card(s, pres, M, py, pw, ph, { fill: WHITE });
    s.addText([{ text: '내 몸의 천연 GLP-1', options: {} }], {
      x: M + 0.35, y: py + 0.30, w: pw - 0.7, h: 0.38, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 15, bold: true, color: SLATE, align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '1~2', options: { fontSize: 56, bold: true, color: SLATE } },
      { text: '  분', options: { fontSize: 22, bold: true, color: SLATE_LT } },
    ], {
      x: M + 0.35, y: py + 0.72, w: pw - 0.7, h: 0.95, isTextBox: true, margin: 0,
      fontFace: K, align: 'left', valign: 'middle',
    });
    s.addText([{ text: 'DPP-4 효소에 바로 잘려 사라집니다.', options: {} }], {
      x: M + 0.35, y: py + 1.72, w: pw - 0.7, h: 0.4, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12, color: SLATE, align: 'left', valign: 'middle',
    });

    // right: drug
    const rx = M + pw + 0.40;
    s.addShape(pres.ShapeType.roundRect, {
      x: rx, y: py, w: pw, h: ph,
      fill: { color: TEAL }, line: { color: TEAL, width: 0 }, rectRadius: 0.09,
      shadow: cardShadow(),
    });
    s.addText([{ text: '주 1회 세마글루타이드 (오젬픽·위고비)', options: {} }], {
      x: rx + 0.35, y: py + 0.30, w: pw - 0.7, h: 0.38, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 15, bold: true, color: 'BCDCE9', align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '약 7', options: { fontSize: 56, bold: true, color: WHITE } },
      { text: '  일', options: { fontSize: 22, bold: true, color: 'BCDCE9' } },
    ], {
      x: rx + 0.35, y: py + 0.72, w: pw - 0.7, h: 0.95, isTextBox: true, margin: 0,
      fontFace: K, align: 'left', valign: 'middle',
    });
    s.addText([{ text: '반감기 약 165시간 → 주 1회 주사로 농도가 유지됩니다.', options: {} }], {
      x: rx + 0.35, y: py + 1.72, w: pw - 0.7, h: 0.4, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12, color: 'D6EAF2', align: 'left', valign: 'middle',
    });

    // how
    s.addText([{ text: '어떻게 오래 버티게 만들었나', options: {} }], {
      x: M, y: 4.35, w: W - 2 * M, h: 0.34, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 13.5, bold: true, color: INK, align: 'left', valign: 'middle',
    });

    const hows = [
      { n: '1', h: '잘리는 자리를 바꿈', b: 'DPP-4가 자르던 부위를 살짝 변형해 분해를 피합니다.' },
      { n: '2', h: '알부민에 업혀 다님', b: '지방산 사슬을 붙여 혈액 단백질에 붙어 다니므로 콩팥으로 덜 빠져나갑니다.' },
      { n: '3', h: '그래서 주 1회', b: '주 1회 주사, 또는 매일 먹는 경구제로도 농도를 유지할 수 있습니다.' },
    ];
    const hg = 0.32, hw = (W - 2 * M - hg * 2) / 3;
    const hy = 4.78, hh = 1.95;
    hows.forEach((it, i) => {
      const x = M + i * (hw + hg);
      card(s, pres, x, hy, hw, hh, { fill: WHITE });
      numDot(s, pres, x + 0.28, hy + 0.28, 0.42, it.n, { fill: TEAL });
      s.addText([{ text: it.h, options: {} }], {
        x: x + 0.82, y: hy + 0.26, w: hw - 1.10, h: 0.46, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 15, bold: true, color: TEAL_DEEP, align: 'left', valign: 'middle',
      });
      s.addText([{ text: it.b, options: {} }], {
        x: x + 0.30, y: hy + 0.84, w: hw - 0.60, h: 0.95, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 12, color: SLATE, lineSpacingMultiple: 1.35,
        align: 'left', valign: 'top',
      });
    });

    footnote(s, '수치는 세마글루타이드 기준이며 약제마다 다릅니다(리라글루타이드는 매일 주사). 출처: Müller TD 외. Mol Metab 2019.');
    s.addNotes('두 가지 변형이 핵심입니다: (1) DPP-4 절단 부위 변형, (2) 지방산 측쇄를 통한 알부민 결합으로 신장 배설 지연. 결과적으로 반감기가 1~2분에서 약 1주로 늘어납니다. 작용 자체는 천연 호르몬과 같은 GLP-1 수용체를 통해 일어난다는 점을 강조하십시오.');
  }

  // ============================================================
  // SLIDE 4 — generic names / brand names
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '실제로 어떤 약들인가 — 성분명과 상품명');
    kicker(s, '성분명이 약의 본명이고, 상품명은 회사가 붙인 이름입니다. 같은 성분이 용도에 따라 다른 이름으로 팔립니다.');

    const hdr = (t) => ({ text: t, options: { fill: { color: TEAL }, color: WHITE, bold: true, fontSize: 12, fontFace: K, align: 'left', valign: 'middle', margin: [0, 0.14, 0, 0.14] } });
    const cell = (t, o) => Object.assign({
      text: t,
      options: Object.assign({
        fontSize: 11.5, fontFace: K, color: INK_SOFT, valign: 'middle',
        align: 'left', margin: [0, 0.14, 0, 0.14], fill: { color: WHITE },
      }, o || {}),
    });

    const OB = ORANGE_TINT;   // rows with an obesity (weight) indication
    const rows = [
      [hdr('성분명'), hdr('상품명'), hdr('투여 방법'), hdr('주요 허가 적응증')],
      [
        cell('세마글루타이드', { bold: true, color: TEAL_DEEP, fill: { color: OB } }),
        cell('위고비 (Wegovy)', { bold: true, fill: { color: OB } }),
        cell('주 1회 피하주사', { fill: { color: OB } }),
        cell('체중 감량 (비만)', { bold: true, color: ORANGE, fill: { color: OB } }),
      ],
      [
        cell('세마글루타이드', { bold: true, color: TEAL_DEEP }),
        cell('오젬픽 (Ozempic)', { bold: true }),
        cell('주 1회 피하주사'),
        cell('2형 당뇨'),
      ],
      [
        cell('세마글루타이드', { bold: true, color: TEAL_DEEP }),
        cell('리벨서스 (Rybelsus)', { bold: true }),
        cell('매일 1회 먹는 알약', { color: TEAL_DEEP, bold: true }),
        cell('2형 당뇨'),
      ],
      [
        cell('터제파타이드', { bold: true, color: TEAL_DEEP, fill: { color: OB } }),
        cell('마운자로 (Mounjaro)', { bold: true, fill: { color: OB } }),
        cell('주 1회 피하주사', { fill: { color: OB } }),
        cell('2형 당뇨 + 체중 감량  ※GLP-1과 GIP에 함께 작용', { bold: true, color: ORANGE, fill: { color: OB } }),
      ],
      [
        cell('리라글루타이드', { bold: true, color: TEAL_DEEP, fill: { color: OB } }),
        cell('삭센다 (Saxenda)', { bold: true, fill: { color: OB } }),
        cell('매일 1회 피하주사', { fill: { color: OB } }),
        cell('체중 감량 (비만)', { bold: true, color: ORANGE, fill: { color: OB } }),
      ],
      [
        cell('리라글루타이드', { bold: true, color: TEAL_DEEP }),
        cell('빅토자 (Victoza)', { bold: true }),
        cell('매일 1회 피하주사'),
        cell('2형 당뇨'),
      ],
      [
        cell('둘라글루타이드', { bold: true, color: TEAL_DEEP }),
        cell('트루리시티 (Trulicity)', { bold: true }),
        cell('주 1회 피하주사'),
        cell('2형 당뇨'),
      ],
    ];

    s.addTable(rows, {
      x: M, y: 1.74, w: W - 2 * M,
      colW: [2.35, 2.85, 2.55, 4.483],
      rowH: 0.46,
      border: { type: 'solid', color: BORDER, pt: 1 },
      autoPage: false,
    });

    // legend + caution
    const by = 5.90;
    card(s, pres, M, by, W - 2 * M, 0.86, { fill: AMBER_TINT, line: 'EBD9B4' });
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.24, y: by + 0.28, w: 0.30, h: 0.30,
      fill: { color: ORANGE_TINT }, line: { color: 'F0CDB6', width: 1 }, rectRadius: 0.06,
    });
    s.addText([
      { text: '주황색 줄 = 체중 감량 적응증으로 허가된 약입니다.  ', options: { bold: true, color: AMBER } },
      { text: '허가 적응증과 국내 공급 상황은 자주 바뀌므로, 처방 전 식품의약품안전처 허가정보를 확인하십시오.', options: { color: INK_SOFT } },
    ], {
      x: M + 0.68, y: by + 0.12, w: W - 2 * M - 0.95, h: 0.62, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12, lineSpacingMultiple: 1.3, align: 'left', valign: 'middle',
    });

    footnote(s, '미국에서는 터제파타이드의 체중 감량 적응증이 젭바운드(Zepbound)로 판매됩니다. 엑세나타이드·릭시세나타이드는 초기 세대 약제입니다. 표는 2026년 중반 기준입니다.', { y: 6.86 });
    s.addNotes('성분명과 상품명의 구분을 반드시 짚어 주십시오. 환자들은 "위고비"와 "오젬픽"을 다른 약으로 알고 오는 경우가 많지만 성분은 같은 세마글루타이드이며 허가 적응증과 용량이 다릅니다. 터제파타이드(마운자로)는 엄밀히 말해 순수 GLP-1 작용제가 아니라 GLP-1/GIP 이중 작용제입니다. 한미약품의 에페글레나타이드 등 국내 개발 약제도 상업화가 추진되고 있습니다.');
  }

  // ============================================================
  // SLIDE 5 — MAIN DIAGRAM: four actions
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '약은 네 곳에 동시에 말을 겁니다');

    const gap = 0.28;
    const cw = (W - 2 * M - gap * 3) / 4;
    const cy = 3.18, ch = 3.42;
    const centers = [0, 1, 2, 3].map(i => M + i * (cw + gap) + cw / 2);
    const hubCx = (centers[0] + centers[3]) / 2;

    // hub
    const hubW = 3.5, hubH = 0.88, hubY = 1.42;
    s.addShape(pres.ShapeType.roundRect, {
      x: hubCx - hubW / 2, y: hubY, w: hubW, h: hubH,
      fill: { color: ORANGE }, line: { type: 'none' }, rectRadius: 0.44,
      shadow: cardShadow(),
    });
    s.addText([{ text: 'GLP-1 작용제', options: {} }], {
      x: hubCx - hubW / 2, y: hubY, w: hubW, h: hubH, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 19, bold: true, color: WHITE, align: 'center', valign: 'middle',
    });

    // stem + rail + 4 arrows
    const railY = 2.76;
    s.addShape(pres.ShapeType.line, {
      x: hubCx, y: hubY + hubH, w: 0, h: railY - (hubY + hubH),
      line: { color: SLATE_LT, width: 1.75 },
    });
    s.addShape(pres.ShapeType.line, {
      x: centers[0], y: railY, w: centers[3] - centers[0], h: 0,
      line: { color: SLATE_LT, width: 1.75 },
    });
    centers.forEach(cx => {
      s.addShape(pres.ShapeType.line, {
        x: cx, y: railY, w: 0, h: cy - railY - 0.04,
        line: { color: SLATE_LT, width: 1.75, endArrowType: 'triangle' },
      });
    });

    const acts = [
      { n: '1', ic: I.brain,   org: '뇌',   what: '배고픔 신호를 줄입니다',   res: '덜 먹고, 더 빨리 배부름' },
      { n: '2', ic: I.stomach, org: '위',   what: '위가 천천히 비워집니다',   res: '포만감이 오래 지속됨' },
      { n: '3', ic: I.drop,    org: '췌장', what: '혈당이 높을 때만\n인슐린을 더 냅니다', res: '식후 혈당이 덜 오름' },
      { n: '4', ic: I.liver,   org: '간',   what: '간이 내보내는 당을\n줄입니다',        res: '공복 혈당이 내려감' },
    ];

    acts.forEach((a, i) => {
      const x = M + i * (cw + gap);
      card(s, pres, x, cy, cw, ch, { fill: WHITE });
      numDot(s, pres, x + 0.20, cy + 0.20, 0.40, a.n, { fill: ORANGE });
      iconBadge(s, pres, x + cw / 2, cy + 0.98, 0.98, a.ic, { bg: TEAL_TINT });
      s.addText([{ text: a.org, options: {} }], {
        x: x + 0.12, y: cy + 1.58, w: cw - 0.24, h: 0.42, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 20, bold: true, color: TEAL_DEEP, align: 'center', valign: 'middle',
      });
      s.addText([{ text: a.what, options: {} }], {
        x: x + 0.16, y: cy + 2.02, w: cw - 0.32, h: 0.72, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 12.5, color: SLATE, lineSpacingMultiple: 1.3,
        align: 'center', valign: 'top',
      });
      s.addShape(pres.ShapeType.roundRect, {
        x: x + 0.16, y: cy + 2.78, w: cw - 0.32, h: 0.48,
        fill: { color: ORANGE_TINT }, line: { color: ORANGE_TINT, width: 0 }, rectRadius: 0.07,
      });
      s.addText([{ text: a.res, options: {} }], {
        x: x + 0.16, y: cy + 2.78, w: cw - 0.32, h: 0.48, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 11, bold: true, color: ORANGE, align: 'center', valign: 'middle',
      });
    });

    footnote(s, '네 가지 작용은 동물 및 사람 연구에서 반복적으로 확인된 확립된 기전입니다. 출처: Drucker DJ. Cell Metab 2018; Müller TD 외. Mol Metab 2019.');
    s.addNotes('이 장이 발표의 중심입니다. 1번(뇌: 시상하부 궁상핵·후뇌 수용체)과 2번(위 배출 지연)이 체중 감소를, 3번(포도당 의존적 인슐린 분비 촉진)과 4번(글루카곤 억제 → 간 당 생성 감소)이 혈당 강하를 설명합니다. 체중 감소의 주된 동력은 식욕 억제이며, 위 배출 지연 효과는 시간이 지나며 약해질 수 있다는 보고가 있습니다(마지막 장의 불확실 항목).');
  }

  // ============================================================
  // SLIDE 6 — glucose-dependent switch
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '“혈당이 높을 때만” 켜지는 스위치');
    kicker(s, '인슐린을 무조건 더 내게 하는 약이 아니라, 혈당이 올라가 있을 때만 작용이 세지는 약입니다.');

    const pw = (W - 2 * M - 0.40) / 2;
    const py = 1.78, ph = 2.98;

    // ON panel
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: py, w: pw, h: ph,
      fill: { color: ORANGE_TINT }, line: { color: 'F0CDB6', width: 1 }, rectRadius: 0.09,
      shadow: cardShadow(),
    });
    s.addText([{ text: 'ON', options: {} }], {
      x: M + 0.38, y: py + 0.28, w: 1.2, h: 0.5, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 26, bold: true, color: ORANGE, align: 'left', valign: 'middle', charSpacing: 2,
    });
    s.addText([{ text: '식사 직후 — 혈당이 올라갈 때', options: {} }], {
      x: M + 0.38, y: py + 0.82, w: pw - 0.76, h: 0.4, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 16, bold: true, color: INK, align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '췌장 베타세포의 인슐린 분비를 강하게 돕습니다', options: { bullet: true, breakLine: true } },
      { text: '혈당을 올리는 글루카곤은 함께 억제합니다', options: { bullet: true, breakLine: true } },
      { text: '결과: 식후 혈당이 완만하게 올라갑니다', options: { bullet: true } },
    ], {
      x: M + 0.42, y: py + 1.34, w: pw - 0.84, h: 1.50, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12.5, color: INK_SOFT, paraSpaceAfter: 8, align: 'left', valign: 'top',
    });

    // OFF panel
    const rx = M + pw + 0.40;
    s.addShape(pres.ShapeType.roundRect, {
      x: rx, y: py, w: pw, h: ph,
      fill: { color: WHITE }, line: { color: BORDER, width: 1 }, rectRadius: 0.09,
      shadow: cardShadow(),
    });
    s.addText([{ text: 'OFF', options: {} }], {
      x: rx + 0.38, y: py + 0.28, w: 1.4, h: 0.5, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 26, bold: true, color: SLATE_LT, align: 'left', valign: 'middle', charSpacing: 2,
    });
    s.addText([{ text: '공복 — 혈당이 정상일 때', options: {} }], {
      x: rx + 0.38, y: py + 0.82, w: pw - 0.76, h: 0.4, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 16, bold: true, color: INK, align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '인슐린을 더 내라는 자극이 거의 사라집니다', options: { bullet: true, breakLine: true } },
      { text: '억제되던 글루카곤도 다시 제 역할을 합니다', options: { bullet: true, breakLine: true } },
      { text: '결과: 혈당을 정상 아래로 끌어내리지 않습니다', options: { bullet: true } },
    ], {
      x: rx + 0.42, y: py + 1.34, w: pw - 0.84, h: 1.50, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12.5, color: SLATE, paraSpaceAfter: 8, align: 'left', valign: 'top',
    });

    // caution
    const by = 5.22;
    card(s, pres, M, by, W - 2 * M, 1.32, { fill: AMBER_TINT, line: 'EBD9B4' });
    iconBadge(s, pres, M + 0.62, by + 0.66, 0.68, I.warn, { bg: WHITE, scale: 0.46 });
    s.addText([
      { text: '그래서 이 약만 쓸 때는 저혈당이 드뭅니다. 단, 예외가 있습니다.  ', options: { bold: true, color: AMBER } },
      { text: '설포닐우레아 계열 약이나 인슐린과 함께 쓰면 저혈당 위험이 올라가므로, 기존 약의 용량을 줄여야 할 수 있습니다. 반드시 처방의와 상의하십시오.', options: { color: INK_SOFT } },
    ], {
      x: M + 1.10, y: by + 0.22, w: W - 2 * M - 1.35, h: 0.90, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 13, lineSpacingMultiple: 1.35, align: 'left', valign: 'middle',
    });

    footnote(s, '“포도당 의존적(glucose-dependent)” 작용이라 부릅니다. 출처: Drucker DJ. Cell Metab 2018.');
    s.addNotes('포도당 의존성이 인크레틴 기반 치료의 가장 중요한 안전성 특징입니다. 설포닐우레아·인슐린 병용 시 저혈당 위험 증가는 약물 자체의 문제가 아니라 병용 약제 때문이며, 실제 진료에서 용량 감량이 필요합니다.');
  }

  // ============================================================
  // SLIDE 7 — efficacy (chart)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '효과: 임상시험에서 확인된 숫자');
    kicker(s, '모두 위약과 비교한 대규모 무작위배정 임상시험 결과입니다.');

    const chW = 7.55, chY = 1.78, chH = 4.42;
    card(s, pres, M, chY, chW, chH, { fill: WHITE });
    s.addText([{ text: '68~72주 후 평균 체중 변화 (%)', options: {} }], {
      x: M + 0.34, y: chY + 0.24, w: chW - 0.68, h: 0.36, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 14, bold: true, color: INK, align: 'left', valign: 'middle',
    });

    s.addChart(
      pres.ChartType.bar,
      [{
        name: '평균 체중 변화 (%)',
        labels: ['위약 · STEP 1', '세마글루타이드 2.4mg', '위약 · SURMOUNT-1', '터제파타이드 15mg'],
        values: [-2.4, -14.9, -3.1, -20.9],
      }],
      {
        x: M + 0.18, y: chY + 0.68, w: chW - 0.44, h: chH - 1.42,
        barDir: 'bar',
        chartColors: [GREY_BAR, TEAL, GREY_BAR, ORANGE],
        showLegend: false,
        showTitle: false,
        showValue: true,
        dataLabelPosition: 'outEnd',
        catAxisLabelPos: 'low',
        dataLabelColor: INK,
        dataLabelFontFace: K,
        dataLabelFontSize: 12,
        dataLabelFontBold: true,
        dataLabelFormatCode: '0.0"%"',
        catAxisLabelColor: INK_SOFT,
        catAxisLabelFontFace: K,
        catAxisLabelFontSize: 11.5,
        catAxisLineShow: false,
        catGridLine: { style: 'none' },
        valAxisLabelColor: SLATE_LT,
        valAxisLabelFontFace: K,
        valAxisLabelFontSize: 10,
        valAxisMaxVal: 0,
        valAxisMinVal: -26,
        valGridLine: { color: 'EDF2F5', size: 1 },
        valAxisLineShow: false,
        barGapWidthPct: 55,
      }
    );
    s.addText([{ text: '세마글루타이드 2.4mg = 위고비, 터제파타이드 15mg = 마운자로. 터제파타이드는 GLP-1과 GIP 두 수용체에 함께 작용합니다.', options: {} }], {
      x: M + 0.34, y: chY + chH - 0.66, w: chW - 0.68, h: 0.5, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 10.5, color: SLATE, lineSpacingMultiple: 1.25, align: 'left', valign: 'top',
    });

    // right column
    const rx = M + chW + 0.34;
    const rw = W - M - rx;

    s.addShape(pres.ShapeType.roundRect, {
      x: rx, y: chY, w: rw, h: 2.34,
      fill: { color: TEAL }, line: { color: TEAL, width: 0 }, rectRadius: 0.09, shadow: cardShadow(),
    });
    iconBadge(s, pres, rx + 0.62, chY + 0.60, 0.66, I.heart, { bg: TEAL_DEEP, scale: 0.5 });
    s.addText([
      { text: '20%', options: { fontSize: 44, bold: true, color: WHITE } },
      { text: '  감소', options: { fontSize: 17, bold: true, color: 'BCDCE9' } },
    ], {
      x: rx + 1.18, y: chY + 0.26, w: rw - 1.5, h: 0.72, isTextBox: true, margin: 0,
      fontFace: K, align: 'left', valign: 'middle',
    });
    s.addText([{ text: '위약보다 주요 심혈관 사건(심혈관 사망·심근경색·뇌졸중)이 20% 적었습니다.\n대상: 심혈관질환이 있고 당뇨는 없는 과체중·비만 성인 (SELECT, 위험비 0.80).', options: {} }], {
      x: rx + 0.36, y: chY + 1.00, w: rw - 0.72, h: 1.20, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 11, color: 'D6EAF2', lineSpacingMultiple: 1.3, align: 'left', valign: 'top',
    });

    const ry2 = chY + 2.54;
    card(s, pres, rx, ry2, rw, chH - 2.54, { fill: AMBER_TINT, line: 'EBD9B4' });
    s.addText([{ text: '숫자를 읽을 때 주의', options: {} }], {
      x: rx + 0.36, y: ry2 + 0.26, w: rw - 0.72, h: 0.36, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 14, bold: true, color: AMBER, align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '서로 다른 시험의 결과이며 직접 비교가 아닙니다', options: { bullet: true, breakLine: true } },
      { text: '식사·운동 상담이 함께 제공되었습니다', options: { bullet: true, breakLine: true } },
      { text: '평균값이며 개인차가 큽니다', options: { bullet: true } },
    ], {
      x: rx + 0.40, y: ry2 + 0.66, w: rw - 0.80, h: 1.12, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 10.5, color: INK_SOFT, paraSpaceAfter: 6, align: 'left', valign: 'top',
    });

    footnote(s, '출처: Wilding JPH 외. N Engl J Med 2021;384:989 (STEP 1) · Jastreboff AM 외. N Engl J Med 2022;387:205 (SURMOUNT-1) · Lincoff AM 외. N Engl J Med 2023;389:2221 (SELECT).');
    s.addNotes('체중 변화는 STEP 1(세마글루타이드 2.4mg, 68주, -14.9% vs 위약 -2.4%)과 SURMOUNT-1(터제파타이드 15mg, 72주, -20.9% vs 위약 -3.1%)입니다. SELECT는 당뇨가 없는 과체중·비만 + 심혈관질환 환자에서 MACE를 20% 감소시켰습니다(HR 0.80, 95% CI 0.72-0.90). 시험 간 직접 비교가 아니라는 점을 반드시 언급하십시오.');
  }

  // ============================================================
  // SLIDE 8 — adverse effects
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '부작용과 알아둘 점');
    kicker(s, '대부분은 앞서 본 작용 기전에서 그대로 따라 나옵니다 — 위가 천천히 비워지니 속이 불편한 것입니다.');

    const rows = [
      { ic: I.stomach, tag: '가장 흔함', t: '구역 · 구토 · 설사 · 변비',
        b: '세마글루타이드 2.4mg 시험에서 구역 약 44%(위약 17%). 주로 초기와 증량 시기에 생기고 시간이 지나며 줄어듭니다. 용량을 천천히 올리는 것이 표준 대처법입니다.' },
      { ic: I.glass, tag: '드물지만 증가', t: '담석 · 담낭 질환',
        b: '담낭 관련 이상반응 2.6% 대 위약 1.2%. 빠른 체중 감소 자체와도 관련이 있습니다.' },
      { ic: I.syringe, tag: '드묾', t: '췌장염',
        b: '보고되어 있으나 드물며, 대규모 심혈관 결과 시험에서 위약보다 뚜렷이 증가한다는 근거는 확인되지 않았습니다.' },
      { ic: I.muscle, tag: '관리 필요', t: '근육량 감소',
        b: '빠진 체중의 일부는 지방이 아닌 제지방(근육 포함)입니다. 충분한 단백질 섭취와 근력 운동이 권고됩니다.' },
      { ic: I.weight, tag: '중요', t: '끊으면 체중이 돌아옵니다',
        b: 'STEP 1 연장 연구에서 중단 1년 후 감량분의 약 3분의 2가 되돌아왔습니다. 비만은 장기 관리가 필요한 상태로 봅니다.' },
    ];

    const ry = 1.72, rh = 0.92, rgap = 0.05;
    rows.forEach((r, i) => {
      const y = ry + i * (rh + rgap);
      card(s, pres, M, y, W - 2 * M, rh, { fill: WHITE, shadow: false });
      iconBadge(s, pres, M + 0.60, y + rh / 2, 0.66, r.ic, { bg: TEAL_TINT, scale: 0.5 });
      s.addShape(pres.ShapeType.roundRect, {
        x: M + 1.12, y: y + 0.17, w: 1.16, h: 0.30,
        fill: { color: ORANGE_TINT }, line: { color: ORANGE_TINT, width: 0 }, rectRadius: 0.07,
      });
      s.addText([{ text: r.tag, options: {} }], {
        x: M + 1.12, y: y + 0.17, w: 1.16, h: 0.30, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 9.5, bold: true, color: ORANGE, align: 'center', valign: 'middle',
      });
      s.addText([{ text: r.t, options: {} }], {
        x: M + 2.40, y: y + 0.14, w: 3.05, h: 0.36, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 14.5, bold: true, color: TEAL_DEEP, align: 'left', valign: 'middle',
      });
      s.addText([{ text: r.b, options: {} }], {
        x: M + 2.40, y: y + 0.46, w: W - 2 * M - 2.70, h: 0.42, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 10.5, color: SLATE, lineSpacingMultiple: 1.18, align: 'left', valign: 'top',
      });
    });

    footnote(s, '갑상선 수질암 병력·가족력이 있으면 사용하지 않습니다(설치류 자료에 근거한 경고이며 사람에서의 인과관계는 미확립 — 다음 장). 출처: Wilding JPH 외. NEJM 2021; Diabetes Obes Metab 2022.', { y: 6.62 });
    s.addNotes('구역 등 위장관 증상은 위 배출 지연의 직접적 결과로 설명하면 청중이 잘 이해합니다. 서서히 증량(dose escalation)이 핵심 대처입니다. 중단 후 체중 재증가는 STEP 1 연장 연구(Diabetes Obes Metab 2022)에서 68주 -17.3% → 120주 -5.6%로 보고되었습니다.');
  }

  // ============================================================
  // SLIDE 9 — established vs uncertain
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '확실한 것과, 아직 모르는 것');
    kicker(s, '근거의 강도를 구분해서 보는 것이 이 주제를 이해하는 가장 정직한 방법입니다.');

    const pw = (W - 2 * M - 0.40) / 2;
    const py = 1.78, ph = 4.70;

    // established
    card(s, pres, M, py, pw, ph, { fill: GREEN_TINT, line: 'C4E0D2' });
    iconBadge(s, pres, M + 0.62, py + 0.60, 0.66, I.check, { bg: WHITE, scale: 0.46 });
    s.addText([{ text: '근거가 확립된 것', options: {} }], {
      x: M + 1.14, y: py + 0.34, w: pw - 1.5, h: 0.52, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 19, bold: true, color: GREEN, align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '앞서 본 네 가지 작용 경로 — 동물과 사람 연구에서 반복 확인되었습니다', options: { bullet: true, breakLine: true } },
      { text: '위약보다 뚜렷한 체중 감소와 혈당 개선 — 다수의 대규모 무작위배정 시험', options: { bullet: true, breakLine: true } },
      { text: '심혈관 고위험군에서 주요 심혈관 사건 감소 — SELECT, LEADER 등', options: { bullet: true, breakLine: true } },
      { text: '위장관 증상이 가장 흔한 부작용이며 대개 용량 증량 시기에 집중됩니다', options: { bullet: true, breakLine: true } },
      { text: '단독 사용 시 저혈당 위험이 낮습니다 — 포도당 의존적 작용 때문입니다', options: { bullet: true } },
    ], {
      x: M + 0.42, y: py + 1.16, w: pw - 0.84, h: ph - 1.50, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12.5, color: INK_SOFT, paraSpaceAfter: 12,
      lineSpacingMultiple: 1.3, align: 'left', valign: 'top',
    });

    // uncertain
    const rx = M + pw + 0.40;
    card(s, pres, rx, py, pw, ph, { fill: AMBER_TINT, line: 'EBD9B4' });
    iconBadge(s, pres, rx + 0.62, py + 0.60, 0.66, I.quest, { bg: WHITE, scale: 0.42 });
    s.addText([{ text: '아직 불확실한 것', options: {} }], {
      x: rx + 1.14, y: py + 0.34, w: pw - 1.5, h: 0.52, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 19, bold: true, color: AMBER, align: 'left', valign: 'middle',
    });
    s.addText([
      { text: '5~10년을 넘는 장기 안전성 자료는 아직 쌓이는 중입니다', options: { bullet: true, breakLine: true } },
      { text: '설치류에서 관찰된 갑상선 C세포 종양이 사람에게도 의미가 있는지는 결론나지 않았습니다', options: { bullet: true, breakLine: true } },
      { text: '근육량 감소가 장기 건강에 어떤 영향을 주는지, 어느 정도까지 막을 수 있는지', options: { bullet: true, breakLine: true } },
      { text: '장기 체중 감소에서 “뇌의 식욕 억제”와 “위 배출 지연” 중 어느 쪽이 더 중요한지 — 위 배출 지연 효과는 시간이 지나며 약해진다는 보고가 있습니다', options: { bullet: true, breakLine: true } },
      { text: '약을 평생 써야 하는지, 언제 어떻게 줄일 수 있는지에 대한 근거는 부족합니다', options: { bullet: true } },
    ], {
      x: rx + 0.42, y: py + 1.16, w: pw - 0.84, h: ph - 1.50, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12.5, color: INK_SOFT, paraSpaceAfter: 12,
      lineSpacingMultiple: 1.3, align: 'left', valign: 'top',
    });

    footnote(s, '“불확실”은 위험하다는 뜻이 아니라 아직 답이 정해지지 않았다는 뜻입니다.');
    s.addNotes('이 장은 청중이 언론 보도와 실제 근거를 구분하도록 돕습니다. 갑상선 C세포 종양은 설치류 데이터에 근거한 경고이며 사람에서의 인과관계는 확립되지 않았습니다. 다만 갑상선 수질암 병력·가족력, 다발성 내분비선종증 2형은 금기입니다.');
  }

  // ============================================================
  // SLIDE 10 — references
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: PAPER };
    slideTitle(s, '근거 자료');
    kicker(s, '이 발표의 수치와 기전 설명은 아래 문헌에서 인용했습니다.');

    const refs = [
      { tag: '기전', t: 'Drucker DJ. Mechanisms of Action and Therapeutic Application of Glucagon-like Peptide-1. Cell Metab. 2018;27(4):740-756.' },
      { tag: '기전', t: 'Müller TD, Finan B, Bloom SR, 외. Glucagon-like peptide 1 (GLP-1). Mol Metab. 2019;30:72-130.' },
      { tag: '체중', t: 'Wilding JPH, Batterham RL, Calanna S, 외. Once-Weekly Semaglutide in Adults with Overweight or Obesity (STEP 1). N Engl J Med. 2021;384(11):989-1002.' },
      { tag: '중단', t: 'Wilding JPH, Batterham RL, Davies M, 외. Weight regain and cardiometabolic effects after withdrawal of semaglutide: the STEP 1 trial extension. Diabetes Obes Metab. 2022;24(8):1553-1564.' },
      { tag: '체중', t: 'Jastreboff AM, Aronne LJ, Ahmad NN, 외. Tirzepatide Once Weekly for the Treatment of Obesity (SURMOUNT-1). N Engl J Med. 2022;387(3):205-216.' },
      { tag: '심혈관', t: 'Lincoff AM, Brown-Frandsen K, Colhoun HM, 외. Semaglutide and Cardiovascular Outcomes in Obesity without Diabetes (SELECT). N Engl J Med. 2023;389(24):2221-2232.' },
    ];

    const ry = 1.72, rh = 0.73, rgap = 0.055;
    refs.forEach((r, i) => {
      const y = ry + i * (rh + rgap);
      card(s, pres, M, y, W - 2 * M, rh, { fill: WHITE, shadow: false });
      s.addShape(pres.ShapeType.roundRect, {
        x: M + 0.24, y: y + 0.21, w: 0.92, h: 0.31,
        fill: { color: TEAL_TINT }, line: { color: TEAL_TINT, width: 0 }, rectRadius: 0.07,
      });
      s.addText([{ text: r.tag, options: {} }], {
        x: M + 0.24, y: y + 0.21, w: 0.92, h: 0.31, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 9.5, bold: true, color: TEAL_DEEP, align: 'center', valign: 'middle',
      });
      s.addText([{ text: r.t, options: {} }], {
        x: M + 1.30, y: y + 0.10, w: W - 2 * M - 1.60, h: 0.54, isTextBox: true, margin: 0,
        fontFace: K, fontSize: 11.5, color: INK_SOFT, lineSpacingMultiple: 1.25,
        align: 'left', valign: 'middle',
      });
    });

    footnote(s, '약제별 허가사항과 용량은 국내 식품의약품안전처 허가정보 및 최신 진료지침을 확인하십시오.', { y: 6.72 });
    s.addNotes('필요하면 국내 지침(대한비만학회 비만 진료지침, 대한당뇨병학회 진료지침)을 추가로 인용하십시오.');
  }

  // ============================================================
  // SLIDE 11 — closing (dark)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    s.addShape(pres.ShapeType.ellipse, {
      x: -2.1, y: 4.05, w: 5.0, h: 5.0,
      fill: { color: ORANGE, transparency: 88 }, line: { type: 'none' },
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: 0.5, y: 5.4, w: 0.62, h: 0.62,
      fill: { color: ORANGE }, line: { type: 'none' },
    });

    s.addText([{ text: '한 문장으로', options: {} }], {
      x: 1.6, y: 1.52, w: 10.2, h: 0.4, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 14.5, bold: true, color: 'E9A87C', charSpacing: 1.5,
      align: 'left', valign: 'middle',
    });
    s.addText([
      { text: 'GLP-1 작용제는 ‘살을 빼는 약’이라기보다,', options: { breakLine: true } },
      { text: '식사 후 몸이 스스로 보내는 포만·혈당 신호를', options: { breakLine: true } },
      { text: '며칠 동안 켜 두는 약입니다.', options: {} },
    ], {
      x: 1.6, y: 2.05, w: 10.4, h: 2.5, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 29, bold: true, color: WHITE, lineSpacingMultiple: 1.35,
      align: 'left', valign: 'top',
    });
    s.addText([{ text: '그래서 효과도, 부작용도, 끊었을 때 되돌아오는 것도\n모두 같은 기전에서 따라 나옵니다.', options: {} }], {
      x: 1.6, y: 4.62, w: 9.6, h: 0.9, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 15, color: 'AFC6D2', lineSpacingMultiple: 1.45,
      align: 'left', valign: 'top',
    });

    s.addShape(pres.ShapeType.roundRect, {
      x: 1.6, y: 5.85, w: 10.2, h: 0.78,
      fill: { color: INK_SOFT }, line: { color: '2E5468', width: 1 }, rectRadius: 0.09,
    });
    s.addText([{ text: '이 자료는 교육용입니다. 실제 사용 여부, 약제 선택, 용량은 개인의 상태에 따라 다르므로 반드시 진료 상담이 필요합니다.', options: {} }], {
      x: 1.85, y: 5.85, w: 9.7, h: 0.78, isTextBox: true, margin: 0,
      fontFace: K, fontSize: 12, color: 'AFC6D2', align: 'left', valign: 'middle',
    });
    s.addNotes('마무리: 기전 → 효과 → 부작용 → 중단 후 재증가가 하나의 논리로 이어진다는 점을 다시 짚어 주십시오. 질의응답에서 가장 많이 나오는 질문은 (1) 평생 써야 하나요 (2) 근육이 빠지지 않나요 (3) 보험이 되나요 입니다.');
  }

  const out = process.argv[2] || 'GLP-1_작용기전_일반인설명.pptx';
  await pres.writeFile({ fileName: out });
  console.log('wrote', out);
}

build().catch(e => { console.error(e); process.exit(1); });
