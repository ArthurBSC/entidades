/**
 * NURA Pitch Deck Generator — v3 (ícones reais + bordas arredondadas)
 * 10 slides · Brasília TecnoGame 2026
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaClipboardList, FaDollarSign, FaBook, FaBullseye, FaGraduationCap,
  FaClipboardCheck, FaComments, FaFingerprint, FaRobot,
  FaRegClock, FaHourglassHalf, FaEnvelopeOpen,
  FaCode, FaBrain, FaLayerGroup, FaMobileAlt,
  FaChartLine, FaSyncAlt, FaCoins,
  FaUserTie, FaLaptopCode, FaCogs,
  FaArrowRight,
} = require("react-icons/fa");

const ASSETS = path.join(__dirname, "assets");

// ===========================================================
// PALETA VISUAL — Midnight Neon
// ===========================================================
const C = {
  bgDark: "0A0A1F",
  bgMid: "131333",
  bgLight: "1E1E47",
  primary: "7C3AED",
  primaryLight: "A78BFA",
  accent: "3B82F6",
  accentWarm: "F59E0B",
  danger: "EF4444",
  success: "10B981",
  textPrimary: "FFFFFF",
  textSecondary: "CBD5E1",
  textMuted: "64748B",
  border: "334155",
};

const FONT_TITLE = "Calibri";
const FONT_BODY = "Calibri";

// ===========================================================
// ICON HELPERS (react-icons → PNG base64)
// ===========================================================
const iconCache = {};

async function iconPng(IconComp, hexColor = "FFFFFF", size = 256) {
  const key = `${IconComp.name}-${hexColor}-${size}`;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color: `#${hexColor}`, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  const data = "image/png;base64," + buf.toString("base64");
  iconCache[key] = data;
  return data;
}

// ===========================================================
// INSTÂNCIA
// ===========================================================
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "NURA";
pres.title = "NURA · Pitch Deck";
pres.subject = "O sistema operacional da educação brasileira";

const W = 13.3;
const H = 7.5;
const R = 0.1; // raio padrão de cantos arredondados

// ===========================================================
// HELPERS de layout
// ===========================================================
function darkBg(slide) {
  slide.background = { color: C.bgDark };
}

function addFooter(slide, numero) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.02,
    fill: { color: C.border }, line: { color: C.border },
  });
  slide.addText("NURA · Pitch Deck · Brasília TecnoGame 2026", {
    x: 0.5, y: H - 0.3, w: 8, h: 0.25,
    fontSize: 9, fontFace: FONT_BODY, color: C.textMuted,
    align: "left", valign: "middle", margin: 0,
  });
  slide.addText(`${numero} / 10`, {
    x: W - 1.5, y: H - 0.3, w: 1, h: 0.25,
    fontSize: 9, fontFace: FONT_BODY, color: C.textMuted,
    align: "right", valign: "middle", margin: 0,
  });
}

function sectionBadge(slide, label, x, y) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 0.15, h: 0.35,
    fill: { color: C.primary }, line: { color: C.primary }, rectRadius: 0.04,
  });
  slide.addText(label.toUpperCase(), {
    x: x + 0.25, y, w: 6, h: 0.35,
    fontSize: 11, fontFace: FONT_BODY, color: C.primaryLight,
    bold: true, charSpacing: 6, valign: "middle", margin: 0,
  });
}

// ===========================================================
// BUILD (async for icon generation)
// ===========================================================
async function build() {
  // Preload de ícones — base64 PNGs universais
  const I = {
    secretaria: await iconPng(FaClipboardList, "A78BFA"),
    financeiro: await iconPng(FaDollarSign, "3B82F6"),
    biblioteca: await iconPng(FaBook, "10B981"),
    processo: await iconPng(FaBullseye, "F59E0B"),
    estagio: await iconPng(FaGraduationCap, "A78BFA"),
    avaliacao: await iconPng(FaClipboardCheck, "3B82F6"),
    comunicacao: await iconPng(FaComments, "10B981"),
    portaria: await iconPng(FaFingerprint, "F59E0B"),
    ia: await iconPng(FaRobot, "FFFFFF"),

    clock: await iconPng(FaRegClock, "F59E0B"),
    hourglass: await iconPng(FaHourglassHalf, "F59E0B"),
    envelope: await iconPng(FaEnvelopeOpen, "F59E0B"),

    code: await iconPng(FaCode, "A78BFA"),
    brain: await iconPng(FaBrain, "A78BFA"),
    layers: await iconPng(FaLayerGroup, "A78BFA"),
    mobile: await iconPng(FaMobileAlt, "A78BFA"),

    chart: await iconPng(FaChartLine, "A78BFA"),
    cycle: await iconPng(FaSyncAlt, "A78BFA"),
    coins: await iconPng(FaCoins, "A78BFA"),

    pm: await iconPng(FaUserTie, "FFFFFF"),
    dev: await iconPng(FaLaptopCode, "FFFFFF"),
    stack: await iconPng(FaCogs, "FFFFFF"),

    arrow: await iconPng(FaArrowRight, "FFFFFF"),
  };

  // =========================================================
  // SLIDE 1 · CAPA
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);

    s.addShape(pres.shapes.OVAL, {
      x: -2, y: -2, w: 6, h: 6,
      fill: { color: C.primary, transparency: 85 }, line: { color: C.primary, transparency: 100 },
    });
    s.addShape(pres.shapes.OVAL, {
      x: W - 4, y: H - 4, w: 6, h: 6,
      fill: { color: C.accent, transparency: 85 }, line: { color: C.accent, transparency: 100 },
    });

    s.addImage({
      path: path.join(ASSETS, "nura-logo.png"),
      x: (W - 3.5) / 2, y: 2.0, w: 3.5, h: 1.0,
      sizing: { type: "contain", w: 3.5, h: 1.0 },
    });

    s.addText("NURA", {
      x: 0, y: 3.2, w: W, h: 1.2,
      fontSize: 96, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "center", valign: "middle", charSpacing: 12,
    });

    s.addText("O sistema operacional da educação brasileira", {
      x: 0, y: 4.5, w: W, h: 0.6,
      fontSize: 22, fontFace: FONT_BODY, color: C.primaryLight,
      align: "center", valign: "middle", italic: true,
    });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: (W - 4) / 2, y: 5.8, w: 4, h: 0.5,
      fill: { color: C.bgMid }, line: { color: C.border, width: 1 }, rectRadius: 0.08,
    });
    s.addText("BRASÍLIA · TECNOGAME 2026", {
      x: (W - 4) / 2, y: 5.8, w: 4, h: 0.5,
      fontSize: 12, fontFace: FONT_BODY, color: C.textSecondary,
      bold: true, align: "center", valign: "middle", charSpacing: 8, margin: 0,
    });
  }

  // =========================================================
  // SLIDE 2 · A DOR
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "1 · A Dor", 0.6, 0.5);

    s.addText("O ensino superior brasileiro opera no caos.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 1.0,
      fontSize: 40, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });

    s.addText("Três cenas que se repetem todos os dias em 2.400 instituições:", {
      x: 0.6, y: 1.9, w: W - 1.2, h: 0.5,
      fontSize: 16, fontFace: FONT_BODY, color: C.textSecondary,
      align: "left", valign: "top", margin: 0,
    });

    const cards = [
      { icon: I.clock, label: "22h", title: "Retrabalho manual",
        desc: "Secretária re-digita a mesma nota em 3 sistemas diferentes. Toda. Semana." },
      { icon: I.hourglass, label: "40%", title: "Tempo desperdiçado",
        desc: "Do tempo do servidor acadêmico é gasto em tarefas que um software deveria fazer." },
      { icon: I.envelope, label: "2 SEM", title: "Alunos sumindo",
        desc: "Estudantes desistem porque ninguém respondeu o e-mail há duas semanas." },
    ];

    const cardW = 3.8, cardH = 3.3, gap = 0.3;
    const startX = (W - (3 * cardW + 2 * gap)) / 2;
    cards.forEach((c, i) => {
      const x = startX + i * (cardW + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.8, w: cardW, h: cardH,
        fill: { color: C.bgMid }, line: { color: C.border, width: 1 }, rectRadius: 0.15,
      });
      // Barra lateral vermelha com topo arredondado (apenas lateral esquerda)
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.8, w: 0.1, h: cardH,
        fill: { color: C.danger }, line: { color: C.danger }, rectRadius: 0.05,
      });
      // Ícone
      s.addImage({
        data: c.icon, x: x + 0.4, y: 3.0, w: 0.7, h: 0.7,
      });
      // Número gigante
      s.addText(c.label, {
        x: x + 0.4, y: 3.8, w: cardW - 0.8, h: 0.9,
        fontSize: 44, fontFace: FONT_TITLE, color: C.accentWarm,
        bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addText(c.title, {
        x: x + 0.4, y: 4.8, w: cardW - 0.8, h: 0.5,
        fontSize: 18, fontFace: FONT_TITLE, color: C.textPrimary,
        bold: true, align: "left", valign: "top", margin: 0,
      });
      s.addText(c.desc, {
        x: x + 0.4, y: 5.3, w: cardW - 0.8, h: 0.9,
        fontSize: 13, fontFace: FONT_BODY, color: C.textSecondary,
        align: "left", valign: "top", margin: 0,
      });
    });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y: 6.4, w: W - 1.2, h: 0.55,
      fill: { color: C.bgLight }, line: { color: C.primary, width: 1 }, rectRadius: 0.1,
    });
    s.addText("Essa dor nós sentimos na pele — como alunos de IES privada.", {
      x: 0.6, y: 6.4, w: W - 1.2, h: 0.55,
      fontSize: 14, fontFace: FONT_BODY, color: C.primaryLight,
      italic: true, align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, 2);
  }

  // =========================================================
  // SLIDE 3 · ALTERNATIVAS
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "2 · Alternativas", 0.6, 0.5);

    s.addText("As duas saídas atuais são sintoma — não solução.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 1.0,
      fontSize: 36, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });

    const colW = 5.8, colH = 3.9, gap = 0.4;
    const startX = (W - (2 * colW + gap)) / 2;

    // Coluna A
    const colA_x = startX;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: colA_x, y: 2.3, w: colW, h: colH,
      fill: { color: C.bgMid }, line: { color: C.danger, width: 2 }, rectRadius: 0.15,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: colA_x, y: 2.3, w: colW, h: 0.7,
      fill: { color: C.danger }, line: { color: C.danger }, rectRadius: 0.15,
    });
    // Cobrindo a parte inferior para o corner superior arredondado se mantém
    s.addShape(pres.shapes.RECTANGLE, {
      x: colA_x, y: 2.6, w: colW, h: 0.4,
      fill: { color: C.danger }, line: { color: C.danger },
    });
    s.addText("OPÇÃO A · FRANKENSTEIN", {
      x: colA_x, y: 2.3, w: colW, h: 0.7,
      fontSize: 16, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "center", valign: "middle", charSpacing: 4, margin: 0,
    });
    s.addText([
      { text: "8–10 fornecedores diferentes", options: { bullet: true, breakLine: true } },
      { text: "Sponte + Pergamum + Zenvia + TOTVS…", options: { bullet: true, breakLine: true } },
      { text: "Custo de integração > custo de licença", options: { bullet: true, breakLine: true } },
      { text: "Dados duplicados em 5 bancos", options: { bullet: true, breakLine: true } },
      { text: "Fragmentação total do processo", options: { bullet: true } },
    ], {
      x: colA_x + 0.4, y: 3.2, w: colW - 0.8, h: colH - 1.0,
      fontSize: 15, fontFace: FONT_BODY, color: C.textSecondary,
      paraSpaceAfter: 6,
    });

    // Coluna B
    const colB_x = startX + colW + gap;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: colB_x, y: 2.3, w: colW, h: colH,
      fill: { color: C.bgMid }, line: { color: C.accentWarm, width: 2 }, rectRadius: 0.15,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: colB_x, y: 2.3, w: colW, h: 0.7,
      fill: { color: C.accentWarm }, line: { color: C.accentWarm }, rectRadius: 0.15,
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: colB_x, y: 2.6, w: colW, h: 0.4,
      fill: { color: C.accentWarm }, line: { color: C.accentWarm },
    });
    s.addText("OPÇÃO B · PLANILHA COMPARTILHADA", {
      x: colB_x, y: 2.3, w: colW, h: 0.7,
      fontSize: 16, fontFace: FONT_TITLE, color: C.bgDark,
      bold: true, align: "center", valign: "middle", charSpacing: 4, margin: 0,
    });
    s.addText([
      { text: "Excel + Google Forms + WhatsApp", options: { bullet: true, breakLine: true } },
      { text: "Zero automação, zero auditoria", options: { bullet: true, breakLine: true } },
      { text: "Dados somem quando alguém sai", options: { bullet: true, breakLine: true } },
      { text: "Sem relatório para MEC/INEP", options: { bullet: true, breakLine: true } },
      { text: "Informalidade total", options: { bullet: true } },
    ], {
      x: colB_x + 0.4, y: 3.2, w: colW - 0.8, h: colH - 1.0,
      fontSize: 15, fontFace: FONT_BODY, color: C.textSecondary,
      paraSpaceAfter: 6,
    });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y: 6.4, w: W - 1.2, h: 0.55,
      fill: { color: C.bgLight }, line: { color: C.primary, width: 1 }, rectRadius: 0.1,
    });
    s.addText("O Brasil nunca teve uma plataforma nativa e unificada para ensino superior.", {
      x: 0.6, y: 6.4, w: W - 1.2, h: 0.55,
      fontSize: 13, fontFace: FONT_BODY, color: C.primaryLight,
      bold: true, italic: true, align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, 3);
  }

  // =========================================================
  // SLIDE 4 · A SOLUÇÃO — Grid de módulos com ícones reais
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "3 · A Solução", 0.6, 0.5);

    s.addText("NURA é essa plataforma.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
      fontSize: 44, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });
    s.addText("Um login. Uma fatura. Uma fonte de verdade.", {
      x: 0.6, y: 1.9, w: W - 1.2, h: 0.5,
      fontSize: 18, fontFace: FONT_BODY, color: C.primaryLight,
      italic: true, align: "left", valign: "top", margin: 0,
    });

    const modules = [
      { name: "Secretaria", icon: I.secretaria, color: C.primaryLight },
      { name: "Financeiro", icon: I.financeiro, color: C.accent },
      { name: "Biblioteca", icon: I.biblioteca, color: C.success },
      { name: "Processo Seletivo", icon: I.processo, color: C.accentWarm },
      { name: "Estágio & TCC", icon: I.estagio, color: C.primaryLight },
      { name: "Avaliação Institucional", icon: I.avaliacao, color: C.accent },
      { name: "Comunicação", icon: I.comunicacao, color: C.success },
      { name: "Portaria Biométrica", icon: I.portaria, color: C.accentWarm },
      { name: "Copiloto de IA", icon: I.ia, color: C.textPrimary, highlight: true },
    ];

    const mW = 3.4, mH = 1.3, mGap = 0.2;
    const gridX = (W - (3 * mW + 2 * mGap)) / 2;
    const gridY = 2.7;

    modules.forEach((m, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = gridX + col * (mW + mGap);
      const y = gridY + row * (mH + mGap);
      const isAI = m.highlight === true;

      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: mW, h: mH,
        fill: { color: isAI ? C.primary : C.bgMid },
        line: { color: isAI ? C.primaryLight : C.border, width: 1 },
        rectRadius: 0.12,
      });
      // Ícone em círculo
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.25, y: y + (mH - 0.75) / 2, w: 0.75, h: 0.75,
        fill: { color: isAI ? C.textPrimary + "20" : C.bgLight },
        line: { color: isAI ? C.textPrimary : m.color, width: 1 },
      });
      s.addImage({
        data: m.icon,
        x: x + 0.425, y: y + (mH - 0.4) / 2, w: 0.4, h: 0.4,
      });
      s.addText(m.name, {
        x: x + 1.15, y, w: mW - 1.3, h: mH,
        fontSize: 15, fontFace: FONT_TITLE,
        color: isAI ? C.textPrimary : C.textSecondary,
        bold: true, align: "left", valign: "middle", margin: 0,
      });
    });

    addFooter(s, 4);
  }

  // =========================================================
  // SLIDE 5 · DIFERENCIAL (SEM a imagem ai-visual)
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "3 · O Diferencial", 0.6, 0.5);

    s.addText("Construído em 2026. De verdade.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
      fontSize: 40, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });
    s.addText("Enquanto concorrentes modernizam 20 anos de legado, nós começamos do zero com IA no núcleo.", {
      x: 0.6, y: 1.85, w: W - 1.2, h: 0.6,
      fontSize: 16, fontFace: FONT_BODY, color: C.textSecondary,
      align: "left", valign: "top", margin: 0,
    });

    // Grid 2x2 de diferenciais ocupando o espaço todo
    const diffs = [
      { icon: I.layers, title: "Stack nativa 2026",
        desc: "Zero dívida técnica. Infraestrutura em nuvem moderna e escalável desde o dia zero." },
      { icon: I.brain, title: "IA desde o primeiro commit",
        desc: "Copiloto Lyra no núcleo — não é feature colada depois, é arquitetura fundamental." },
      { icon: I.code, title: "Multi-tenant de verdade",
        desc: "Uma instância, milhares de instituições isoladas por slug com zero cross-contamination." },
      { icon: I.mobile, title: "UX mobile-first",
        desc: "Gestores e alunos usam do celular sem perder função. Responsivo de fato, não no slide." },
    ];

    const dW = 5.8, dH = 1.85, dGap = 0.35;
    const dStartX = (W - (2 * dW + dGap)) / 2;
    const dStartY = 2.9;

    diffs.forEach((d, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = dStartX + col * (dW + dGap);
      const y = dStartY + row * (dH + dGap);

      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: dW, h: dH,
        fill: { color: C.bgMid }, line: { color: C.border, width: 1 }, rectRadius: 0.15,
      });
      // Ícone circular
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.35, y: y + 0.4, w: 1.0, h: 1.0,
        fill: { color: C.primary }, line: { color: C.primaryLight, width: 2 },
      });
      s.addImage({
        data: d.icon,
        x: x + 0.6, y: y + 0.65, w: 0.5, h: 0.5,
      });
      // Título
      s.addText(d.title, {
        x: x + 1.55, y: y + 0.3, w: dW - 1.7, h: 0.5,
        fontSize: 20, fontFace: FONT_TITLE, color: C.textPrimary,
        bold: true, align: "left", valign: "middle", margin: 0,
      });
      // Desc
      s.addText(d.desc, {
        x: x + 1.55, y: y + 0.8, w: dW - 1.7, h: 0.95,
        fontSize: 13, fontFace: FONT_BODY, color: C.textSecondary,
        align: "left", valign: "top", margin: 0,
      });
    });

    addFooter(s, 5);
  }

  // =========================================================
  // SLIDE 6 · MODELO DE NEGÓCIO
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "4 · Modelo de Negócio", 0.6, 0.5);

    s.addText("SaaS B2B · R$ 80 por aluno ativo / mês", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
      fontSize: 36, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });
    s.addText("Receita recorrente, previsível e escalável — mensal, por aluno ativo.", {
      x: 0.6, y: 1.85, w: W - 1.2, h: 0.5,
      fontSize: 15, fontFace: FONT_BODY, color: C.textSecondary,
      align: "left", valign: "top", margin: 0,
    });

    // Calculadora
    const calcX = 1.0, calcY = 2.9, calcW = W - 2.0, calcH = 2.4;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: calcX, y: calcY, w: calcW, h: calcH,
      fill: { color: C.bgMid }, line: { color: C.primary, width: 2 }, rectRadius: 0.2,
    });
    s.addText("UMA INSTITUIÇÃO MÉDIA GERA", {
      x: calcX, y: calcY + 0.3, w: calcW, h: 0.4,
      fontSize: 12, fontFace: FONT_BODY, color: C.primaryLight,
      bold: true, align: "center", valign: "middle", charSpacing: 6, margin: 0,
    });
    s.addText("1.500 alunos × R$ 80 × 12 meses", {
      x: calcX, y: calcY + 0.8, w: calcW, h: 0.6,
      fontSize: 22, fontFace: FONT_BODY, color: C.textSecondary,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("= R$ 1,44 milhão de ARR", {
      x: calcX, y: calcY + 1.45, w: calcW, h: 0.8,
      fontSize: 44, fontFace: FONT_TITLE, color: C.primaryLight,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    // Mini-cards com ícones
    const mini = [
      { icon: I.chart, title: "MARGEM BRUTA", value: "~75%" },
      { icon: I.cycle, title: "CHURN B2B ALVO", value: "< 5% / ano" },
      { icon: I.coins, title: "TICKET MÉDIO", value: "R$ 120k / mês" },
    ];
    const miW = 3.8, miH = 1.2, miGap = 0.3;
    const miStartX = (W - (3 * miW + 2 * miGap)) / 2;
    mini.forEach((m, i) => {
      const x = miStartX + i * (miW + miGap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 5.6, w: miW, h: miH,
        fill: { color: C.bgLight }, line: { color: C.border, width: 1 }, rectRadius: 0.12,
      });
      // Ícone em círculo
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.25, y: 5.75, w: 0.9, h: 0.9,
        fill: { color: C.bgDark }, line: { color: C.primary, width: 1 },
      });
      s.addImage({
        data: m.icon, x: x + 0.47, y: 5.97, w: 0.46, h: 0.46,
      });
      s.addText(m.title, {
        x: x + 1.3, y: 5.65, w: miW - 1.5, h: 0.45,
        fontSize: 11, fontFace: FONT_BODY, color: C.textMuted,
        bold: true, charSpacing: 4, align: "left", valign: "bottom", margin: 0,
      });
      s.addText(m.value, {
        x: x + 1.3, y: 6.1, w: miW - 1.5, h: 0.55,
        fontSize: 20, fontFace: FONT_TITLE, color: C.textPrimary,
        bold: true, align: "left", valign: "top", margin: 0,
      });
    });

    addFooter(s, 6);
  }

  // =========================================================
  // SLIDE 7 · MERCADO
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "4 · Mercado", 0.6, 0.5);

    s.addText("Mercado bilionário subatendido.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
      fontSize: 40, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });

    const kpis = [
      { value: "2.400", label: "IES PRIVADAS NO BRASIL", fs: 52 },
      { value: "8,5 MI", label: "ALUNOS NO ENSINO SUPERIOR", fs: 52 },
      { value: "R$ 800M–1,5 BI", label: "TAM ANUAL BRASIL", fs: 28 },
    ];
    const kW = 3.8, kH = 2.4, kGap = 0.3;
    const kStartX = (W - (3 * kW + 2 * kGap)) / 2;
    kpis.forEach((k, i) => {
      const x = kStartX + i * (kW + kGap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.2, w: kW, h: kH,
        fill: { color: C.bgMid }, line: { color: C.accent, width: 1 }, rectRadius: 0.15,
      });
      s.addText(k.value, {
        x, y: 2.4, w: kW, h: 1.4,
        fontSize: k.fs, fontFace: FONT_TITLE, color: C.primaryLight,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(k.label, {
        x: x + 0.2, y: 3.8, w: kW - 0.4, h: 0.5,
        fontSize: 11, fontFace: FONT_BODY, color: C.textSecondary,
        bold: true, charSpacing: 4, align: "center", valign: "top", margin: 0,
      });
    });

    // Meta 36 meses
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 1.0, y: 5.1, w: W - 2.0, h: 1.7,
      fill: { color: C.primary }, line: { color: C.primaryLight, width: 2 }, rectRadius: 0.2,
    });
    s.addText("META · 36 MESES", {
      x: 1.0, y: 5.2, w: W - 2.0, h: 0.4,
      fontSize: 13, fontFace: FONT_BODY, color: C.textPrimary,
      bold: true, charSpacing: 8, align: "center", valign: "middle", margin: 0,
    });
    s.addText("50 IES  ·  75.000 alunos  ·  R$ 72 MI ARR", {
      x: 1.0, y: 5.7, w: W - 2.0, h: 1.0,
      fontSize: 36, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, 7);
  }

  // =========================================================
  // SLIDE 8 · TIME (Guilherme, Arthur, Rômulo)
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "Time · Quem constrói", 0.6, 0.5);

    s.addText("Faminto. Baixo custo. Skin no jogo máximo.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
      fontSize: 38, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });
    s.addText("Três estudantes vivendo a dor que estão resolvendo.", {
      x: 0.6, y: 1.85, w: W - 1.2, h: 0.5,
      fontSize: 15, fontFace: FONT_BODY, color: C.textSecondary,
      italic: true, align: "left", valign: "top", margin: 0,
    });

    const team = [
      { initial: "G", icon: I.pm, role: "PRODUCT MANAGER", name: "Guilherme",
        exp: "6 meses em produto · Aluno de IES privada" },
      { initial: "A", icon: I.dev, role: "DESENVOLVEDOR BACKEND", name: "Arthur",
        exp: "6 meses em engenharia · Aluno de IES privada" },
      { initial: "R", icon: I.stack, role: "DESENVOLVEDOR FRONTEND", name: "Rômulo",
        exp: "6 meses em engenharia · Aluno de IES privada" },
    ];
    const tW = 3.8, tH = 3.5, tGap = 0.3;
    const tStartX = (W - (3 * tW + 2 * tGap)) / 2;
    team.forEach((t, i) => {
      const x = tStartX + i * (tW + tGap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.6, w: tW, h: tH,
        fill: { color: C.bgMid }, line: { color: C.border, width: 1 }, rectRadius: 0.2,
      });
      // Avatar: círculo grande com ícone do papel dentro
      s.addShape(pres.shapes.OVAL, {
        x: x + (tW - 1.5) / 2, y: 2.9, w: 1.5, h: 1.5,
        fill: { color: C.primary }, line: { color: C.primaryLight, width: 2 },
      });
      // Inicial grande
      s.addText(t.initial, {
        x: x + (tW - 1.5) / 2, y: 2.9, w: 1.5, h: 1.5,
        fontSize: 48, fontFace: FONT_TITLE, color: C.textPrimary,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      // Nome
      s.addText(t.name, {
        x: x + 0.2, y: 4.6, w: tW - 0.4, h: 0.5,
        fontSize: 22, fontFace: FONT_TITLE, color: C.textPrimary,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      // Cargo
      s.addText(t.role, {
        x: x + 0.2, y: 5.1, w: tW - 0.4, h: 0.4,
        fontSize: 11, fontFace: FONT_BODY, color: C.primaryLight,
        bold: true, charSpacing: 4, align: "center", valign: "middle", margin: 0,
      });
      // Exp
      s.addText(t.exp, {
        x: x + 0.2, y: 5.5, w: tW - 0.4, h: 0.6,
        fontSize: 12, fontFace: FONT_BODY, color: C.textSecondary,
        align: "center", valign: "top", margin: 0,
      });
    });

    s.addText("Vamos construir 10× mais rápido e 5× mais barato que qualquer time sênior.", {
      x: 0.6, y: 6.5, w: W - 1.2, h: 0.5,
      fontSize: 15, fontFace: FONT_BODY, color: C.primaryLight,
      italic: true, align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, 8);
  }

  // =========================================================
  // SLIDE 9 · O PEDIDO
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);
    sectionBadge(s, "5 · O Pedido", 0.6, 0.5);

    s.addText("Captando R$ 500k via SAFE.", {
      x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
      fontSize: 38, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "left", valign: "top", margin: 0,
    });
    s.addText("Pré-money R$ 2,5M · Oversubscription até R$ 700k · Runway de 18 meses.", {
      x: 0.6, y: 1.85, w: W - 1.2, h: 0.5,
      fontSize: 16, fontFace: FONT_BODY, color: C.primaryLight,
      italic: true, align: "left", valign: "top", margin: 0,
    });

    const uses = [
      { pct: "40%", label: "Vendas & BD", desc: "Closers, parcerias\ninstitucionais" },
      { pct: "40%", label: "Engenharia", desc: "Completar produto\ne infraestrutura" },
      { pct: "20%", label: "Operações", desc: "Jurídico, compliance,\nLGPD, cloud" },
    ];
    const uW = 3.8, uH = 3.2, uGap = 0.3;
    const uStartX = (W - (3 * uW + 2 * uGap)) / 2;
    uses.forEach((u, i) => {
      const x = uStartX + i * (uW + uGap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.8, w: uW, h: uH,
        fill: { color: C.bgMid }, line: { color: C.border, width: 1 }, rectRadius: 0.2,
      });
      s.addShape(pres.shapes.OVAL, {
        x: x + (uW - 1.6) / 2, y: 3.0, w: 1.6, h: 1.6,
        fill: { color: C.bgDark }, line: { color: C.primary, width: 4 },
      });
      s.addText(u.pct, {
        x: x + (uW - 1.6) / 2, y: 3.0, w: 1.6, h: 1.6,
        fontSize: 28, fontFace: FONT_TITLE, color: C.primaryLight,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(u.label, {
        x: x + 0.2, y: 4.8, w: uW - 0.4, h: 0.5,
        fontSize: 18, fontFace: FONT_TITLE, color: C.textPrimary,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(u.desc, {
        x: x + 0.2, y: 5.3, w: uW - 0.4, h: 0.7,
        fontSize: 12, fontFace: FONT_BODY, color: C.textSecondary,
        align: "center", valign: "top", margin: 0,
      });
    });

    // CTA com ícone de flecha
    const ctaW = 6.5, ctaH = 0.75;
    const ctaX = (W - ctaW) / 2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: ctaX, y: 6.3, w: ctaW, h: ctaH,
      fill: { color: C.primary }, line: { color: C.primaryLight, width: 1 }, rectRadius: 0.35,
    });
    s.addText("Converse comigo após o pitch", {
      x: ctaX, y: 6.3, w: ctaW - 0.5, h: ctaH,
      fontSize: 18, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addImage({
      data: I.arrow,
      x: ctaX + ctaW - 0.6, y: 6.48, w: 0.4, h: 0.4,
    });

    addFooter(s, 9);
  }

  // =========================================================
  // SLIDE 10 · FECHAMENTO
  // =========================================================
  {
    const s = pres.addSlide();
    darkBg(s);

    s.addShape(pres.shapes.OVAL, {
      x: W - 5, y: -3, w: 8, h: 8,
      fill: { color: C.primary, transparency: 88 }, line: { color: C.primary, transparency: 100 },
    });
    s.addShape(pres.shapes.OVAL, {
      x: -3, y: H - 4, w: 7, h: 7,
      fill: { color: C.accent, transparency: 88 }, line: { color: C.accent, transparency: 100 },
    });

    s.addImage({
      path: path.join(ASSETS, "nura-logo.png"),
      x: (W - 3.0) / 2, y: 1.5, w: 3.0, h: 0.9,
      sizing: { type: "contain", w: 3.0, h: 0.9 },
    });

    s.addText("NURA", {
      x: 0, y: 2.7, w: W, h: 1.4,
      fontSize: 108, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "center", valign: "middle", charSpacing: 14,
    });

    s.addText("O sistema operacional da educação brasileira.", {
      x: 0, y: 4.3, w: W, h: 0.6,
      fontSize: 20, fontFace: FONT_BODY, color: C.primaryLight,
      italic: true, align: "center", valign: "middle",
    });

    s.addText("Obrigado.", {
      x: 0, y: 5.2, w: W, h: 0.8,
      fontSize: 40, fontFace: FONT_TITLE, color: C.textPrimary,
      bold: true, align: "center", valign: "middle",
    });

    // Contatos com bordas bem arredondadas — pill shape
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: (W - 9) / 2, y: 6.3, w: 9, h: 0.65,
      fill: { color: C.bgMid }, line: { color: C.primary, width: 1 }, rectRadius: 0.32,
    });
    s.addText("nextgenbrotherss@gmail.com   ·   www.nurasys.com.br", {
      x: (W - 9) / 2, y: 6.3, w: 9, h: 0.65,
      fontSize: 14, fontFace: FONT_BODY, color: C.textPrimary,
      align: "center", valign: "middle", charSpacing: 3, margin: 0,
    });
  }

  // =========================================================
  // TRANSIÇÕES — Morph suave entre slides
  // =========================================================
  // Aplicadas via XML direto no arquivo gerado (pptxgenjs não expõe isso nativamente).
  // Orientação para o usuário: abrir no PowerPoint, selecionar "Transições" → Morph/Fade.

  await pres.writeFile({ fileName: path.join(__dirname, "NURA - PITCH.pptx") });
  console.log("\n✅ Apresentação v3 gerada.\n");
}

build().catch(err => {
  console.error("❌ Erro:", err);
  process.exit(1);
});
