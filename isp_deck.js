const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaShieldAlt, FaSitemap, FaExclamationTriangle, FaBrain,
  FaRobot, FaBuilding, FaSearch, FaTools, FaLightbulb,
  FaChartLine, FaNetworkWired, FaEye, FaClipboardList, FaKey
} = require("react-icons/fa");
const { MdOutlineAccountTree, MdSecurity } = require("react-icons/md");

// ─── Palette ─────────────────────────────────────────────────────────────────
const C = {
  dark:    "1A2340",   // deep navy
  mid:     "243261",   // mid navy
  accent:  "E84C4C",   // alert red
  gold:    "F5A623",   // warning gold
  teal:    "2EC4B6",   // calm teal
  white:   "FFFFFF",
  offWhite:"F4F6FA",
  muted:   "8A97B0",
  card:    "EEF1F8",
};

// ─── Icon helper ─────────────────────────────────────────────────────────────
async function iconPng(IconComp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const makeShadow = () => ({
  type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.14
});

// ─── Main ─────────────────────────────────────────────────────────────────────
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 1 — TITLE
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Left accent block
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 3.8, h: 5.625,
      fill: { color: C.mid }
    });

    // Red "danger" shape top-right corner
    s.addShape(pres.shapes.RECTANGLE, {
      x: 7.5, y: 0, w: 2.5, h: 1.1,
      fill: { color: C.accent }
    });

    // Shield icon top-right
    const shieldIcon = await iconPng(FaShieldAlt, "#FFFFFF", 256);
    s.addImage({ data: shieldIcon, x: 8.2, y: 0.1, w: 0.85, h: 0.85 });

    // Left: label tag
    s.addText("ISP · TPRM · CYBERSÉCURITÉ", {
      x: 0.3, y: 0.35, w: 3.2, h: 0.35,
      fontSize: 9, bold: true, color: C.teal,
      charSpacing: 3, align: "left", margin: 0
    });

    // Left: main visual icon
    const eyeIcon = await iconPng(FaEye, "#2EC4B6", 256);
    s.addImage({ data: eyeIcon, x: 0.6, y: 0.9, w: 2.4, h: 2.4 });

    // Left: slide number label
    s.addText("2025", {
      x: 0.3, y: 4.9, w: 3.2, h: 0.4,
      fontSize: 10, color: C.muted, align: "left", margin: 0
    });

    // Right: main title
    s.addText("Les Angles Morts", {
      x: 4.1, y: 0.7, w: 5.6, h: 1.0,
      fontSize: 40, bold: true, color: C.white,
      align: "left", margin: 0
    });
    s.addText("des Missions ISP", {
      x: 4.1, y: 1.6, w: 5.6, h: 0.9,
      fontSize: 40, bold: true, color: C.teal,
      align: "left", margin: 0
    });

    // Separator line simulation via narrow rect
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.1, y: 2.6, w: 5.0, h: 0.04,
      fill: { color: C.accent }
    });

    s.addText("Gouvernance · Conformité · Risques émergents · Outillage", {
      x: 4.1, y: 2.75, w: 5.6, h: 0.45,
      fontSize: 12, color: C.muted, italic: true, align: "left", margin: 0
    });

    s.addText([
      { text: "Consultant Cybersécurité Senior", options: { bold: true, breakLine: true } },
      { text: "Gestion des risques fournisseurs (TPRM) — Vision terrain" }
    ], {
      x: 4.1, y: 3.35, w: 5.6, h: 0.8,
      fontSize: 12, color: C.white, align: "left", margin: 0
    });

    // Bottom row: 3 tags
    const tags = ["Indépendance vs Dépendance", "Conformité vs Réalité", "Gouvernance & RFP"];
    tags.forEach((t, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 4.1 + i * 1.95, y: 4.55, w: 1.8, h: 0.45,
        fill: { color: C.mid }, rectRadius: 0.05
      });
      s.addText(t, {
        x: 4.1 + i * 1.95, y: 4.55, w: 1.8, h: 0.45,
        fontSize: 8, color: C.teal, align: "center", bold: true, margin: 0
      });
    });

    s.addNotes("Slide d'introduction. Cadrer le périmètre : missions ISP dans un GIE avec dépendance forte au Groupe. Annoncer les 6 angles morts couverts.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 2 — ÉCOSYSTÈME : Indépendance vs Dépendance
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    // Header band
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.dark }
    });
    s.addText("01", {
      x: 0.25, y: 0.1, w: 0.6, h: 0.55,
      fontSize: 11, bold: true, color: C.accent, align: "center", margin: 0
    });
    s.addText("L'ÉCOSYSTÈME GIE — INDÉPENDANCE VS DÉPENDANCE", {
      x: 1.0, y: 0.15, w: 8.5, h: 0.5,
      fontSize: 15, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 1
    });

    const siteIcon = await iconPng(FaSitemap, "#2EC4B6", 256);
    s.addImage({ data: siteIcon, x: 8.5, y: 0.1, w: 0.6, h: 0.6 });

    // Two columns
    // Left: GIE card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: 1.05, w: 4.4, h: 3.8,
      fill: { color: C.white }, rectRadius: 0.08,
      shadow: makeShadow()
    });
    s.addText("GIE (Local)", {
      x: 0.5, y: 1.15, w: 4.0, h: 0.45,
      fontSize: 14, bold: true, color: C.dark, align: "left", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 1.58, w: 4.0, h: 0.03,
      fill: { color: C.teal }
    });
    s.addText([
      { text: "✔  Autonomie de gestion des risques", options: { breakLine: true } },
      { text: "✔  Seuils de criticité propres (impact financier/réglementaire moindre)", options: { breakLine: true } },
      { text: "✔  Réévaluation systématique des services déjà audités Groupe", options: { breakLine: true } },
      { text: "⚠  Tensions avec les équipes Groupe challengées sur leurs analyses" }
    ], {
      x: 0.5, y: 1.7, w: 4.0, h: 2.9,
      fontSize: 11.5, color: C.dark, align: "left", margin: 0,
      paraSpaceAfter: 8
    });

    // Right: Groupe card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.3, y: 1.05, w: 4.4, h: 3.8,
      fill: { color: C.white }, rectRadius: 0.08,
      shadow: makeShadow()
    });
    s.addText("Groupe Operations (GO)", {
      x: 5.5, y: 1.15, w: 4.0, h: 0.45,
      fontSize: 14, bold: true, color: C.dark, align: "left", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: 1.58, w: 4.0, h: 0.03,
      fill: { color: C.accent }
    });
    s.addText([
      { text: "◉  Gestion des datacenters (fournisseur de services critique)", options: { breakLine: true } },
      { text: "◉  Politiques de sécurité divergentes de celles du GIE", options: { breakLine: true } },
      { text: "◉  Peu collaboratif face aux remises en question", options: { breakLine: true } },
      { text: "◉  Périmètre de responsabilité flou entre entités" }
    ], {
      x: 5.5, y: 1.7, w: 4.0, h: 2.9,
      fontSize: 11.5, color: C.dark, align: "left", margin: 0,
      paraSpaceAfter: 8
    });

    // Arrow between columns
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.72, y: 2.6, w: 0.56, h: 0.04,
      fill: { color: C.muted }
    });
    s.addText("↔", {
      x: 4.65, y: 2.45, w: 0.7, h: 0.35,
      fontSize: 20, color: C.muted, align: "center", margin: 0
    });

    // Bottom callout
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: 5.0, w: 9.4, h: 0.4,
      fill: { color: C.accent, transparency: 85 }, rectRadius: 0.05
    });
    s.addText("ANGLE MORT : La \"double évaluation\" consomme des ressources sans valeur ajoutée incrémentale et fragilise la relation inter-entités.", {
      x: 0.5, y: 5.0, w: 9.0, h: 0.4,
      fontSize: 10, color: C.accent, bold: true, align: "left", italic: true, margin: 0
    });

    s.addNotes("Insister sur la tension structurelle : le GIE doit se conformer aux politiques Groupe tout en maintenant sa propre gouvernance. Exemple concret : audit datacenter remis en question.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 3 — TPRM : Standardisation & Automatisation
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.mid }
    });
    s.addText("02", {
      x: 0.25, y: 0.1, w: 0.6, h: 0.55,
      fontSize: 11, bold: true, color: C.gold, align: "center", margin: 0
    });
    s.addText("LE PROCESSUS TPRM — STANDARDISATION & AUTOMATISATION", {
      x: 1.0, y: 0.15, w: 8.5, h: 0.5,
      fontSize: 15, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 1
    });

    const toolsIcon = await iconPng(FaTools, "#F5A623", 256);
    s.addImage({ data: toolsIcon, x: 8.5, y: 0.1, w: 0.6, h: 0.6 });

    // Process flow: 3 steps
    const steps = [
      { num: "1", title: "Certification ISO 27001", desc: "Socle de référence pour évaluer la maturité formelle du fournisseur.", color: C.teal },
      { num: "2", title: "Questionnaire de sécurité interne", desc: "Couche complémentaire pour tester la réalité opérationnelle au-delà des certifications.", color: C.gold },
      { num: "3", title: "Scoring automatisé des écarts", desc: "Identification rapide des gaps de maturité selon les réponses fournies.", color: C.accent },
    ];

    steps.forEach((step, i) => {
      const x = 0.4 + i * 3.15;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.1, w: 2.9, h: 3.2,
        fill: { color: C.mid }, rectRadius: 0.08,
        shadow: makeShadow()
      });
      // Circle number
      s.addShape(pres.shapes.OVAL, {
        x: x + 1.05, y: 1.2, w: 0.8, h: 0.8,
        fill: { color: step.color }
      });
      s.addText(step.num, {
        x: x + 1.05, y: 1.2, w: 0.8, h: 0.8,
        fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle", margin: 0
      });
      s.addText(step.title, {
        x: x + 0.15, y: 2.15, w: 2.6, h: 0.7,
        fontSize: 12, bold: true, color: step.color, align: "center", margin: 0
      });
      s.addText(step.desc, {
        x: x + 0.15, y: 2.95, w: 2.6, h: 1.2,
        fontSize: 11, color: C.offWhite, align: "center", margin: 0
      });
    });

    // Arrows between steps
    [1, 2].forEach(i => {
      s.addText("→", {
        x: 0.4 + i * 3.15 - 0.5, y: 2.3, w: 0.55, h: 0.4,
        fontSize: 22, color: C.muted, align: "center", margin: 0
      });
    });

    // Callout
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 4.55, w: 9.2, h: 0.85,
      fill: { color: "1E2D5A" }, rectRadius: 0.08
    });
    s.addText([
      { text: "Limite structurelle : ", options: { bold: true, color: C.gold } },
      { text: "L'automatisation du scoring réduit la charge de tri, mais ne résout pas l'absence d'outil TPRM dédié. La traçabilité reste fragile tant que tout repose sur Excel/SharePoint.", options: { color: C.offWhite } }
    ], {
      x: 0.65, y: 4.6, w: 8.7, h: 0.75,
      fontSize: 11, align: "left", margin: 0
    });

    s.addNotes("Souligner que cette approche en 3 temps est pragmatique mais souffre d'un manque d'outillage. Le scoring automatisé est un palliatif, pas une solution pérenne.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 4 — Dissonance Cognitive : Conformité vs Réalité Cyber
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.dark }
    });
    s.addText("03", {
      x: 0.25, y: 0.1, w: 0.6, h: 0.55,
      fontSize: 11, bold: true, color: C.accent, align: "center", margin: 0
    });
    s.addText("DISSONANCE COGNITIVE — CONFORMITÉ vs RÉALITÉ CYBER", {
      x: 1.0, y: 0.15, w: 8.5, h: 0.5,
      fontSize: 15, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 1
    });

    const brainIcon = await iconPng(FaBrain, "#E84C4C", 256);
    s.addImage({ data: brainIcon, x: 8.5, y: 0.1, w: 0.6, h: 0.6 });

    // Left column: problem
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: 1.05, w: 4.5, h: 2.5,
      fill: { color: C.white }, rectRadius: 0.08, shadow: makeShadow()
    });
    s.addText("Le Problème", {
      x: 0.5, y: 1.15, w: 4.1, h: 0.4,
      fontSize: 13, bold: true, color: C.dark, align: "left", margin: 0
    });
    s.addText([
      { text: "La criticité cybersécurité est \"otage\" d'autres domaines :", options: { breakLine: true, bold: true } },
      { text: " ", options: { breakLine: true } },
      { text: "  • Héritée de la Data Privacy (DP)", options: { breakLine: true } },
      { text: "  • DP se base sur la volumétrie des données", options: { breakLine: true } },
      { text: "  • Pas sur leur nature ou leur sensibilité réelle", options: { breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "Résultat : la cybersécurité n'est pas évaluée pour ses propres enjeux." }
    ], {
      x: 0.5, y: 1.65, w: 4.1, h: 1.8,
      fontSize: 11, color: C.dark, align: "left", margin: 0
    });

    // Right column: case study
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.05, w: 4.5, h: 2.5,
      fill: { color: C.white }, rectRadius: 0.08, shadow: makeShadow()
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.05, w: 4.5, h: 0.42,
      fill: { color: C.gold, transparency: 20 }, rectRadius: 0.08
    });
    s.addText("📌  Cas réel : Social Listening", {
      x: 5.4, y: 1.1, w: 4.1, h: 0.32,
      fontSize: 11, bold: true, color: C.dark, align: "left", margin: 0
    });
    s.addText([
      { text: "Outil de veille sur réseaux sociaux :", options: { breakLine: true, bold: true } },
      { text: " ", options: { breakLine: true } },
      { text: "  Données publiques + volume élevé", options: { breakLine: true } },
      { text: "  → Classé \"TRÈS CRITIQUE\" par DP", options: { breakLine: true, color: C.accent, bold: true } },
      { text: " ", options: { breakLine: true } },
      { text: "  Risque cyber réel : limité", options: { breakLine: true } },
      { text: "  → Criticité disproportionnée vs. enjeu" }
    ], {
      x: 5.4, y: 1.55, w: 4.1, h: 1.9,
      fontSize: 11, color: C.dark, align: "left", margin: 0
    });

    // Bottom: disconnect diagram
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: 3.75, w: 9.4, h: 1.55,
      fill: { color: C.dark }, rectRadius: 0.08
    });
    s.addText("Déconnexion structurelle", {
      x: 0.55, y: 3.85, w: 4.0, h: 0.35,
      fontSize: 12, bold: true, color: C.teal, align: "left", margin: 0
    });

    // 3 mini boxes
    const items = [
      { label: "Politique AXA", sub: "Conformité globale", col: C.muted },
      { label: "Data Privacy", sub: "Volume ≠ sensibilité", col: C.gold },
      { label: "Réalité Cyber", sub: "Risque opérationnel", col: C.accent },
    ];
    items.forEach((item, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.4 + i * 3.15, y: 4.25, w: 2.9, h: 0.85,
        fill: { color: item.col, transparency: 75 }, rectRadius: 0.06
      });
      s.addText(item.label, {
        x: 0.4 + i * 3.15, y: 4.28, w: 2.9, h: 0.35,
        fontSize: 11, bold: true, color: C.white, align: "center", margin: 0
      });
      s.addText(item.sub, {
        x: 0.4 + i * 3.15, y: 4.63, w: 2.9, h: 0.35,
        fontSize: 10, color: C.offWhite, align: "center", italic: true, margin: 0
      });
    });
    // arrows
    s.addText("≠", { x: 3.35, y: 4.35, w: 0.8, h: 0.5, fontSize: 20, color: C.accent, align: "center", margin: 0 });
    s.addText("≠", { x: 6.5, y: 4.35, w: 0.8, h: 0.5, fontSize: 20, color: C.accent, align: "center", margin: 0 });

    s.addNotes("Point clé : la criticité cyber ne doit pas être un sous-produit de la classification DP. Illustrer avec le cas Social Listening. Ouvrir sur la nécessité d'un référentiel cyber autonome.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 5 — Angles Morts : IA & 4th Party
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.mid }
    });
    s.addText("04", {
      x: 0.25, y: 0.1, w: 0.6, h: 0.55,
      fontSize: 11, bold: true, color: C.gold, align: "center", margin: 0
    });
    s.addText("NOUVELLES TECHNOLOGIES & RISQUE 4TH PARTY", {
      x: 1.0, y: 0.15, w: 8.5, h: 0.5,
      fontSize: 15, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 1
    });

    const robotIcon = await iconPng(FaRobot, "#F5A623", 256);
    s.addImage({ data: robotIcon, x: 8.5, y: 0.1, w: 0.6, h: 0.6 });

    // Left: IA / Shadow IT
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: 1.1, w: 4.5, h: 4.25,
      fill: { color: C.mid }, rectRadius: 0.08, shadow: makeShadow()
    });
    s.addShape(pres.shapes.OVAL, {
      x: 0.5, y: 1.2, w: 0.65, h: 0.65,
      fill: { color: C.gold }
    });
    const aiIcon = await iconPng(FaBrain, "#1A2340", 256);
    s.addImage({ data: aiIcon, x: 0.55, y: 1.25, w: 0.55, h: 0.55 });

    s.addText("IA & Shadow IT", {
      x: 1.3, y: 1.25, w: 3.3, h: 0.5,
      fontSize: 14, bold: true, color: C.gold, align: "left", margin: 0
    });

    s.addText([
      { text: "Problème central :", options: { bold: true, color: C.offWhite, breakLine: true } },
      { text: "Le fournisseur ne reconnaît pas lui-même l'usage d'IA dans sa solution.", options: { color: C.offWhite, breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "Exemple :", options: { bold: true, color: C.gold, breakLine: true } },
      { text: "Bibliothèques ML pour l'analyse de sentiments embarquées sans déclaration formelle.", options: { color: C.offWhite, breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "Conséquences :", options: { bold: true, color: C.offWhite, breakLine: true } },
      { text: "  • Périmètre de traitement des données opaque", options: { color: C.muted, breakLine: true } },
      { text: "  • Risque de biais et de fuite hors périmètre contractuel", options: { color: C.muted, breakLine: true } },
      { text: "  • Impossibilité de notifier ni d'évaluer le risque réel" }
    ], {
      x: 0.5, y: 1.95, w: 4.1, h: 3.1,
      fontSize: 11, align: "left", margin: 0, paraSpaceAfter: 3
    });

    // Right: 4th Party
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.1, w: 4.5, h: 4.25,
      fill: { color: C.mid }, rectRadius: 0.08, shadow: makeShadow()
    });
    s.addShape(pres.shapes.OVAL, {
      x: 5.4, y: 1.2, w: 0.65, h: 0.65,
      fill: { color: C.accent }
    });
    const netIcon = await iconPng(FaNetworkWired, "#F4F6FA", 256);
    s.addImage({ data: netIcon, x: 5.45, y: 1.25, w: 0.55, h: 0.55 });

    s.addText("Risque 4th Party", {
      x: 6.2, y: 1.25, w: 3.3, h: 0.5,
      fontSize: 14, bold: true, color: C.accent, align: "left", margin: 0
    });

    s.addText([
      { text: "Chaîne de sous-traitance invisible :", options: { bold: true, color: C.offWhite, breakLine: true } },
      { text: "Elle dépasse le périmètre cartographié sans visibilité contractuelle.", options: { color: C.offWhite, breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "Zones de rupture :", options: { bold: true, color: C.accent, breakLine: true } },
      { text: "  • Prestataires de maintenance OT non déclarés", options: { color: C.muted, breakLine: true } },
      { text: "  • Services tiers accédant à des données sensibles", options: { color: C.muted, breakLine: true } },
      { text: "  • Absence de clause contractuelle sur les sous-traitants", options: { color: C.muted, breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "Résultat :", options: { bold: true, color: C.offWhite, breakLine: true } },
      { text: "Rupture de la chaîne de confiance sans mécanisme de détection." }
    ], {
      x: 5.4, y: 1.95, w: 4.1, h: 3.1,
      fontSize: 11, align: "left", margin: 0, color: C.offWhite, paraSpaceAfter: 3
    });

    s.addNotes("Ces deux risques sont structurellement hors-radar des questionnaires TPRM classiques. Illustrer avec un cas IA non déclarée. Souligner l'urgence de clauses contractuelles spécifiques IA et 4th Party.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 6 — Gouvernance : Le Péché Originel du RFP
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.dark }
    });
    s.addText("05", {
      x: 0.25, y: 0.1, w: 0.6, h: 0.55,
      fontSize: 11, bold: true, color: C.accent, align: "center", margin: 0
    });
    s.addText("GOUVERNANCE — LE \"PÉCHÉ ORIGINEL\" DU RFP", {
      x: 1.0, y: 0.15, w: 8.5, h: 0.5,
      fontSize: 15, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 1
    });

    const buildingIcon = await iconPng(FaBuilding, "#E84C4C", 256);
    s.addImage({ data: buildingIcon, x: 8.5, y: 0.1, w: 0.6, h: 0.6 });

    // Timeline: 4 stages
    const stages = [
      { label: "Idéation", sub: "Besoin métier identifié", icon: "💡", bad: false },
      { label: "RFP / Sélection", sub: "Critères : finance & commercial", icon: "💰", bad: true },
      { label: "Contractualisation", sub: "Sécurité en bout de chaîne", icon: "📄", bad: true },
      { label: "Remédiation", sub: "ISP rattrape le niveau requis", icon: "🚒", bad: true },
    ];

    // Timeline line
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 2.85, w: 8.4, h: 0.06,
      fill: { color: C.muted }
    });

    stages.forEach((stage, i) => {
      const x = 0.7 + i * 2.2;
      // Circle on timeline
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.65, y: 2.6, w: 0.55, h: 0.55,
        fill: { color: stage.bad ? C.accent : C.teal }
      });
      s.addText(stage.icon, {
        x: x + 0.65, y: 2.6, w: 0.55, h: 0.55,
        fontSize: 14, align: "center", valign: "middle", margin: 0
      });

      // Card above or below alternating
      const isAbove = i % 2 === 0;
      const cardY = isAbove ? 1.05 : 3.35;

      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: cardY, w: 2.0, h: 1.3,
        fill: { color: stage.bad ? "FFF0F0" : "F0FFF8" }, rectRadius: 0.07,
        shadow: makeShadow()
      });
      s.addText(stage.label, {
        x, y: cardY + 0.1, w: 2.0, h: 0.4,
        fontSize: 12, bold: true, color: stage.bad ? C.accent : C.teal,
        align: "center", margin: 0
      });
      s.addText(stage.sub, {
        x, y: cardY + 0.55, w: 2.0, h: 0.65,
        fontSize: 10, color: C.dark, align: "center", italic: true, margin: 0
      });

      // Connector line
      const lineY1 = isAbove ? cardY + 1.3 : 3.15;
      const lineY2 = isAbove ? 2.6 : cardY;
      s.addShape(pres.shapes.LINE, {
        x: x + 0.97, y: lineY1, w: 0, h: lineY2 - lineY1,
        line: { color: stage.bad ? C.accent : C.teal, width: 1.5, dashType: "dash" }
      });
    });

    // Bottom callout
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: 4.9, w: 9.4, h: 0.52,
      fill: { color: C.dark }, rectRadius: 0.07
    });
    s.addText([
      { text: "Conséquence : ", options: { bold: true, color: C.accent } },
      { text: "La sécurité n'est plus un critère éliminatoire mais un centre de remédiation forcé. Chaque onboarding devient une course contre la montre.", options: { color: C.offWhite } }
    ], {
      x: 0.55, y: 4.9, w: 9.0, h: 0.52,
      fontSize: 11, align: "left", margin: 0
    });

    s.addNotes("Faire ressortir que ce problème est gouvernancenel et non technique. La sécurité doit être un critère éliminatoire dès le RFP, pas une case à cocher en fin de processus.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 7 — Faillite de l'Outillage
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.mid }
    });
    s.addText("06", {
      x: 0.25, y: 0.1, w: 0.6, h: 0.55,
      fontSize: 11, bold: true, color: C.gold, align: "center", margin: 0
    });
    s.addText("FAILLITE DE L'OUTILLAGE & PERTE DE TRAÇABILITÉ", {
      x: 1.0, y: 0.15, w: 8.5, h: 0.5,
      fontSize: 15, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 1
    });

    const searchIcon = await iconPng(FaSearch, "#F5A623", 256);
    s.addImage({ data: searchIcon, x: 8.5, y: 0.1, w: 0.6, h: 0.6 });

    // Left: current state
    s.addText("État actuel", {
      x: 0.3, y: 1.05, w: 4.4, h: 0.4,
      fontSize: 13, bold: true, color: C.accent, align: "left", margin: 0
    });

    const currentItems = [
      "Fichiers Excel & PowerPoint sur SharePoint",
      "Absence de versioning et de traçabilité",
      "Multiplicaion de trackers manuels",
      "Réconciliation chronophage d'informations",
      "Réunions de coordination au détriment de l'analyse",
    ];
    currentItems.forEach((item, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: 1.55 + i * 0.68, w: 4.5, h: 0.58,
        fill: { color: C.mid }, rectRadius: 0.06
      });
      s.addShape(pres.shapes.OVAL, {
        x: 0.4, y: 1.62 + i * 0.68, w: 0.38, h: 0.38,
        fill: { color: C.accent }
      });
      s.addText("✕", {
        x: 0.4, y: 1.62 + i * 0.68, w: 0.38, h: 0.38,
        fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0
      });
      s.addText(item, {
        x: 0.9, y: 1.58 + i * 0.68, w: 3.8, h: 0.52,
        fontSize: 10.5, color: C.offWhite, align: "left", valign: "middle", margin: 0
      });
    });

    // Right: impact + transition
    s.addText("Impact ISP", {
      x: 5.2, y: 1.05, w: 4.5, h: 0.4,
      fontSize: 13, bold: true, color: C.gold, align: "left", margin: 0
    });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.55, w: 4.5, h: 1.5,
      fill: { color: C.mid }, rectRadius: 0.08
    });

    // Donut-like stat
    s.addShape(pres.shapes.OVAL, {
      x: 5.4, y: 1.65, w: 1.2, h: 1.2,
      fill: { color: C.accent, transparency: 20 }
    });
    s.addText("70%", {
      x: 5.4, y: 1.65, w: 1.2, h: 0.6,
      fontSize: 20, bold: true, color: C.white, align: "center", valign: "bottom", margin: 0
    });
    s.addText("tâches\nadmin", {
      x: 5.4, y: 2.25, w: 1.2, h: 0.4,
      fontSize: 9, color: C.offWhite, align: "center", margin: 0
    });

    s.addText([
      { text: "La charge administrative représente l'essentiel du temps ISP.", options: { breakLine: true, bold: true, color: C.offWhite } },
      { text: " ", options: { breakLine: true } },
      { text: "L'analyse de risque profonde est sacrifiée au profit de la réconciliation de données.", options: { color: C.muted } }
    ], {
      x: 6.8, y: 1.65, w: 2.8, h: 1.3,
      fontSize: 11, align: "left", margin: 0
    });

    // ProcessUnity note
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 3.2, w: 4.5, h: 2.1,
      fill: { color: "1E2D5A" }, rectRadius: 0.08
    });
    s.addText("ProcessUnity — en cours d'implémentation", {
      x: 5.4, y: 3.3, w: 4.1, h: 0.4,
      fontSize: 12, bold: true, color: C.teal, align: "left", margin: 0
    });
    s.addText([
      { text: "Statut : ", options: { bold: true, color: C.gold, breakLine: false } },
      { text: "Non adapté aux besoins réels du GIE à ce jour", options: { color: C.offWhite, breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "Risque : ", options: { bold: true, color: C.accent, breakLine: false } },
      { text: "Outil imposé sans co-construction → adoption limitée", options: { color: C.offWhite } }
    ], {
      x: 5.4, y: 3.8, w: 4.1, h: 1.3,
      fontSize: 11, align: "left", margin: 0
    });

    s.addNotes("Quantifier si possible le temps passé en réconciliation vs. analyse. ProcessUnity est une opportunité mais son implémentation doit être pilotée par les besoins GIE, pas imposée par le Groupe.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 8 — Recommandations Stratégiques
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.85,
      fill: { color: C.accent }
    });
    s.addText("RECOMMANDATIONS STRATÉGIQUES", {
      x: 0.5, y: 0.15, w: 8.8, h: 0.55,
      fontSize: 17, bold: true, color: C.white, align: "left", margin: 0, charSpacing: 2
    });

    const lightIcon = await iconPng(FaLightbulb, "#FFFFFF", 256);
    s.addImage({ data: lightIcon, x: 8.6, y: 0.1, w: 0.6, h: 0.6 });

    // 3 recommendation cards
    const recs = [
      {
        num: "01",
        title: "ANTICIPER",
        sub: "Shift Left Security",
        color: C.teal,
        points: [
          "Intégrer la sécurité dès l'idéation du projet",
          "Critère éliminatoire dans le RFP avant sélection financière",
          "Gate de sécurité obligatoire avant contractualisation",
        ]
      },
      {
        num: "02",
        title: "CLARIFIER",
        sub: "RACI & Responsabilités",
        color: C.gold,
        points: [
          "Définir un RACI formel : DP / Sécurité / Résilience",
          "Éviter les renvois de responsabilités inter-domaines",
          "Construire un référentiel de criticité cyber autonome",
        ]
      },
      {
        num: "03",
        title: "AUTOMATISER",
        sub: "Sortir du tout-Excel",
        color: C.accent,
        points: [
          "Piloter l'implémentation ProcessUnity par les besoins GIE",
          "Libérer la capacité d'analyse à haute valeur ajoutée",
          "Intégrer clauses IA et 4th Party dans les contrats fournisseurs",
        ]
      }
    ];

    recs.forEach((rec, i) => {
      const x = 0.3 + i * 3.2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.05, w: 3.0, h: 4.3,
        fill: { color: C.mid }, rectRadius: 0.1, shadow: makeShadow()
      });

      // Number badge
      s.addShape(pres.shapes.OVAL, {
        x: x + 1.1, y: 1.12, w: 0.8, h: 0.8,
        fill: { color: rec.color }
      });
      s.addText(rec.num, {
        x: x + 1.1, y: 1.12, w: 0.8, h: 0.8,
        fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle", margin: 0
      });

      s.addText(rec.title, {
        x: x + 0.1, y: 2.05, w: 2.8, h: 0.45,
        fontSize: 16, bold: true, color: rec.color, align: "center", margin: 0
      });
      s.addText(rec.sub, {
        x: x + 0.1, y: 2.5, w: 2.8, h: 0.3,
        fontSize: 10, color: C.muted, align: "center", italic: true, margin: 0
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.5, y: 2.85, w: 2.0, h: 0.03,
        fill: { color: rec.color, transparency: 50 }
      });

      rec.points.forEach((pt, j) => {
        s.addText("→  " + pt, {
          x: x + 0.15, y: 3.0 + j * 0.73, w: 2.7, h: 0.65,
          fontSize: 10.5, color: C.offWhite, align: "left", margin: 0
        });
      });
    });

    s.addNotes("Conclure sur le message clé : ces trois axes sont complémentaires. Anticiper évite la remédiation. Clarifier débloque les tensions inter-domaines. Automatiser libère la valeur ajoutée. Court terme : RACI + gate RFP. Moyen terme : ProcessUnity piloté par GIE.");
  }

  // ══════════════════════════════════════════════════════════════════════════
  // Write file
  // ══════════════════════════════════════════════════════════════════════════
  await pres.writeFile({ fileName: "/mnt/user-data/outputs/ISP_Angles_Morts.pptx" });
  console.log("✅ Presentation written.");
}

build().catch(err => {
  console.error("ERROR:", err);
  process.exit(1);
});
