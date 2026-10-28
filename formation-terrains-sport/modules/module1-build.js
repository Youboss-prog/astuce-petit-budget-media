// Module 1 — Sol, eau, air : la physique du substrat sableux des terrains de sport
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const { applyTheme } = require("/root/.claude/skills/synced/7ad88e06-8b42-40df-80c5-767aa6e4ff67_95b44de5-b5bc-4944-a291-4ab47213765c/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "Module1_Sol_eau_air.pptx";

const THEME = {
  name: "Terrain - Sol eau air",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1C2B22", // vert-noir profond
    lt1: "FFFFFF",
    dk2: "4A5A4F", // gris-vert (texte secondaire)
    lt2: "EEF2EC", // teinte claire des cartes
    accent1: "2F7D4A", // gazon / racines
    accent2: "C8962E", // sable
    accent3: "2B78A8", // eau
    accent4: "8FC1DB", // air
    accent5: "B83A2E", // alerte / erreurs
    accent6: "7A5C3E", // limon / sol fin
    hlink: "2B78A8",
    folHlink: "6D5A8E",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.author = "Formation terrains de sport";
pres.title = "Module 1 — Sol, eau, air";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITRE",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.55, w: 8.8, h: 1.1, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.8, w: 8.8, h: 0.9, fontSize: 18, color: C.accent4, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.7, w: 8.8, h: 0.5, fontSize: 16, color: C.accent2, bold: true, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.2, w: 8.8, h: 1.2, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Module 1 · Sol, eau, air", options: { x: 0.6, y: 5.15, w: 5, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ],
});
pres.defineSlideMaster({
  title: "SOMBRE",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.28, w: 9.0, h: 0.72, fontSize: 30, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Module 1 · Sol, eau, air", options: { x: 0.5, y: 5.2, w: 5, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENU",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.5, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.28, w: 9.0, h: 0.72, fontSize: 30, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Module 1 · Sol, eau, air", options: { x: 0.5, y: 5.2, w: 5, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.text2, align: "right" },
});

// ---------- Helpers ----------
async function iconData(Icon, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const ICONS = {};
async function loadIcons() {
  const list = {
    sol: gi.GiStonePile, eau: fa.FaTint, air: fa.FaWind, racine: gi.GiPlantRoots, jeu: fa.FaFutbol,
    duree: fa.FaClock, cout: fa.FaCoins, dispo: fa.FaRunning, cible: fa.FaBullseye, loupe: fa.FaSearch,
    alerte: fa.FaExclamationTriangle, livre: fa.FaBook, check: fa.FaClipboardCheck, ampoule: fa.FaLightbulb,
    labo: fa.FaFlask, regle: fa.FaRulerVertical, couches: fa.FaLayerGroup, therm: fa.FaThermometerHalf,
    balance: fa.FaBalanceScale, oeil: fa.FaEye, outils: fa.FaTools, pousse: fa.FaSeedling, question: fa.FaQuestion,
  };
  for (const [k, v] of Object.entries(list)) ICONS[k] = await iconData(v);
}
let objN = 0;
const nm = (p) => `${p}-${++objN}`;
function iconCircle(slide, key, x, y, d, fill) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: nm("pastille") });
  const p = d * 0.25;
  slide.addImage({ data: ICONS[key], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icone"), altText: key });
}
function card(slide, x, y, w, h, fill = C.background2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: nm("carte") });
}
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", objectName: nm("texte") }, o));
}
function content(title, section) {
  const s = pres.addSlide({ masterName: "CONTENU", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
}
function sectionSlide(label, title, section) {
  pres.addSection({ title: section });
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle: section });
  s.addText(label, { placeholder: "body" });
  s.addText(title, { placeholder: "title" });
  return s;
}

async function build() {
  await loadIcons();

  // ===== 1. Titre =====
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITRE", sectionTitle: "Introduction" });
  s.addText("Sol, eau, air", { placeholder: "title" });
  s.addText("Module 1 · La physique du substrat sableux des terrains de sport", { placeholder: "body" });
  iconCircle(s, "sol", 0.6, 4.15, 0.6, C.accent2);
  iconCircle(s, "eau", 1.35, 4.15, 0.6, C.accent3);
  iconCircle(s, "air", 2.1, 4.15, 0.6, C.accent4);
  iconCircle(s, "racine", 2.85, 4.15, 0.6, C.accent1);
  txt(s, "Phase 1 · Fondations — durée conseillée : 4 semaines à 5 h", { x: 3.7, y: 4.25, w: 5.7, h: 0.4, fontSize: 14, color: C.background1, valign: "middle" });
  s.addNotes(
    "NOTES DU FORMATEUR\n\n" +
    "Ce module est la fondation de toute la formation. L'arrosage (M3), la chimie du sol (M5), la fertilisation (M6) et l'aération (M7) reposent tous sur ce qui est expliqué ici : comment l'eau et l'air circulent dans un sol sableux, et ce que le piétinement y change.\n\n" +
    "Il répond directement à deux lacunes repérées pendant le diagnostic : la question 3 (sable contre limon, restée sans réponse) et l'erreur E5 (« un sol compacté manque d'eau »). Il prépare aussi la correction de l'erreur E1 (« j'aère à la profondeur des racines »).\n\n" +
    "Comment utiliser ce diaporama : lisez chaque diapositive, puis ces notes. Les notes contiennent les explications détaillées. À la fin, l'étude de cas ne donne pas la solution : vous envoyez votre raisonnement dans la conversation, et on l'analyse ensemble."
  );

  // ===== 2. Objectifs =====
  s = content("Ce que vous saurez faire", "Introduction");
  const objs = [
    ["cible", "Expliquer pourquoi un sable résiste mieux au piétinement qu’un limon"],
    ["racine", "Relier compaction, eau et oxygène à l’état des racines"],
    ["loupe", "Lire une carotte de sol et une analyse granulométrique"],
    ["regle", "Choisir une profondeur d’aération à partir d’un diagnostic"],
  ];
  objs.forEach(([ic, t], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.3 + Math.floor(i / 2) * 1.55;
    card(s, x, y, 4.4, 1.35);
    iconCircle(s, ic, x + 0.25, y + 0.33, 0.7, C.accent1);
    txt(s, [{ text: `Objectif ${i + 1}`, options: { bold: true, color: C.accent1, fontSize: 12, breakLine: true } }, { text: t, options: { fontSize: 15 } }], { x: x + 1.15, y: y + 0.2, w: 3.05, h: 1.0, valign: "middle" });
  });
  txt(s, "Validation : vous saurez appliquer ces 4 compétences à un terrain que vous ne connaissez pas.", { x: 0.5, y: 4.55, w: 9, h: 0.4, fontSize: 13, italic: true, color: C.text2 });
  s.addNotes(
    "Chaque objectif est mesurable : à la fin du module, l'évaluation vérifie que vous savez le faire, pas seulement le réciter.\n\n" +
    "Objectif 1 — vous devez pouvoir l'expliquer sans notes, avec le mécanisme (taille des grains, gros pores qui résistent ou non).\n" +
    "Objectif 2 — vous devez pouvoir citer les quatre besoins des racines (eau, oxygène, espace, nutriments) et dire ce que la compaction change pour chacun.\n" +
    "Objectif 3 — vous devez savoir quoi regarder dans une carotte (couches, couleur, odeur, racines) et lire les classes d'une analyse granulométrique.\n" +
    "Objectif 4 — vous devez partir d'un problème localisé dans le profil pour choisir l'outil et la profondeur, et non d'une règle fixe.\n\n" +
    "Rappel de la règle de validation : un sujet est validé au niveau 4 quand vous l'appliquez correctement à deux contextes différents, par exemple un terrain hybride à Marseille et un gazon naturel à Bordeaux."
  );

  // ===== 3. Pourquoi c'est important =====
  s = content("Pourquoi ce sujet compte", "Introduction");
  const why = [
    ["jeu", C.accent1, "Qualité de jeu", "Stabilité des appuis, dureté de la surface, rebond du ballon : tout dépend de l’état du sol."],
    ["pousse", C.accent1, "Durabilité", "Des racines profondes = un gazon qui récupère plus vite après l’usure."],
    ["cout", C.accent2, "Coûts", "Placage, rénovation, eau et engrais perdus : un sol mal compris coûte cher."],
    ["dispo", C.accent3, "Disponibilité", "Un terrain qui draine mal reste fermé après la pluie."],
  ];
  why.forEach(([ic, col, t, d], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.3, 2.1, 2.75);
    iconCircle(s, ic, x + 0.7, 1.5, 0.7, col);
    txt(s, t, { x: x + 0.15, y: 2.35, w: 1.8, h: 0.4, fontSize: 16, bold: true, align: "center" });
    txt(s, d, { x: x + 0.15, y: 2.8, w: 1.8, h: 1.15, fontSize: 13, color: C.text2 });
  });
  txt(s, "Ce que vous voyez en surface est souvent la conséquence de ce qui se passe dessous.", { x: 0.5, y: 4.35, w: 9, h: 0.5, fontSize: 17, italic: true, bold: true, color: C.accent1 });
  s.addNotes(
    "Le lien avec votre pratique : à la question 4 du diagnostic, vous cherchiez la cause d'une zone jaune en surface (maladie, insecte) avant de regarder le sol et l'eau. Ce module vous donne les bases pour regarder sous la surface en premier.\n\n" +
    "Qualité de jeu : un sol saturé d'eau perd sa stabilité ; un sol compacté et sec devient dur. Les deux modifient les appuis des joueurs et le comportement du ballon.\n\n" +
    "Durabilité : la récupération d'un gazon après usure dépend de ses réserves et de ses racines. Des racines bloquées dans les 4 premiers centimètres rendent le gazon fragile à la chaleur, à la sécheresse et à l'arrachement.\n\n" +
    "Coûts : un placage en cours de saison, une rénovation anticipée, de l'eau qui ruisselle au lieu de s'infiltrer, de l'azote perdu dans un sol saturé — tout cela a un coût direct.\n\n" +
    "Disponibilité : pour un staff technique, un terrain fermé est un terrain inutile. C'est un argument qui parle à l'entraîneur (lien avec la question 8)."
  );

  // ===== SECTION 1 =====
  s = sectionSlide("Partie 1", "Ce qui compose le sol", "1. Ce qui compose le sol");
  s.addNotes("Partie 1 : la texture (taille des grains) et la porosité (les vides entre les grains). C'est la base de tout le reste.");

  // ===== 4. Taille des particules =====
  s = content("Sable, limon, argile : une question de taille", "1. Ce qui compose le sol");
  // échelle : grain moyen 0,5 mm => ballon 22 cm (facteur 440)
  const scale = [
    { lab: "Sable", lim: "0,05 à 2 mm", ex: "Grain de 0,5 mm", obj: "= un ballon de football", d: 1.5, col: C.accent2 },
    { lab: "Limon", lim: "0,002 à 0,05 mm", ex: "Particule de 0,02 mm", obj: "= un petit pois", d: 0.32, col: C.accent6 },
    { lab: "Argile", lim: "moins de 0,002 mm", ex: "Particule de 0,002 mm", obj: "= un grain de semoule", d: 0.1, col: C.text2 },
  ];
  scale.forEach((p, i) => {
    const cx = 1.6 + i * 3.1;
    const cy = 2.2;
    s.addShape(pres.shapes.OVAL, { x: cx - p.d / 2, y: cy - p.d / 2 + 0.2, w: p.d, h: p.d, fill: { color: p.col }, line: { color: p.col, width: 0 }, objectName: nm("grain") });
    txt(s, p.lab, { x: cx - 1.4, y: 3.25, w: 2.8, h: 0.4, fontSize: 18, bold: true, align: "center" });
    txt(s, p.lim, { x: cx - 1.4, y: 3.65, w: 2.8, h: 0.3, fontSize: 14, color: C.text2, align: "center" });
    txt(s, [{ text: p.ex, options: { breakLine: true } }, { text: p.obj, options: { bold: true } }], { x: cx - 1.4, y: 4.0, w: 2.8, h: 0.6, fontSize: 13, color: C.text1, align: "center" });
  });
  txt(s, "Tout est agrandi 440 fois. Mêmes limites dans les classifications américaine (USDA) et française.", { x: 0.5, y: 4.75, w: 9, h: 0.35, fontSize: 11, italic: true, color: C.text2 });
  s.addNotes(
    "La texture d'un sol, c'est la proportion de sable, de limon et d'argile. Ces trois catégories se distinguent uniquement par la taille des particules :\n" +
    "• argile : moins de 2 micromètres (0,002 mm) ;\n" +
    "• limon : de 2 à 50 micromètres (0,002 à 0,05 mm) — la classification française distingue les limons fins (2 à 20 µm) et les limons grossiers (20 à 50 µm) ;\n" +
    "• sable : de 50 micromètres à 2 mm — la classification française distingue les sables fins (50 à 200 µm) et les sables grossiers (200 µm à 2 mm).\n\n" +
    "L'analogie : si on agrandit tout 440 fois, un grain de sable moyen de 0,5 mm devient un ballon de 22 cm, une particule de limon de 0,02 mm devient un pois de 9 mm, et une particule d'argile de 0,002 mm devient un grain de semoule de moins d'1 mm.\n\n" +
    "Pourquoi c'est important : entre des ballons, il reste de grands vides. Entre des pois et des grains de semoule tassés ensemble, les vides sont minuscules. Toute la suite du module découle de cette différence.\n\n" +
    "Retour sur votre réponse à la question 2 : « sablo-limoneux » désigne une texture qui contient une part importante de limon. Ce n'est pas le substrat attendu sur un terrain d'élite, construit sur un sable presque pur. D'où la mission : vérifier s'il existe une analyse granulométrique de votre substrat."
  );

  // ===== 5. Tous les sables ne se valent pas =====
  s = content("Tous les sables ne se valent pas", "1. Ce qui compose le sol");
  const segs = [
    { t: "Argile", s: "< 0,002", col: C.text2, w: 0.9 },
    { t: "Limon", s: "0,002–0,05", col: C.accent6, w: 1.1 },
    { t: "Sable très fin", s: "0,05–0,10", col: "E3C98F", w: 1.3 },
    { t: "Sable fin", s: "0,10–0,25", col: "D9B15E", w: 1.3 },
    { t: "Sable moyen", s: "0,25–0,50", col: C.accent2, w: 1.4 },
    { t: "Sable grossier", s: "0,50–1,0", col: "A87A1F", w: 1.4 },
    { t: "Très grossier", s: "1,0–2,0", col: "7D5B17", w: 1.6 },
  ];
  let sx = 0.5;
  const barY = 2.35;
  segs.forEach((g, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: sx, y: barY, w: g.w, h: 0.7, fill: { color: g.col }, line: { color: C.background1, width: 1.5 }, objectName: nm("classe") });
    const dark = i >= 4 || i <= 1;
    txt(s, g.t, { x: sx + 0.04, y: barY + 0.04, w: g.w - 0.08, h: 0.36, fontSize: 11, bold: true, align: "center", valign: "middle", color: dark ? C.background1 : C.text1 });
    txt(s, g.s + " mm", { x: sx + 0.04, y: barY + 0.38, w: g.w - 0.08, h: 0.28, fontSize: 10, align: "center", valign: "middle", color: dark ? C.background1 : C.text1 });
    g.x = sx;
    sx += g.w;
  });
  // repère cible : moyen + grossier
  const tx = segs[4].x, tw = segs[4].w + segs[5].w;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: tx, y: 1.3, w: tw, h: 0.85, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: nm("cible") });
  txt(s, [{ text: "Au moins 60 %", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "entre 0,25 et 1 mm", options: { fontSize: 12 } }], { x: tx + 0.05, y: 1.35, w: tw - 0.1, h: 0.75, color: C.background1, align: "center", valign: "middle" });
  const fx = segs[0].x, fw = segs[0].w + segs[1].w + segs[2].w;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: fx, y: 1.3, w: fw, h: 0.85, rectRadius: 0.08, fill: { color: C.accent5 }, line: { color: C.accent5, width: 0 }, objectName: nm("fines") });
  txt(s, [{ text: "Fines : 10 % maximum", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "dont limon ≤ 5 %, argile ≤ 3 %", options: { fontSize: 12 } }], { x: fx + 0.05, y: 1.35, w: fw - 0.1, h: 0.75, color: C.background1, align: "center", valign: "middle" });
  card(s, 0.5, 3.35, 9, 1.55);
  iconCircle(s, "alerte", 0.7, 3.62, 0.65, C.accent2);
  txt(s, [
    { text: "Attention au contexte de ces valeurs", options: { bold: true, breakLine: true } },
    { text: "Ce sont les recommandations de l’USGA pour les greens de golf. Elles servent souvent de repère pour les terrains de sport, mais ce n’est pas une norme football. Retenez surtout le principe : peu de particules fines, et des grains de taille assez homogène.", options: { fontSize: 13, color: C.text2 } },
  ], { x: 1.55, y: 3.5, w: 7.75, h: 1.3, fontSize: 14, valign: "middle" });
  s.addNotes(
    "Le sable lui-même se divise en cinq classes (classification USDA) : très fin (0,05 à 0,10 mm), fin (0,10 à 0,25), moyen (0,25 à 0,50), grossier (0,50 à 1,0) et très grossier (1,0 à 2,0 mm). Le schéma n'est pas à l'échelle : il sert à situer les classes les unes par rapport aux autres.\n\n" +
    "Les recommandations de l'USGA (United States Golf Association) pour la construction des greens sont le document de référence le plus cité au monde pour les substrats sableux. Les points clés : au moins 60 % des grains entre 0,25 et 1 mm (sable moyen et grossier), et une limite stricte sur les particules fines : 10 % au maximum de sable très fin + limon + argile, dont 5 % au maximum de limon et 3 % au maximum d'argile.\n\n" +
    "Pourquoi les fines sont limitées : quelques pourcents de particules fines suffisent à boucher les vides entre les gros grains (vous le verrez sur la diapositive 8).\n\n" +
    "Limite importante : ces valeurs ont été établies pour des greens de golf, pas pour des terrains de football. Les terrains de football utilisent des repères proches, mais chaque constructeur, chaque fabricant de système hybride et chaque pays peut avoir ses propres spécifications. Ne citez jamais ces chiffres comme « la norme » d'un terrain de football.\n\n" +
    "Ce qu'un responsable doit savoir faire : lire une courbe granulométrique de laboratoire, repérer la part de fines, et vérifier qu'un sable de sablage est compatible avec le substrat en place (diapositive 17)."
  );

  // ===== 6. Les pores =====
  s = content("Les pores : là où vivent l’eau et l’air", "1. Ce qui compose le sol");
  // schéma : boîte avec grains et films d'eau
  const bx = 0.5, by = 1.25, bw = 4.4, bh = 3.6;
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: C.accent4 }, line: { color: C.text2, width: 1 }, objectName: nm("cadre-pores") });
  const gd = 0.62;
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 7; c++) {
      const off = r % 2 ? gd * 0.5 : 0;
      const x = bx + 0.12 + c * gd + off, y = by + 0.1 + r * gd * 0.93;
      if (x + gd > bx + bw - 0.02) continue;
      // film d'eau (anneau bleu) autour du grain
      s.addShape(pres.shapes.OVAL, { x: x - 0.04, y: y - 0.04, w: gd * 0.92 + 0.08, h: gd * 0.92 + 0.08, fill: { color: C.accent3 }, line: { color: C.accent3, width: 0 }, objectName: nm("film") });
      s.addShape(pres.shapes.OVAL, { x, y, w: gd * 0.92, h: gd * 0.92, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: nm("grain") });
    }
  }
  // légende
  const leg = [
    [C.accent2, "Grain de sable"],
    [C.accent3, "Eau retenue autour des grains et dans les petits pores"],
    [C.accent4, "Air dans les gros pores"],
  ];
  leg.forEach(([col, t], i) => {
    s.addShape(pres.shapes.OVAL, { x: 5.25, y: 1.35 + i * 0.45, w: 0.28, h: 0.28, fill: { color: col }, line: { color: C.text2, width: 0.5 }, objectName: nm("legende") });
    txt(s, t, { x: 5.65, y: 1.32 + i * 0.45, w: 3.85, h: 0.36, fontSize: 13, valign: "middle" });
  });
  const stats = [
    ["35–55 %", "Porosité totale", C.text1],
    ["15–30 %", "Gros pores : air et drainage", C.accent4],
    ["15–25 %", "Petits pores : réserve d’eau", C.accent3],
  ];
  stats.forEach(([v, l, col], i) => {
    const x = 5.25 + i * 1.45;
    card(s, x, 2.85, 1.35, 1.6);
    txt(s, v, { x: x + 0.05, y: 2.95, w: 1.25, h: 0.55, fontSize: 20, bold: true, align: "center", valign: "middle", color: col === C.accent4 ? C.accent3 : col });
    txt(s, l, { x: x + 0.08, y: 3.55, w: 1.19, h: 0.8, fontSize: 11, align: "center", color: C.text2 });
  });
  txt(s, "Repères USGA pour un substrat de green, en % du volume du sol.", { x: 5.25, y: 4.55, w: 4.25, h: 0.3, fontSize: 10, italic: true, color: C.text2 });
  s.addNotes(
    "Un sol, ce n'est pas que des grains : c'est aussi des vides entre les grains, appelés pores. Dans un sable de terrain, ces vides représentent environ un tiers à la moitié du volume.\n\n" +
    "Il existe deux grandes familles de pores :\n" +
    "• les gros pores (macropores) : après une pluie, l'eau s'en écoule par gravité et ils se remplissent d'air. Ce sont eux qui assurent le drainage et l'oxygénation des racines ;\n" +
    "• les petits pores (micropores) : ils retiennent l'eau par capillarité, comme une éponge. C'est la réserve d'eau dans laquelle les racines puisent.\n\n" +
    "La limite de taille entre gros et petits pores varie selon les auteurs : retenez le principe (l'un se vide après la pluie, l'autre retient l'eau) plutôt qu'un chiffre.\n\n" +
    "Les repères USGA pour un substrat de green : porosité totale 35 à 55 %, porosité d'air 15 à 30 %, porosité capillaire 15 à 25 %. Ces valeurs sont mesurées en laboratoire sur un échantillon, à une succion donnée (méthode ASTM F1815). Elles ne se mesurent pas avec un outil de terrain.\n\n" +
    "L'équilibre est la clé : trop de gros pores, le sol sèche trop vite et retient mal les nutriments ; trop de petits pores, le sol reste gorgé d'eau et manque d'air. Le schéma est simplifié : dans un vrai sol, les grains ne sont ni ronds ni rangés."
  );

  // ===== 7. Pourquoi le sable résiste =====
  s = content("Pourquoi le sable résiste au piétinement", "1. Ce qui compose le sol");
  function panel(px, mixed) {
    const pw = 3.0, ph = 2.25, py = 1.3;
    s.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: mixed ? C.accent3 : C.accent4 }, line: { color: C.text2, width: 1 }, objectName: nm("cadre") });
    const d = 0.5;
    const rows = [];
    for (let r = 0; r < 5; r++) {
      const row = [];
      for (let c = 0; c < 6; c++) {
        const x = px + 0.08 + c * d + (r % 2 ? d * 0.5 : 0), y = py + 0.06 + r * d * 0.86;
        if (x + d > px + pw - 0.02 || y + d > py + ph) continue;
        row.push([x + d * 0.475, y + d * 0.475]);
        s.addShape(pres.shapes.OVAL, { x, y, w: d * 0.95, h: d * 0.95, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: nm("grain") });
      }
      rows.push(row);
    }
    if (mixed) {
      // particules fines logées dans les vides entre trois grains
      const fd = 0.13;
      rows.forEach((row, r) => {
        for (let i = 0; i + 1 < row.length; i++) {
          const mx = (row[i][0] + row[i + 1][0]) / 2;
          [1, -1].forEach((sg) => {
            const cy = row[i][1] + sg * d * 0.86 / 3;
            if ((sg === 1 && r + 1 < rows.length) || (sg === -1 && r > 0)) {
              if (cy - fd / 2 > py && cy + fd / 2 < py + ph) s.addShape(pres.shapes.OVAL, { x: mx - fd / 2, y: cy - fd / 2, w: fd, h: fd, fill: { color: C.accent6 }, line: { color: C.accent6, width: 0 }, objectName: nm("fine") });
            }
          });
        }
      });
    }
  }
  panel(0.5, false);
  panel(3.8, true);
  txt(s, [{ text: "Sable homogène", options: { bold: true, breakLine: true } }, { text: "Les grains s’appuient les uns sur les autres. Les gros pores tiennent sous la charge.", options: { fontSize: 13, color: C.text2 } }], { x: 0.5, y: 3.65, w: 3.0, h: 1.2, fontSize: 15 });
  txt(s, [{ text: "Grains de tailles variées", options: { bold: true, breakLine: true } }, { text: "Sous la charge, les fines comblent les vides. Il ne reste que de petits pores, pleins d’eau.", options: { fontSize: 13, color: C.text2 } }], { x: 3.8, y: 3.65, w: 3.0, h: 1.2, fontSize: 15 });
  card(s, 7.1, 1.3, 2.4, 3.55, C.text1);
  txt(s, [
    { text: "Mais le sable a ses limites", options: { bold: true, color: C.accent2, breakLine: true, fontSize: 15 } },
    { text: "Il retient peu d’eau.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Il retient peu les nutriments.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Il manque de cohésion : la surface peut être instable.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "D’où le rôle des racines… et des fibres hybrides.", options: { bold: true, color: C.accent4 } },
  ], { x: 7.25, y: 1.45, w: 2.1, h: 3.3, fontSize: 13, color: C.background1, paraSpaceAfter: 6 });
  s.addNotes(
    "C'est la réponse à la question 3 du diagnostic, restée sans réponse.\n\n" +
    "Pourquoi un sable homogène résiste : ses grains, de taille proche, se touchent comme des billes dans un seau. Sous le poids d'un joueur, ils se réarrangent un peu, mais il reste toujours de grands vides entre eux. Les gros pores survivent, le drainage et l'aération continuent.\n\n" +
    "Pourquoi un sol à granulométrie étalée (limon, ou sable chargé de fines) se compacte : sous la pression, les petites particules viennent se loger dans les vides laissés par les grosses. Le sol devient plus dense, les gros pores disparaissent, il ne reste que de petits pores qui retiennent l'eau. Dans un limon, les agrégats (les petites mottes qui créaient des vides) sont en plus écrasés par le piétinement.\n\n" +
    "Les limites du sable : il retient peu d'eau (d'où des arrosages plus fréquents, voir module M3), il retient mal les nutriments (voir modules M5 et M6) et il manque de cohésion. Un sable pur sans racines est instable sous les crampons.\n\n" +
    "Lien avec votre terrain : les fibres de votre système hybride servent en partie à compenser ce manque de stabilité, en ancrant la surface et les racines.\n\n" +
    "Nuance d'expert : un sable de terrain n'est pas immunisé contre la compaction. Il se tasse aussi, surtout en surface, quand la matière organique et les fines s'y accumulent (diapositive 16)."
  );

  // ===== SECTION 2 =====
  s = sectionSlide("Partie 2", "Compaction, eau et oxygène", "2. Compaction, eau, oxygène");
  s.addNotes("Partie 2 : ce que le piétinement fait au sol, et ce que le sol compacté fait aux racines. Cette partie corrige l'erreur E5 du diagnostic.");

  // ===== 8. L'humidité décide =====
  s = content("L’humidité décide de la compaction", "2. Compaction, eau, oxygène");
  s.addChart(pres.charts.LINE, [{ name: "Densité après compactage", labels: ["Très sec", "", "", "", "Optimum", "", "", "", "Détrempé"], values: [1.52, 1.6, 1.67, 1.72, 1.745, 1.72, 1.66, 1.58, 1.5] }], {
    x: 0.75, y: 1.2, w: 5.0, h: 3.35,
    chartColors: [HEX.accent6], lineSize: 3, lineDataSymbol: "none", lineSmooth: true,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: HEX.dk2, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt",
    showCatAxisTitle: true, catAxisTitle: "Humidité du sol au moment du passage →", catAxisTitleColor: HEX.dk2, catAxisTitleFontSize: 11,
    showLegend: false, showTitle: false, showValue: false,
  });
  txt(s, "Tassement obtenu →", { x: 0.0, y: 2.6, w: 1.4, h: 0.3, fontSize: 11, color: C.text2, rotate: 270, align: "center" });
  const zones = [
    [C.accent2, "Sol sec", "Les particules frottent entre elles : le sol résiste au tassement."],
    [C.accent5, "Sol humide", "L’eau aide les particules à glisser : c’est là que le sol se tasse le plus."],
    [C.accent3, "Sol détrempé", "L’eau occupe les pores : le sol se malaxe, perd sa structure, devient boue."],
  ];
  zones.forEach(([col, t, d], i) => {
    const y = 1.25 + i * 1.12;
    card(s, 6.0, y, 3.5, 1.0);
    s.addShape(pres.shapes.OVAL, { x: 6.15, y: y + 0.36, w: 0.28, h: 0.28, fill: { color: col }, line: { color: col, width: 0 }, objectName: nm("zone") });
    txt(s, [{ text: t, options: { bold: true, breakLine: true } }, { text: d, options: { fontSize: 12, color: C.text2 } }], { x: 6.55, y: y + 0.08, w: 2.85, h: 0.86, fontSize: 14, valign: "middle" });
  });
  txt(s, "Courbe schématique : la forme est typique, la position de l’optimum dépend du sol. Un sable propre y est beaucoup moins sensible qu’un limon.", { x: 0.5, y: 4.62, w: 9, h: 0.45, fontSize: 11, italic: true, color: C.text2 });
  s.addNotes(
    "Cette courbe s'appelle courbe Proctor, du nom de l'essai de laboratoire qui la mesure : on compacte le même sol avec la même énergie, à différentes humidités, et on mesure la densité obtenue.\n\n" +
    "Ce qu'elle montre : pour une même énergie de compactage (le même nombre de joueurs, la même machine), le tassement obtenu dépend de l'humidité du sol.\n" +
    "• Sol sec : les particules frottent entre elles et résistent au réarrangement.\n" +
    "• Sol humide, proche de l'optimum : l'eau lubrifie les contacts, les particules glissent et se tassent au maximum.\n" +
    "• Sol détrempé : les pores sont pleins d'eau, et l'eau ne se comprime pas. Le sol ne se tasse plus beaucoup, mais il se déforme et se malaxe : la structure est détruite et la surface devient boueuse. Les dégâts sont différents mais tout aussi graves.\n\n" +
    "Important : la courbe est schématique, les valeurs ne sont pas réelles. Un sable propre et homogène a une courbe beaucoup plus plate qu'un limon : il est peu sensible à l'humidité. C'est une autre raison de construire les terrains sur sable.\n\n" +
    "Application : passer une machine lourde ou faire jouer intensivement sur un sol humide chargé de fines (zone avec matière organique en surface, zone mal drainée) cause plus de tassement qu'à un autre moment. Le jour et l'état du sol au moment d'une intervention comptent autant que l'intervention elle-même."
  );

  // ===== 9. Compaction ≠ usure =====
  s = content("Compaction ou usure ?", "2. Compaction, eau, oxygène");
  const rowsCU = [
    ["", "Usure", "Compaction"],
    ["Où ?", "Au-dessus du sol", "Dans le sol"],
    ["Ce qui est abîmé", "Feuilles et collets arrachés, écrasés", "Pores écrasés, sol dense et asphyxié"],
    ["Ce qu’on voit", "Gazon clairsemé, terre à nu", "Eau qui stagne, sol dur, racines courtes"],
    ["Réponse", "Regarnir, protéger, faire tourner les exercices", "Aérer, décompacter, gérer l’eau"],
  ];
  s.addTable(rowsCU.map((r, i) => r.map((c, j) => ({
    text: c,
    options: {
      bold: i === 0 || j === 0,
      color: i === 0 ? C.background1 : (j === 0 ? C.text1 : C.text1),
      fill: { color: i === 0 ? (j === 1 ? C.accent1 : j === 2 ? C.accent6 : C.text1) : (j === 0 ? C.background2 : C.background1) },
      fontSize: 14, valign: "middle", margin: [4, 8, 4, 8],
    },
  }))), { x: 0.5, y: 1.25, w: 9.0, colW: [1.9, 3.55, 3.55], rowH: 0.5, border: { type: "solid", pt: 1, color: HEX.lt2 } });
  card(s, 0.5, 4.0, 9.0, 0.9, C.background2);
  iconCircle(s, "alerte", 0.65, 4.13, 0.62, C.accent5);
  txt(s, [{ text: "Devant les buts, les deux se cumulent. ", options: { bold: true } }, { text: "Regarnir une zone compactée sans traiter la compaction mène à une mauvaise reprise : les jeunes plants dépérissent faute d’oxygène et d’enracinement." }], { x: 1.45, y: 4.08, w: 7.9, h: 0.75, fontSize: 13, valign: "middle" });
  s.addNotes(
    "Cette distinction est au cœur de l'erreur E5 du diagnostic. À la question 3, vous parliez d'une plante « étranglée » sans distinguer ce qui se passe au-dessus et en dessous du sol.\n\n" +
    "L'usure est une agression mécanique sur la plante : les crampons arrachent, cisaillent et écrasent les feuilles et le collet (le point de croissance à la base de la plante). Le gazon récupère s'il a des réserves et si son collet est intact.\n\n" +
    "La compaction est une agression sur le sol : les gros pores sont écrasés, l'air disparaît, l'eau stagne, le sol durcit. La plante en souffre indirectement, par ses racines.\n\n" +
    "Les deux se traitent différemment. Un regarnissage règle un problème d'usure, pas un problème de compaction. Une aération règle un problème de compaction, pas un problème d'usure.\n\n" +
    "Question de raisonnement à vous poser sur le terrain : « Le gazon a-t-il disparu parce qu'il a été arraché, ou parce que le sol ne lui permet plus de vivre ? » Souvent, c'est les deux, mais dans des proportions différentes selon les zones."
  );

  // ===== 10. Les 4 besoins des racines =====
  s = content("Ce que la compaction fait aux racines", "2. Compaction, eau, oxygène");
  const needs = [
    ["eau", C.accent3, "Eau", "Elle s’infiltre moins vite et stagne en surface. Puis, en été, la même zone devient dure et sèche."],
    ["air", C.accent4, "Oxygène", "Les gros pores, pleins d’air, sont écrasés en premier. Les racines s’asphyxient."],
    ["regle", C.accent6, "Espace", "Le sol devient trop dur : la pointe des racines n’arrive plus à s’y enfoncer."],
    ["labo", C.accent1, "Nutriments", "Moins de racines et moins d’énergie pour absorber. En sol saturé, une partie de l’azote se perd en gaz."],
  ];
  needs.forEach(([ic, col, t, d], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.25 + Math.floor(i / 2) * 1.5;
    card(s, x, y, 4.4, 1.35);
    iconCircle(s, ic, x + 0.2, y + 0.33, 0.7, col);
    txt(s, [{ text: t, options: { bold: true, breakLine: true, fontSize: 16 } }, { text: d, options: { fontSize: 13, color: C.text2 } }], { x: x + 1.05, y: y + 0.12, w: 3.2, h: 1.1, valign: "middle" });
  });
  txt(s, [{ text: "Signe d’alerte : ", options: { bold: true, color: C.accent5 } }, { text: "le pâturin annuel (Poa annua) s’installe. Il supporte les sols tassés et humides mieux que vos espèces semées." }], { x: 0.5, y: 4.3, w: 9.0, h: 0.6, fontSize: 14, valign: "middle" });
  s.addNotes(
    "La méthode à retenir : face à un problème de sol, passez en revue les quatre besoins des racines (eau, oxygène, espace, nutriments) et demandez-vous ce que le problème change pour chacun.\n\n" +
    "Eau : contrairement à ce que vous pensiez (« l'absorption d'eau serait quasi nulle »), un sol compacté souffre souvent d'un excès d'eau après la pluie ou l'arrosage. L'eau entre moins vite mais reste piégée dans les petits pores. En été, la même zone sèche en surface et durcit. La plante peut donc subir l'excès puis le manque d'eau.\n\n" +
    "Oxygène : c'est le mécanisme principal (diapositive suivante).\n\n" +
    "Espace : quand la résistance du sol augmente, l'allongement des racines ralentit fortement (diapositive 14). Les racines restent en surface, et le gazon devient sensible à la sécheresse, à la chaleur et à l'arrachement.\n\n" +
    "Nutriments : absorber des nutriments demande de l'énergie, produite par la respiration des racines, qui a besoin d'oxygène. Et dans un sol saturé, des bactéries transforment une partie de l'azote nitrique en gaz qui s'échappe : c'est la dénitrification. Conséquence pratique : ajouter de l'engrais sur une zone compactée qui jaunit traite le symptôme, pas la cause.\n\n" +
    "Le pâturin annuel est un indicateur : sa présence en tache dans une zone de forte utilisation oriente vers une cause (compaction, humidité excessive), pas vers une solution (désherber)."
  );

  // ===== 11. Oxygène =====
  s = content("L’oxygène, le facteur invisible", "2. Compaction, eau, oxygène");
  card(s, 0.5, 1.25, 3.6, 3.6, C.text1);
  txt(s, "≈ 10 000 fois", { x: 0.6, y: 1.65, w: 3.4, h: 0.9, fontSize: 40, bold: true, color: C.accent4, align: "center", valign: "middle" });
  txt(s, "plus lent : la diffusion de l’oxygène dans l’eau, comparée à l’air", { x: 0.8, y: 2.6, w: 3.0, h: 0.9, fontSize: 15, color: C.background1, align: "center" });
  txt(s, "Un pore rempli d’eau est presque un mur pour l’oxygène.", { x: 0.8, y: 3.65, w: 3.0, h: 0.9, fontSize: 13, italic: true, color: C.accent2, align: "center" });
  const ox = [
    ["eau", C.accent3, "Sol saturé", "Il peut devenir pauvre en oxygène, voire sans oxygène, en une journée environ."],
    ["therm", C.accent5, "Chaleur : double peine", "Les racines respirent plus quand il fait chaud, alors que l’eau chaude contient moins d’oxygène dissous."],
    ["racine", C.accent1, "Gazon de climat tempéré", "Ses racines poussent le mieux entre 10 et 18 °C environ. Au-delà de 21 °C, leur croissance chute et elles meurent plus vite en sol humide."],
  ];
  ox.forEach(([ic, col, t, d], i) => {
    const y = 1.25 + i * 1.22;
    iconCircle(s, ic, 4.4, y + 0.18, 0.62, col);
    txt(s, [{ text: t, options: { bold: true, breakLine: true } }, { text: d, options: { fontSize: 12.5, color: C.text2 } }], { x: 5.2, y: y, w: 4.3, h: 1.1, fontSize: 15, valign: "middle" });
  });
  s.addNotes(
    "C'est le mécanisme que vous n'avez pas cité à la question 3, et c'est le plus important.\n\n" +
    "Les racines respirent : elles consomment de l'oxygène pour produire l'énergie dont elles ont besoin pour pousser et absorber l'eau et les nutriments. Cet oxygène vient de l'air contenu dans les gros pores du sol.\n\n" +
    "L'oxygène circule environ 10 000 fois moins vite dans l'eau que dans l'air. Quand les gros pores disparaissent (compaction) ou se remplissent d'eau (saturation), l'oxygène n'arrive presque plus jusqu'aux racines. Les sources consultées donnent toutes cet ordre de grandeur (10 000, soit 10⁴), avec de légères variations selon les conditions.\n\n" +
    "Un sol saturé peut devenir anaérobie (sans oxygène) en environ 24 heures selon un article technique de l'archive MSU. Ce délai dépend de la température et de l'activité biologique du sol.\n\n" +
    "En été à Marseille, c'est la double peine : la respiration des racines augmente avec la température, alors que la solubilité de l'oxygène dans l'eau diminue quand l'eau se réchauffe. Un sol à la fois chaud et gorgé d'eau est donc particulièrement dangereux pour les racines. Lien avec la question 6 : un arrosage de 10 minutes à 14 h n'est pas anodin sur une zone qui draine mal.\n\n" +
    "Repères pour les gazons de climat tempéré (comme le pâturin des prés et le ray-grass anglais) : croissance des racines optimale vers 10 à 18 °C (50 à 65 °F), en baisse rapide au-dessus d'environ 21 °C (70 °F), avec une mortalité racinaire accrue en sol humide et pauvre en oxygène (Université du Nebraska). Ce sont des ordres de grandeur, qui varient selon l'espèce et la variété."
  );

  // ===== 12. Résistance mécanique =====
  s = content("2 MPa : un repère, pas un seuil", "2. Compaction, eau, oxygène");
  // jauge
  const gx = 0.6, gy = 1.55, gw = 5.2;
  const bands = [[C.accent1, "Faible", 0.35], [C.accent2, "Croissance ralentie", 0.3], [C.accent5, "Fortement limitante", 0.35]];
  let bxp = gx;
  bands.forEach(([col, lab, f]) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bxp, y: gy, w: gw * f, h: 0.55, fill: { color: col }, line: { color: C.background1, width: 1.5 }, objectName: nm("jauge") });
    txt(s, lab, { x: bxp + 0.05, y: gy + 0.05, w: gw * f - 0.1, h: 0.45, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    bxp += gw * f;
  });
  const ticks = [[0, "0"], [0.35, "≈ 0,8"], [0.65, "≈ 2"], [1.0, "3 MPa et +"]];
  ticks.forEach(([f, l]) => txt(s, l, { x: gx + gw * f - 0.6, y: gy + 0.6, w: 1.2, h: 0.3, fontSize: 11, color: C.text2, align: "center" }));
  txt(s, "Résistance mesurée au pénétromètre (MPa)", { x: gx, y: gy - 0.38, w: gw, h: 0.3, fontSize: 12, bold: true, color: C.text2 });
  txt(s, [
    { text: "Ce que dit la recherche", options: { bold: true, breakLine: true, fontSize: 15 } },
    { text: "L’allongement des racines est souvent divisé par deux entre 0,8 et 2 MPa.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "2 MPa est un repère courant de forte limitation.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "La même zone lit plus « dur » quand elle sèche.", options: { bullet: { indent: 14 } } },
  ], { x: 0.6, y: 2.75, w: 5.2, h: 2.1, fontSize: 13, paraSpaceAfter: 6 });
  card(s, 6.2, 1.25, 3.3, 3.6, C.background2);
  iconCircle(s, "alerte", 7.5, 1.45, 0.7, C.accent5);
  txt(s, [
    { text: "Règle d’or", options: { bold: true, breakLine: true, fontSize: 16, color: C.accent5 } },
    { text: "Toute lecture de pénétromètre ou de Clegg se note avec l’humidité du sol au même endroit.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "Rappel E3 : le Clegg mesure la dureté de surface, pas la résistance du sol en profondeur.", options: { fontSize: 12, color: C.text2 } },
  ], { x: 6.4, y: 2.3, w: 2.9, h: 2.45, fontSize: 14, align: "left" });
  s.addNotes(
    "Le pénétromètre mesure la force nécessaire pour enfoncer une pointe dans le sol, exprimée en mégapascals (MPa). Il imite (grossièrement) l'effort que doit fournir la pointe d'une racine.\n\n" +
    "Ce que dit la recherche :\n" +
    "• Une revue scientifique de référence (Bengough et al., 2011, Journal of Experimental Botany) indique que l'allongement des racines est typiquement divisé par deux quand la résistance dépasse 0,8 à 2 MPa, en sol remanié et sans stress hydrique.\n" +
    "• Une valeur de 2 MPa est couramment citée comme limitante pour la croissance des racines. Certaines études citent une restriction nette au-delà de 2,5 MPa.\n\n" +
    "Pourquoi c'est un repère et non un seuil : la résistance dépend fortement de l'humidité. Le même sol, avec la même compaction, peut donner 1 MPa humide et plus de 2 MPa sec. En été, une lecture élevée peut donc signifier « sol sec » avant de signifier « sol compacté ». La même revue montre que, dans des sols de terrain, la part des mesures au-dessus de 2 MPa passe d'environ 10 % en sol humide à près de 50 % en sol plus sec.\n\n" +
    "La jauge est simplifiée : les limites de couleur sont indicatives et varient selon l'espèce et le sol.\n\n" +
    "Conséquence pour votre pratique : ne comparez jamais deux lectures prises à des humidités différentes. Mesurez la zone à problème et une zone témoin le même jour, à la même profondeur, et notez l'humidité de chacune.\n\n" +
    "Rappel de l'erreur E3, répétée pendant le diagnostic : le Clegg ne mesure pas la compaction en profondeur. Il mesure la décélération d'une masse qui tombe sur la surface (en Gmax). Sa valeur dépend de la compaction, mais aussi de l'humidité, du feutre, de la densité du gazon et des fibres."
  );

  // ===== SECTION 3 =====
  s = sectionSlide("Partie 3", "Comment l’eau circule", "3. Circulation de l'eau");
  s.addNotes("Partie 3 : ce qui empêche l'eau d'entrer et de circuler dans le profil. La surface, les couches, et le sable qui repousse l'eau.");

  // ===== 13. Surface commande =====
  s = content("Les premiers centimètres commandent", "3. Circulation de l'eau");
  // profil
  const px0 = 0.6, pw0 = 3.3;
  s.addShape(pres.shapes.RECTANGLE, { x: px0, y: 1.25, w: pw0, h: 0.35, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: nm("gazon") });
  s.addShape(pres.shapes.RECTANGLE, { x: px0, y: 1.6, w: pw0, h: 0.75, fill: { color: "3A2E22" }, line: { color: "3A2E22", width: 0 }, objectName: nm("couche-mo") });
  s.addShape(pres.shapes.RECTANGLE, { x: px0, y: 2.35, w: pw0, h: 2.45, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: nm("sable") });
  txt(s, "Gazon", { x: px0 + pw0 + 0.15, y: 1.25, w: 1.7, h: 0.35, fontSize: 12, valign: "middle" });
  txt(s, [{ text: "0 à 3 cm", options: { bold: true, breakLine: true } }, { text: "matière organique accumulée", options: { fontSize: 11, color: C.text2 } }], { x: px0 + pw0 + 0.15, y: 1.62, w: 1.9, h: 0.7, fontSize: 12, valign: "middle" });
  txt(s, [{ text: "Sable du substrat", options: { bold: true, breakLine: true } }, { text: "draine bien… si l’eau arrive jusqu’à lui", options: { fontSize: 11, color: C.text2 } }], { x: px0 + pw0 + 0.15, y: 2.6, w: 1.9, h: 0.9, fontSize: 12, valign: "middle" });
  // flèches eau
  [0.6, 1.5, 2.4].forEach((dx) => {
    s.addShape(pres.shapes.DOWN_ARROW, { x: px0 + dx, y: 0.98 + 0.1, w: 0.25, h: 0.42, fill: { color: C.accent3 }, line: { color: C.accent3, width: 0 }, objectName: nm("pluie") });
  });
  txt(s, [
    { text: "Ce qui se passe", options: { bold: true, breakLine: true, fontSize: 15 } },
    { text: "La matière organique (racines mortes, feutre) s’accumule en surface.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Elle retient l’eau comme une éponge et freine l’infiltration.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Elle favorise aussi le sable qui repousse l’eau et certaines maladies.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Le sablage régulier sert d’abord à la diluer.", options: { bullet: { indent: 14 } } },
  ], { x: 6.15, y: 1.25, w: 3.35, h: 2.6, fontSize: 13, paraSpaceAfter: 5 });
  card(s, 6.15, 3.95, 3.35, 0.9, C.background2);
  txt(s, [{ text: "À construction (repère USGA) : ", options: { bold: true } }, { text: "drainage de 15 à 30 cm par heure en laboratoire, 1 à 5 % de matière organique." }], { x: 6.3, y: 4.0, w: 3.1, h: 0.8, fontSize: 11.5, valign: "middle" });
  s.addNotes(
    "Message clé : un substrat sableux bien construit peut drainer très vite en profondeur, mais si l'eau ne traverse pas les premiers centimètres, cela ne sert à rien.\n\n" +
    "Avec les années, la matière organique s'accumule en surface : racines mortes, tiges, feutre. Cette couche se comporte comme une éponge : elle retient l'eau près de la surface, réduit l'infiltration, garde la surface molle et humide, et crée un milieu favorable à certaines maladies. Elle peut aussi rendre le sable juste en dessous hydrophobe (diapositive 18).\n\n" +
    "Le rôle premier du sablage régulier est de diluer cette matière organique dans les premiers centimètres (Université du Nebraska). L'aération à dents creuses retire une partie de la couche, et le sablage remplit les trous avec du sable propre.\n\n" +
    "Repères USGA, à utiliser avec prudence : un substrat de green neuf doit drainer de 15 à 30 cm par heure en laboratoire (conductivité hydraulique saturée, gamme « normale ») et contenir 1 à 5 % de matière organique en masse (idéalement 2 à 4 %). Ce sont des spécifications de construction, pas des seuils de suivi d'un terrain en service.\n\n" +
    "Pour suivre la matière organique : on mesure le taux de matière organique par perte au feu, en laboratoire, par tranches de profondeur (par exemple 0–2 cm et 2–4 cm). Toujours préciser la profondeur de l'échantillon, sinon le chiffre ne veut rien dire.\n\n" +
    "Lien avec votre pratique : quand l'eau stagne après une averse, commencez par regarder les premiers centimètres (carotte) avant de soupçonner les drains."
  );

  // ===== 14. Couches =====
  s = content("Deux couches différentes = de l’eau bloquée", "3. Circulation de l'eau");
  const lx = 0.6, lw = 3.2;
  s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 1.3, w: lw, h: 0.3, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: nm("gazon") });
  s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 1.6, w: lw, h: 1.1, fill: { color: "E3C98F" }, line: { color: "E3C98F", width: 0 }, objectName: nm("couche-fine") });
  s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 2.25, w: lw, h: 0.45, fill: { color: C.accent3 }, transparency: 35, line: { color: C.accent3, width: 0 }, objectName: nm("eau-perchee") });
  s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 2.7, w: lw, h: 2.1, fill: { color: "A87A1F" }, line: { color: "A87A1F", width: 0 }, objectName: nm("couche-grossiere") });
  txt(s, [{ text: "Sable plus fin", options: { bold: true, breakLine: true } }, { text: "apporté au sablage", options: { fontSize: 11, color: C.text2 } }], { x: lx + lw + 0.15, y: 1.6, w: 1.9, h: 0.6, fontSize: 12, valign: "middle" });
  txt(s, [{ text: "Eau bloquée", options: { bold: true, color: C.accent3, breakLine: true } }, { text: "au-dessus de l’interface", options: { fontSize: 11, color: C.text2 } }], { x: lx + lw + 0.15, y: 2.2, w: 1.9, h: 0.55, fontSize: 12, valign: "middle" });
  txt(s, [{ text: "Sable du substrat", options: { bold: true, breakLine: true } }, { text: "plus grossier", options: { fontSize: 11, color: C.text2 } }], { x: lx + lw + 0.15, y: 3.2, w: 1.9, h: 0.6, fontSize: 12, valign: "middle" });
  txt(s, [
    { text: "Le principe", options: { bold: true, breakLine: true, fontSize: 15 } },
    { text: "L’eau ne passe dans la couche du dessous qu’une fois la couche du dessus saturée.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "Les conséquences", options: { bold: true, breakLine: true, fontSize: 15 } },
    { text: "Excès d’eau près de la surface", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Racines qui s’arrêtent à l’interface", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Gazon qui se décolle à ce niveau", options: { bullet: { indent: 14 } } },
  ], { x: 6.1, y: 1.3, w: 3.4, h: 2.6, fontSize: 13 });
  card(s, 6.1, 3.95, 3.4, 0.9, C.text1);
  txt(s, "Tout nouveau sable de sablage se vérifie en laboratoire avant usage.", { x: 6.25, y: 4.0, w: 3.1, h: 0.8, fontSize: 13, bold: true, color: C.accent2, valign: "middle" });
  s.addNotes(
    "C'est l'un des principes les plus contre-intuitifs de la physique des sols : quand une couche de texture fine repose sur une couche plus grossière, l'eau ne descend pas librement. Elle reste retenue dans la couche fine jusqu'à ce que celle-ci soit saturée, puis seulement elle passe dans la couche du dessous. On parle de nappe perchée.\n\n" +
    "Pourquoi : les petits pores de la couche fine retiennent l'eau plus fort que les gros pores de la couche grossière ne l'attirent. L'eau « préfère » rester là où elle est.\n\n" +
    "Conséquences : la zone au-dessus de l'interface reste gorgée d'eau (donc pauvre en oxygène), les racines ont tendance à ne pas traverser l'interface, et le gazon peut se décoller à ce niveau. Même quelques centimètres d'un matériau différent suffisent à créer ce problème.\n\n" +
    "Cause fréquente sur un terrain en service : un sablage avec un sable différent du substrat en place (plus fin, ou plus chargé en fines), par exemple quand le fournisseur habituel est en rupture de stock. L'USGA Green Section insiste sur l'utilisation d'un sable de sablage compatible avec le substrat, vérifié par analyse de laboratoire.\n\n" +
    "Nuance d'expert : la nappe perchée n'est pas toujours un défaut. Les greens construits selon la méthode USGA placent volontairement le substrat sableux sur une couche de gravier, pour garder une réserve d'eau dans le bas du substrat. Le problème vient des couches non prévues, près de la surface."
  );

  // ===== 15. Sable hydrophobe =====
  s = content("Le sable qui repousse l’eau", "3. Circulation de l'eau");
  txt(s, [
    { text: "Ce que vous voyez", options: { bold: true, breakLine: true, fontSize: 15 } },
    { text: "Des zones sèches irrégulières qui jaunissent malgré un arrosage normal, surtout en été.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "La cause", options: { bold: true, breakLine: true, fontSize: 15 } },
    { text: "Les grains se couvrent de composés organiques (matière organique, champignons) qui repoussent l’eau. C’est souvent juste sous le feutre.", options: {} },
  ], { x: 0.5, y: 1.25, w: 4.3, h: 2.6, fontSize: 13 });
  txt(s, "Test de la goutte d’eau sur une carotte", { x: 5.1, y: 1.25, w: 4.4, h: 0.35, fontSize: 15, bold: true });
  const wd = [
    ["< 5 s", "Mouillable", C.accent1],
    ["5–60 s", "Légèrement hydrophobe", "D9B15E"],
    ["60–600 s", "Fortement hydrophobe", C.accent2],
    ["600–3 600 s", "Sévèrement hydrophobe", "C86A2E"],
    ["> 3 600 s", "Extrêmement hydrophobe", C.accent5],
  ];
  wd.forEach(([t, l, col], i) => {
    const y = 1.7 + i * 0.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.1, y, w: 1.35, h: 0.42, rectRadius: 0.06, fill: { color: col }, line: { color: col, width: 0 }, objectName: nm("classe-wdpt") });
    txt(s, t, { x: 5.1, y, w: 1.35, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, l, { x: 6.6, y, w: 2.9, h: 0.42, fontSize: 13, valign: "middle" });
  });
  txt(s, "Classes de Dekker et Ritsema (1999), temps de pénétration d’une goutte.", { x: 5.1, y: 4.25, w: 4.4, h: 0.3, fontSize: 10, italic: true, color: C.text2 });
  card(s, 0.5, 3.95, 4.3, 0.9, C.background2);
  txt(s, [{ text: "Toujours comparer : ", options: { bold: true } }, { text: "zone sèche et zone saine, même profondeur, même jour." }], { x: 0.65, y: 4.0, w: 4.0, h: 0.8, fontSize: 13, valign: "middle" });
  s.addNotes(
    "C'est l'une des hypothèses qui manquaient dans votre réponse à la question 4 (zone qui jaunit fin juin à Marseille).\n\n" +
    "Les grains de sable peuvent se couvrir d'un film de composés organiques, issus de la décomposition de la matière organique ou produits par certains champignons du sol. Ce film repousse l'eau : l'arrosage ruisselle ou reste en surface, et le sol en dessous reste sec. On parle de zones sèches localisées (en anglais, localized dry spot).\n\n" +
    "Où la chercher : l'hydrophobie est généralement concentrée près de la surface. Une étude de l'Université de Californie sur des greens a même trouvé le feutre mouillable, et le sable juste en dessous fortement hydrophobe.\n\n" +
    "Le test de la goutte d'eau (WDPT, pour water drop penetration time) : on prélève une carotte, on dépose une goutte d'eau tous les centimètres le long du profil, et on chronomètre le temps qu'elle met à pénétrer. La classification de Dekker et Ritsema (1999) est la référence en recherche : moins de 5 s mouillable, 5 à 60 s légèrement hydrophobe, 60 à 600 s fortement, 600 à 3 600 s sévèrement, plus de 3 600 s extrêmement hydrophobe. Certains guides pratiques de golf considèrent qu'au-delà de 10 à 15 secondes, il y a un problème.\n\n" +
    "Attention : le résultat dépend de l'humidité de la carotte (un échantillon sec repousse plus l'eau). Comparez toujours avec une carotte prise dans une zone saine voisine, le même jour.\n\n" +
    "Les agents mouillants (surfactants) sont l'outil de gestion habituel. On les verra dans le module M3 (arrosage)."
  );

  // ===== 16. Hybride =====
  s = content("Ce que change un terrain hybride", "3. Circulation de l'eau");
  const hy = [
    ["couches", C.accent1, "Fibres plantées", "Des fibres synthétiques sont cousues dans le substrat sableux, selon une grille de l’ordre du centimètre."],
    ["regle", C.accent2, "Profondeur variable", "De 4,5 à 18 cm selon le système et le choix du club. C’est l’information clé pour régler l’aération."],
    ["racine", C.accent1, "Stabilité", "Les racines s’ancrent autour des fibres : la surface résiste mieux au cisaillement des crampons."],
    ["outils", C.accent5, "Contraintes", "Outils et profondeurs d’aération à valider avec les consignes du fabricant de votre système."],
  ];
  hy.forEach(([ic, col, t, d], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.25 + Math.floor(i / 2) * 1.5;
    card(s, x, y, 4.4, 1.35);
    iconCircle(s, ic, x + 0.2, y + 0.33, 0.7, col);
    txt(s, [{ text: t, options: { bold: true, breakLine: true, fontSize: 16 } }, { text: d, options: { fontSize: 13, color: C.text2 } }], { x: x + 1.05, y: y + 0.12, w: 3.2, h: 1.1, valign: "middle" });
  });
  txt(s, [{ text: "Exemples : ", options: { bold: true } }, { text: "GrassMaster, fibres de 20 cm dont 18 cm dans le sol, grille d’environ 2 × 2 cm. SIS, grilles de 20 × 10 à 20 × 30 mm, de 4,5 à 18 cm de profondeur." }], { x: 0.5, y: 4.3, w: 9, h: 0.6, fontSize: 12, color: C.text2, valign: "middle" });
  s.addNotes(
    "Les systèmes hybrides cousus (comme GrassMaster ou SISGrass) consistent à injecter des fibres synthétiques dans un substrat sableux. Le gazon naturel pousse autour, et ses racines s'entremêlent aux fibres.\n\n" +
    "Données fabricants (à vérifier pour votre propre système) :\n" +
    "• GrassMaster : fibres de 20 cm de long, injectées à 18 cm de profondeur, 2 cm restant en surface, selon une grille d'environ 2 × 2 cm (source secondaire : Wikipédia, chiffres non sourcés par le fabricant).\n" +
    "• SISGrass : espacements proposés de 20 × 10 mm à 20 × 30 mm, profondeur de couture de 45 à 180 mm selon le sport, l'utilisation, le sol et le budget (source : SIS Pitches, fabricant).\n\n" +
    "Correction de la question 2 : vous aviez dit « une fibre tous les 1 mm ». C'est impossible (cela ferait un million de fibres par m²). L'ordre de grandeur est le centimètre.\n\n" +
    "Ce que les fibres changent : elles améliorent la stabilité de la surface et sa résistance à l'arrachement, et elles compensent en partie le manque de cohésion du sable. Elles ne changent pas la physique du sable : la compaction, l'accumulation de matière organique, les couches et l'hydrophobie existent aussi sur un terrain hybride.\n\n" +
    "Ce qu'elles imposent : la profondeur et le type d'outil d'aération doivent être compatibles avec les fibres. Les consignes du fabricant de votre système font foi. C'est pour cela que la mission « nom du système et profondeur des fibres » est prioritaire."
  );

  // ===== SECTION 4 =====
  s = sectionSlide("Partie 4", "Mesurer et appliquer", "4. Mesurer et appliquer");
  s.addNotes("Partie 4 : comment mesurer ce que vous venez d'apprendre, et comment l'appliquer sur le terrain.");

  // ===== 17. Mesures labo =====
  const tblOpts = (rows, colW) => rows.map((r, i) => r.map((c, j) => ({
    text: c,
    options: { bold: i === 0 || j === 0, color: i === 0 ? C.background1 : C.text1, fill: { color: i === 0 ? C.text1 : (i % 2 ? C.background1 : C.background2) }, fontSize: 12, valign: "middle", margin: [3, 6, 3, 6] },
  })));
  s = content("Données et indicateurs : au laboratoire", "4. Mesurer et appliquer");
  const lab = [
    ["Mesure", "Ce qu’elle vous dit", "Unité", "Piège courant"],
    ["Granulométrie", "Répartition des grains par classe de taille", "% par classe", "Un seul échantillon ne représente pas tout le terrain"],
    ["Porosité d’air et capillaire", "Équilibre entre air et réserve d’eau", "% du volume", "Mesure sur carotte non remaniée, sinon résultat faux"],
    ["Conductivité hydraulique", "Vitesse à laquelle l’eau traverse le substrat", "cm/h", "Le labo mesure le substrat, pas la surface feutrée"],
    ["Matière organique (perte au feu)", "Accumulation organique par tranche", "% de la masse", "Sans la profondeur de l’échantillon, le chiffre ne veut rien dire"],
  ];
  s.addTable(tblOpts(lab), { x: 0.5, y: 1.25, w: 9.0, colW: [2.1, 2.9, 1.2, 2.8], rowH: 0.55, border: { type: "solid", pt: 1, color: HEX.lt2 } });
  txt(s, "Ces analyses servent à comprendre un terrain et à suivre son évolution d’une année à l’autre. Un résultat isolé ne suffit pas.", { x: 0.5, y: 4.45, w: 9, h: 0.45, fontSize: 13, italic: true, color: C.text2 });
  s.addNotes(
    "Les analyses de laboratoire répondent à la question : « De quoi est fait mon sol, et comment se comporte-t-il ? »\n\n" +
    "Granulométrie : elle donne le pourcentage de grains dans chaque classe de taille. C'est la base pour savoir si le substrat est conforme, et pour vérifier qu'un sable de sablage est compatible avec le substrat en place.\n\n" +
    "Porosité et conductivité hydraulique : elles se mesurent sur des carottes non remaniées (prélevées sans écraser la structure) selon des méthodes normalisées (par exemple ASTM F1815). Une carotte écrasée ou tassée au prélèvement donne un résultat faux.\n\n" +
    "Matière organique : mesurée par perte au feu (on brûle l'échantillon et on mesure la masse perdue). Toujours par tranches (0–2 cm, 2–4 cm…) : la matière organique est concentrée en surface, et un échantillon de 0–10 cm « dilue » le problème.\n\n" +
    "Erreur d'analyse courante : prendre un seul échantillon au centre du terrain et conclure sur tout le terrain. Les zones de forte utilisation (surfaces de réparation, couloirs d'échauffement) évoluent différemment.\n\n" +
    "Erreur de raisonnement courante : comparer des résultats de laboratoires différents sans vérifier que les méthodes sont les mêmes."
  );

  // ===== 18. Mesures terrain =====
  s = content("Données et indicateurs : sur le terrain", "4. Mesurer et appliquer");
  const field = [
    ["Outil", "Ce qu’il mesure vraiment", "Unité", "Piège courant"],
    ["Sonde d’humidité", "Eau dans le sol, sur la longueur des tiges", "% du volume", "La profondeur lue dépend des tiges utilisées"],
    ["Pénétromètre", "Résistance du sol en profondeur", "MPa", "Augmente quand le sol sèche"],
    ["Clegg", "Dureté de la surface", "Gmax", "Ce n’est pas la compaction ; dépend de la masse du marteau"],
    ["Infiltromètre (anneau)", "Vitesse d’entrée de l’eau par la surface", "mm/h", "Dépend de l’humidité de départ"],
    ["Test de la goutte", "Hydrophobie du sable", "secondes", "Dépend de l’humidité de la carotte"],
  ];
  s.addTable(tblOpts(field), { x: 0.5, y: 1.25, w: 9.0, colW: [2.0, 3.0, 1.2, 2.8], rowH: 0.5, border: { type: "solid", pt: 1, color: HEX.lt2 } });
  card(s, 0.5, 4.35, 9, 0.6, C.text1);
  txt(s, "Une mesure sans humidité, sans unité et sans zone témoin ne prouve rien.", { x: 0.65, y: 4.38, w: 8.7, h: 0.54, fontSize: 15, bold: true, color: C.accent2, valign: "middle", align: "center" });
  s.addNotes(
    "Les outils de terrain répondent à la question : « Comment se comporte mon terrain aujourd'hui, à cet endroit ? »\n\n" +
    "Sonde d'humidité : elle mesure la part d'eau dans le volume de sol traversé par ses tiges. Avec des tiges de 4 cm, vous lisez l'humidité des 4 premiers centimètres ; avec des tiges de 12 cm, une moyenne sur 12 cm. Deux lectures prises avec des tiges différentes ne sont pas comparables. Mission : vérifier la longueur des tiges de votre sonde.\n\n" +
    "Pénétromètre : voir la diapositive 14. Toujours noter l'humidité au même point.\n\n" +
    "Clegg : il mesure la dureté de surface en Gmax. Les résultats changent selon la masse du marteau (0,5 kg ou 2,25 kg). Le seuil de 100 Gmax que vous citiez correspond à la limite de sécurité imposée par la NFL avec le marteau de 2,25 kg : c'est un seuil de sécurité des joueurs, pas un déclencheur d'aération.\n\n" +
    "Infiltromètre : on enfonce un anneau dans le sol, on le remplit d'eau et on mesure la vitesse à laquelle le niveau baisse. Il mesure ce qui se passe réellement en surface, feutre compris. Le résultat dépend de l'humidité du sol au départ : un sol déjà humide absorbe moins vite.\n\n" +
    "Le NDVI n'apparaît pas ici : il mesure l'état du gazon, pas celui du sol. Il sert à repérer et délimiter un problème, pas à en trouver la cause.\n\n" +
    "La règle de la zone témoin : chaque mesure dans une zone à problème se compare à la même mesure, le même jour, dans une zone saine voisine. Sans témoin, vous ne savez pas si la valeur est anormale."
  );

  // ===== 19. Application : carotte =====
  s = content("Application : lire une carotte de sol", "4. Mesurer et appliquer");
  const steps = [
    ["1", "Choisir 3 points", "Zone très utilisée, zone moyenne, zone témoin saine."],
    ["2", "Prélever", "Carotte de 15 à 20 cm, sans écraser le profil. Accord du responsable."],
    ["3", "Observer", "Couches, couleur, odeur d’œuf pourri, profondeur et couleur des racines."],
    ["4", "Mesurer au même point", "Humidité, Clegg, goutte d’eau. Noter l’heure."],
    ["5", "Noter et photographier", "Photo avec une règle. Fiche datée."],
  ];
  steps.forEach(([n, t, d], i) => {
    const x = 0.5 + i * 1.84;
    card(s, x, 1.3, 1.7, 2.75);
    s.addShape(pres.shapes.OVAL, { x: x + 0.5, y: 1.45, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: nm("etape") });
    txt(s, n, { x: x + 0.5, y: 1.45, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, t, { x: x + 0.1, y: 2.25, w: 1.5, h: 0.55, fontSize: 14, bold: true, align: "center", valign: "middle" });
    txt(s, d, { x: x + 0.12, y: 2.85, w: 1.46, h: 1.15, fontSize: 12, color: C.text2, align: "center" });
  });
  txt(s, [{ text: "Ce qui doit vous alerter : ", options: { bold: true, color: C.accent5 } }, { text: "une couche sombre et dense en surface, des racines qui s’arrêtent net, une interface entre deux sables, une odeur d’œuf pourri (sol sans oxygène)." }], { x: 0.5, y: 4.2, w: 9, h: 0.7, fontSize: 13, valign: "middle" });
  s.addNotes(
    "La carotte de sol est l'outil de diagnostic le plus simple et le plus riche. Elle montre directement le profil : ce que vous ne voyez pas depuis la surface.\n\n" +
    "Étape 1 — Toujours trois points minimum : la zone à problème, une zone intermédiaire et une zone témoin saine. C'est la comparaison qui donne du sens.\n\n" +
    "Étape 2 — Prélever 15 à 20 cm avec un carottier ou une bêche, en évitant d'écraser le profil. Sur un terrain hybride, demandez l'accord de votre responsable et le type d'outil adapté, à cause des fibres.\n\n" +
    "Étape 3 — Ce qu'on regarde :\n" +
    "• les couches : une couche sombre et dense de matière organique en surface, une interface nette entre deux sables différents ;\n" +
    "• la couleur : une couche noire avec une odeur d'œuf pourri indique un sol privé d'oxygène (conditions anaérobies) ;\n" +
    "• les racines : jusqu'où descendent-elles ? S'arrêtent-elles net à une profondeur précise (couche, interface, zone dure) ? Sont-elles blanches et saines, ou brunes et molles ?\n\n" +
    "Étape 4 — Mesurer au même point et au même moment : humidité, Clegg, test de la goutte. Sinon, impossible de relier ce que vous voyez à ce que vous mesurez.\n\n" +
    "Étape 5 — Photo avec une règle graduée posée le long de la carotte, et une fiche datée. Refaire la même série au même endroit à chaque saison : c'est ainsi qu'on suit l'évolution d'un terrain.\n\n" +
    "C'est votre mission terrain de ce module (voir la dernière diapositive)."
  );

  // ===== 20. Application : profondeur d'aération =====
  s = content("Application : choisir une profondeur d’aération", "4. Mesurer et appliquer");
  const aer = [
    ["Problème observé", "Où ?", "Objectif", "Approche"],
    ["Matière organique accumulée", "0 à 3 cm", "Retirer et diluer", "Dents creuses peu profondes + sablage"],
    ["Surface tassée, eau qui stagne", "Premiers cm", "Rouvrir des gros pores", "Dents juste au-delà de la zone tassée"],
    ["Couche ou interface dense", "À x cm", "La traverser", "Dents qui descendent sous la couche"],
    ["Semelle sous l’aération habituelle", "Sous la profondeur habituelle", "La casser", "Décompactage plus profond, occasionnel"],
  ];
  s.addTable(tblOpts(aer), { x: 0.5, y: 1.25, w: 9.0, colW: [2.6, 1.6, 1.9, 2.9], rowH: 0.52, border: { type: "solid", pt: 1, color: HEX.lt2 } });
  card(s, 0.5, 3.95, 9.0, 0.95, C.background2);
  iconCircle(s, "check", 0.65, 4.1, 0.65, C.accent1);
  txt(s, [{ text: "Toujours vérifier : ", options: { bold: true } }, { text: "profondeur des fibres et consignes du fabricant · humidité du sol (ni sec, ni détrempé) · calendrier des séances · varier les profondeurs au fil de la saison." }], { x: 1.5, y: 4.0, w: 7.85, h: 0.85, fontSize: 13, valign: "middle" });
  s.addNotes(
    "C'est la correction de l'erreur E1 du diagnostic : « j'aère à 13 cm parce que c'est la longueur des racines ».\n\n" +
    "Le raisonnement attendu part toujours du problème : Quel problème ? → Où se situe-t-il dans le profil (la carotte vous le dit) ? → Quel objectif ? → Quel outil ? → Quelle profondeur et quel moment ? → Comment vérifier que ça a marché ?\n\n" +
    "Le tableau ne donne volontairement pas de profondeurs en centimètres pour chaque ligne : elles dépendent de votre profil, de votre système hybride et de vos machines. La « bonne » profondeur, c'est celle qui atteint la zone à problème.\n\n" +
    "Sur la semelle : aérer toujours à la même profondeur peut créer une couche tassée au fond des trous. Des travaux de l'Université d'État du Michigan ont montré que les dents compactent les parois et le fond de chaque trou. Les dents pleines y seraient plus sujettes. Le remède classique est de faire varier la profondeur et d'intégrer de temps en temps un passage plus profond.\n\n" +
    "Les contraintes : la profondeur des fibres et les consignes du fabricant de votre système ; l'humidité du sol (un sol trop sec résiste et casse les dents ou soulève la surface, un sol détrempé se lisse et se malaxe autour des dents) ; le calendrier des séances (temps de récupération de la surface).\n\n" +
    "Et vérifier que ça a marché : refaire les mêmes mesures (infiltration, humidité, carotte) quelques semaines après, aux mêmes points."
  );

  // ===== 21. Étude de cas — situation =====
  s = content("Étude de cas : l’eau ne part plus", "4. Mesurer et appliquer");
  const caseL = [
    ["Situation", "Terrain hybride n°2, Marseille, fin octobre. Orage de 30 mm en une heure."],
    ["Symptôme", "L’eau stagne 6 à 8 h dans les deux surfaces de réparation et sur une bande de 5 m le long de la touche (échauffement des remplaçants). Ailleurs : absorbée en moins de 30 min."],
    ["Carottes", "Surface de réparation : couche sombre et dense de 0 à 3 cm, racines arrêtées vers 4 cm, sable clair dessous. Centre : racines jusqu’à 10 cm."],
  ];
  caseL.forEach(([t, d], i) => {
    const y = 1.25 + i * 1.2;
    card(s, 0.5, y, 5.4, 1.1);
    txt(s, [{ text: t, options: { bold: true, color: C.accent1, breakLine: true } }, { text: d, options: { fontSize: 12.5 } }], { x: 0.65, y: y + 0.06, w: 5.1, h: 0.98, fontSize: 13, valign: "middle" });
  });
  card(s, 6.1, 1.25, 3.4, 3.5, C.text1);
  txt(s, [
    { text: "Autres informations", options: { bold: true, color: C.accent2, breakLine: true, fontSize: 14 } },
    { text: "Clegg 2,25 kg, le lendemain : surface 92, centre 68 Gmax.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Sablage de juin fait avec le sable d’un nouveau fournisseur, sans analyse.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Aération mensuelle à 8 cm depuis avril.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Deux séances par jour. Trêve internationale dans 3 semaines : 5 jours sans séance.", options: { bullet: { indent: 14 }, breakLine: true } },
    { text: "Budget d’intervention limité.", options: { bullet: { indent: 14 } } },
  ], { x: 6.25, y: 1.38, w: 3.1, h: 3.3, fontSize: 12, color: C.background1, paraSpaceAfter: 4 });
  txt(s, "Cas fictif construit pour l’exercice. Les informations sont volontairement incomplètes.", { x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 10, italic: true, color: C.text2 });
  s.addNotes(
    "Lisez attentivement la situation. Elle est construite pour ressembler à un vrai problème : plusieurs causes possibles, des informations partielles, des contraintes de calendrier et de budget.\n\n" +
    "Ne cherchez pas « la » bonne réponse tout de suite. Le but est de montrer votre raisonnement.\n\n" +
    "Pas de solution dans ce diaporama : vous envoyez votre réponse dans la conversation, je l'analyse, puis on approfondit.\n\n" +
    "Conseil de méthode : relisez les diapositives 11, 12, 16, 17 et 24 avant de répondre. Et posez-vous la question : qu'est-ce qui distingue les zones où l'eau stagne des zones où elle part ?"
  );

  // ===== 22. Étude de cas — questions =====
  s = content("Étude de cas : votre raisonnement", "4. Mesurer et appliquer");
  const qs = [
    "Séparez les symptômes des causes possibles.",
    "Formulez au moins quatre hypothèses.",
    "Quelles informations vous manquent ?",
    "Quelles mesures et observations faites-vous, où et dans quel ordre ?",
    "Quelles solutions envisagez-vous ? Avantages, risques, coût, délai.",
    "Votre décision, justifiée par rapport au calendrier et au budget.",
  ];
  qs.forEach((q, i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.25 + Math.floor(i / 2) * 1.05;
    card(s, x, y, 4.4, 0.92);
    s.addShape(pres.shapes.OVAL, { x: x + 0.18, y: y + 0.18, w: 0.56, h: 0.56, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: nm("num") });
    txt(s, String(i + 1), { x: x + 0.18, y: y + 0.18, w: 0.56, h: 0.56, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, q, { x: x + 0.9, y: y + 0.08, w: 3.35, h: 0.76, fontSize: 13.5, valign: "middle" });
  });
  txt(s, [{ text: "Envoyez votre réponse dans la conversation. ", options: { bold: true, color: C.accent1 } }, { text: "Pas de solution ici : on l’analyse ensemble." }], { x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 14 });
  s.addNotes(
    "Les six étapes correspondent à la méthode de diagnostic que vous travaillerez en détail dans le module M2. Prenez-les dans l'ordre.\n\n" +
    "1. Symptômes et causes : l'eau qui stagne est un symptôme. Le Clegg élevé est une mesure, pas une cause. Qu'est-ce qui pourrait expliquer chaque symptôme ?\n" +
    "2. Hypothèses : au moins quatre, et chacune doit pouvoir expliquer pourquoi l'eau stagne ici et pas ailleurs.\n" +
    "3. Informations manquantes : qu'auriez-vous besoin de savoir pour départager vos hypothèses ?\n" +
    "4. Mesures : lesquelles, à quels endroits (n'oubliez pas les zones témoins), dans quel ordre (la plus rapide et la plus informative d'abord).\n" +
    "5. Solutions : pour chacune, ce qu'elle règle, ce qu'elle risque, ce qu'elle coûte et combien de temps elle immobilise le terrain.\n" +
    "6. Décision : ce que vous feriez maintenant, ce que vous feriez pendant la trêve, et ce que vous changeriez pour éviter que le problème revienne.\n\n" +
    "Une bonne réponse n'a pas besoin d'être longue. Elle doit être structurée et justifiée."
  );

  // ===== 23. Erreurs d'expert =====
  pres.addSection({ title: "Conclusion" });
  s = content("Erreurs que font aussi les experts", "Conclusion");
  const errs = [
    ["Sabler avec un sable « qui y ressemble »", "Rupture de stock, prix. Résultat : une couche qui bloque l’eau pendant des années."],
    ["Aérer toujours à la même profondeur", "La machine est réglée une fois pour toutes. Résultat : une semelle au fond des trous."],
    ["Lire une mesure sans noter l’humidité", "On confond « sol sec » et « sol compacté », puis on traite le mauvais problème."],
    ["Accuser les drains quand l’eau stagne en surface", "La cause est souvent dans les premiers centimètres, beaucoup moins chère à traiter."],
    ["Mettre de l’engrais sur une zone asphyxiée qui jaunit", "Le symptôme est traité, la cause reste. L’azote peut même se perdre."],
    ["Ne pas suivre la matière organique", "Elle s’accumule lentement, sans signe visible, jusqu’au jour où l’eau ne passe plus."],
  ];
  errs.forEach(([t, d], i) => {
    const x = 0.5 + (i % 3) * 3.07, y = 1.25 + Math.floor(i / 3) * 1.85;
    card(s, x, y, 2.9, 1.7);
    iconCircle(s, "alerte", x + 0.15, y + 0.15, 0.5, C.accent5);
    txt(s, t, { x: x + 0.75, y: y + 0.1, w: 2.05, h: 0.62, fontSize: 13, bold: true, valign: "middle" });
    txt(s, d, { x: x + 0.15, y: y + 0.78, w: 2.6, h: 0.85, fontSize: 11.5, color: C.text2 });
  });
  s.addNotes(
    "Ces erreurs ne sont pas des erreurs de débutant. Elles sont faites par des professionnels expérimentés, parce qu'elles sont commodes ou parce que leurs conséquences arrivent longtemps après.\n\n" +
    "1. Le sable de sablage : changer de sable sans analyse est tentant quand le fournisseur habituel est en rupture. Mais une couche de quelques centimètres d'un sable incompatible peut créer une nappe perchée qui dure des années (on ne retire pas une couche facilement).\n\n" +
    "2. La profondeur d'aération fixe : par habitude, la machine garde le même réglage toute l'année. Risque : une semelle au fond des trous (diapositive 24).\n\n" +
    "3. La mesure sans humidité : c'est l'erreur E4 du diagnostic. Une lecture de Clegg ou de pénétromètre élevée en été peut simplement signifier que le sol est sec.\n\n" +
    "4. Les drains : quand l'eau stagne en surface sur une zone de forte utilisation alors qu'elle part ailleurs, regardez d'abord les premiers centimètres. Les drains sont plus profonds et concernent en général des surfaces plus larges. Ne les excluez pas, mais ne commencez pas par eux.\n\n" +
    "5. L'engrais sur une zone asphyxiée : la couleur revient parfois un peu, ce qui donne l'impression que ça marche. Mais la cause (manque d'oxygène, racines courtes) reste, et une partie de l'azote se perd par dénitrification.\n\n" +
    "6. La matière organique : elle s'accumule de quelques fractions de pourcent par an, sans symptôme visible, jusqu'à un seuil où l'infiltration s'effondre. Seul un suivi régulier par tranches permet de l'anticiper."
  );

  // ===== 24. À retenir =====
  s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Conclusion" });
  s.addText("À retenir", { placeholder: "title" });
  const keep = [
    "La taille et la régularité des grains décident du comportement du sol.",
    "Les gros pores portent l’air et le drainage. Le piétinement les détruit en premier.",
    "Un sol compacté manque d’oxygène avant de manquer d’eau.",
    "L’humidité au moment du passage décide des dégâts.",
    "Les premiers centimètres commandent l’infiltration.",
    "Deux couches différentes = de l’eau bloquée.",
    "Une mesure sans humidité, sans unité et sans témoin ne prouve rien.",
  ];
  keep.forEach((k, i) => {
    const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4;
    const x = 0.5 + col * 4.6, y = 1.25 + row * 0.95;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.1, w: 0.6, h: 0.6, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: nm("num") });
    txt(s, String(i + 1), { x, y: y + 0.1, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: C.text1, align: "center", valign: "middle" });
    txt(s, k, { x: x + 0.8, y, w: 3.6, h: 0.8, fontSize: 14.5, color: C.background1, valign: "middle" });
  });
  s.addNotes(
    "Ces sept principes doivent être retenus par cœur, mais surtout compris : pour chacun, vous devez pouvoir expliquer le mécanisme et donner un exemple sur votre terrain.\n\n" +
    "Exercice de révision : expliquez chaque principe à voix haute, en une ou deux phrases, comme si vous le présentiez à un nouveau collègue. Si vous bloquez sur l'un d'eux, relisez la diapositive correspondante :\n" +
    "1 → diapositives 5, 6 et 8 ;\n2 → diapositives 7 et 8 ;\n3 → diapositives 12, 13 et 14 ;\n4 → diapositive 10 ;\n5 → diapositive 16 ;\n6 → diapositive 17 ;\n7 → diapositives 14, 21 et 22."
  );

  // ===== 25. Votre semaine =====
  s = content("Votre travail pour ce module", "Conclusion");
  const plan = [
    ["livre", C.accent1, "Lire", "Diaporama et notes du formateur. Relire les parties 2 et 3."],
    ["question", C.accent2, "Étude de cas", "Envoyer votre raisonnement en 6 étapes dans la conversation."],
    ["loupe", C.accent3, "Mission terrain", "3 carottes (zone usée, moyenne, témoin) + photos + humidité et Clegg aux mêmes points. Avec l’accord de votre responsable."],
    ["check", C.accent6, "Évaluation finale", "Questions de raisonnement, cas inédit, transfert sur gazon naturel à Bordeaux, note d’une page."],
  ];
  plan.forEach(([ic, col, t, d], i) => {
    const y = 1.22 + i * 0.93;
    card(s, 0.5, y, 9.0, 0.82);
    iconCircle(s, ic, 0.65, y + 0.11, 0.6, col);
    txt(s, t, { x: 1.45, y: y + 0.08, w: 2.0, h: 0.66, fontSize: 15, bold: true, valign: "middle" });
    txt(s, d, { x: 3.45, y: y + 0.08, w: 5.9, h: 0.66, fontSize: 13, color: C.text2, valign: "middle" });
  });
  s.addNotes(
    "Organisation conseillée sur 4 semaines, à 5 h par semaine :\n" +
    "• Semaine 1 : lecture complète du diaporama et des notes, puis étude de cas (réponse dans la conversation).\n" +
    "• Semaine 2 : analyse de votre réponse avec moi, approfondissement, puis mission terrain (les 3 carottes).\n" +
    "• Semaine 3 : interprétation de vos carottes et de vos mesures avec moi. Vérification des missions du bilan : nom du système hybride et profondeur des fibres, analyse granulométrique, masse du marteau Clegg, longueur des tiges de la sonde.\n" +
    "• Semaine 4 : révision des 7 principes, puis évaluation finale.\n\n" +
    "L'évaluation finale comprendra : des questions de raisonnement (pas de récitation), une nouvelle étude de cas, une question de transfert (le même problème sur le gazon naturel de Bordeaux, sous un climat océanique) et une note d'une page à destination de votre responsable.\n\n" +
    "Rappel de validation : le module est validé quand vous appliquez correctement les notions à deux contextes différents."
  );

  // ===== 26. Sources =====
  s = content("Sources", "Conclusion");
  const src = [
    "USGA — Recommandations pour la construction des greens (révision 2004, résumé, archive MSU)",
    "Pitchcare — USGA golf putting green construction materials",
    "Bengough et al. (2011) — Root elongation, water stress, and mechanical impedance, J. Exp. Bot.",
    "Université du Nebraska — High temps are hard on cool-season grasses",
    "Archive MSU — Articles sur la saturation, l’oxygène et la chaleur (2004, 2005)",
    "Iowa State University — Soil Aeration (manuel de science du sol)",
    "Université du Nebraska — Topdressing 101 (GCSAA, 2025)",
    "USGA Green Section — Compatibilité des sables de sablage",
    "GCM — Verdure : Water drops through time (test de la goutte)",
    "UMass — Management of compaction : coring ; GCSAA — Giving soil compaction the old heave ho",
    "SIS Pitches — SISGrass (fabricant) ; Wikipédia — GrassMaster",
    "Penn State / STMA — Dureté des terrains et Gmax",
  ];
  txt(s, src.slice(0, 6).map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < 5 } })), { x: 0.5, y: 1.2, w: 4.4, h: 3.6, fontSize: 11, paraSpaceAfter: 6 });
  txt(s, src.slice(6).map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < 5 } })), { x: 5.1, y: 1.2, w: 4.4, h: 3.6, fontSize: 11, paraSpaceAfter: 6 });
  txt(s, "Liens complets dans les notes du formateur.", { x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 10, italic: true, color: C.text2 });
  s.addNotes(
    "Liens consultables (vérifiés lors de la préparation du module) :\n\n" +
    "USGA, propriétés physiques recommandées (résumé de la révision 2004) : https://archive.lib.msu.edu/tic/stnew/article/2004sum9.pdf\n" +
    "Pitchcare, USGA golf putting green construction materials : https://www.pitchcare.com/blogs/news/usga-golf-putting-green-construction-materials\n" +
    "Bengough, McKenzie, Hallett, Valentine (2011), Root elongation, water stress, and mechanical impedance : https://discovery.dundee.ac.uk/en/publications/root-elongation-water-stress-and-mechanical-impedance-a-review-of/\n" +
    "Université du Nebraska, chaleur et gazons de climat tempéré : https://turf.unl.edu/news/high-temps-are-hard-cool-season-grasses/\n" +
    "Archive MSU, sol saturé et anaérobiose : https://archive.lib.msu.edu/tic/bigga/gki/article/2004oct26.pdf\n" +
    "Archive MSU, chaleur, humidité et pertes de racines : https://archive.lib.msu.edu/tic/bigga/gki/article/2005aug14.pdf\n" +
    "Iowa State University, Soil Aeration : https://iastate.pressbooks.pub/isudp-2025-201/chapter/soil-aeration/\n" +
    "Université du Nebraska, Topdressing 101 : https://turf.unl.edu/sites/unl.edu.ianr.agronomy-horticulture.turf/files/media/file/GCSAA2025_Topdressing101.pdf\n" +
    "USGA Green Section (document sur les sables de sablage) : https://www.usga.org/content/dam/usga/pdf/imported/course-care/100117.pdf\n" +
    "GCM, Water drops through time : https://gcmonline.com/research/news/verdure-water-drops-through-time\n" +
    "UMass, Management of compaction : coring : https://www.umass.edu/agriculture-food-environment/turf/fact-sheets/management-of-compaction-coring\n" +
    "GCSAA, Giving soil compaction the old heave ho : https://www.gcsaa.org/docs/default-source/research-and-information/agronomy/giving-soil-compaction-the-old-heave-ho.pdf?sfvrsn=2\n" +
    "SIS Pitches, SISGrass : https://www.sispitches.com/hybrid-pitches/sisgrass-lite\n" +
    "Wikipédia, GrassMaster : https://en.wikipedia.org/wiki/GrassMaster\n" +
    "Penn State, Gmax : https://plantscience.psu.edu/research/centers/ssrc/documents/ftfgmax.pdf\n" +
    "STMA, Field Hardness Testing : https://www.stma.org/wp-content/uploads/2019/09/Field-Hardness-Testing-STMAInstitute-logo.pdf\n\n" +
    "Limites à connaître :\n" +
    "• Les valeurs USGA concernent les greens de golf ; elles servent de repère, pas de norme pour le football.\n" +
    "• Les chiffres GrassMaster viennent d'une source secondaire non sourcée.\n" +
    "• Le seuil de 2 MPa et les températures racinaires sont des ordres de grandeur, qui varient selon le sol, l'espèce et la méthode.\n" +
    "• La classification Dekker et Ritsema (1999) est citée d'après des sources secondaires (GCM, archive MSU)."
  );

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("OK", OUT);
}
build().catch((e) => { console.error(e); process.exit(1); });
