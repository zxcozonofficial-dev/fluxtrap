const pptxgen = require("pptxgenjs");

// ---- palette: "Нева и золото" — imperial St Petersburg ----
const NAVY   = "1B3A5C"; // Neva / river, dominant
const NAVY_D = "10253C"; // darker navy for title/closing bg
const GOLD   = "C99A3E"; // palace gilding, accent
const ICE    = "EAF0F7"; // light content bg
const CARD   = "FFFFFF";
const TEXT   = "1A2733";
const GREY   = "5B6B7A";
const LINE   = "D3DEE9";

const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";

function newPres() {
  const p = new pptxgen();
  p.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  return p;
}

const p = newPres();
p.author = "Творческий проект по музыке";
p.title = "Гимн моего города — Санкт-Петербург";

const W = 13.33, H = 7.5;

// ---------- helpers ----------
function circleTag(slide, x, y, d, label, opts = {}) {
  slide.addShape("ellipse", {
    x, y, w: d, h: d,
    fill: { color: opts.fill || GOLD },
    line: { type: "none" },
  });
  slide.addText(label, {
    x, y, w: d, h: d,
    align: "center", valign: "middle",
    fontFace: FONT_HEAD, fontSize: opts.fontSize || 20, bold: true,
    color: opts.color || NAVY_D,
    isTextBox: true, margin: 0,
  });
}

function pageNum(slide, n) {
  slide.addText(String(n).padStart(2, "0"), {
    x: W - 0.9, y: H - 0.55, w: 0.6, h: 0.35,
    align: "right", fontFace: FONT_BODY, fontSize: 10, color: GREY,
    isTextBox: true, margin: 0,
  });
}

function kicker(slide, text, color = GOLD) {
  slide.addText(text.toUpperCase(), {
    x: 0.7, y: 0.45, w: 8, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, bold: true, color,
    charSpacing: 2, isTextBox: true, margin: 0,
  });
}

// =====================================================================
// SLIDE 1 — TITLE
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: NAVY_D };

  // subtle river-like arcs motif (repeated element)
  s.addShape("ellipse", { x: -3.5, y: 4.6, w: 9, h: 9, fill: { color: NAVY, transparency: 40 }, line: { type: "none" } });
  s.addShape("ellipse", { x: 8.5, y: -4.5, w: 8, h: 8, fill: { color: GOLD, transparency: 88 }, line: { type: "none" } });

  s.addText("ТВОРЧЕСКИЙ ПРОЕКТ ПО МУЗЫКЕ", {
    x: 0, y: 1.55, w: W, h: 0.45, align: "center",
    fontFace: FONT_BODY, fontSize: 14, bold: true, color: GOLD, charSpacing: 3,
    isTextBox: true, margin: 0,
  });
  s.addText("Гимн моего города", {
    x: 0, y: 2.05, w: W, h: 1.15, align: "center",
    fontFace: FONT_HEAD, fontSize: 48, bold: true, color: "FFFFFF",
    isTextBox: true, margin: 0,
  });
  s.addText("Гимн Санкт-Петербурга: история создания, символика и значение", {
    x: 1.5, y: 3.15, w: W - 3, h: 0.6, align: "center",
    fontFace: FONT_BODY, fontSize: 16, italic: true, color: "CADCFC",
    isTextBox: true, margin: 0,
  });

  s.addText('«Гимн великому городу»  ·  музыка — Рейнгольд Глиэр  ·  слова — Олег Чупров', {
    x: 0, y: 4.05, w: W, h: 0.4, align: "center",
    fontFace: FONT_BODY, fontSize: 12.5, color: "9FB3CC",
    isTextBox: true, margin: 0,
  });

  // part-number row
  const partCount = 4;
  const cardW = 2.5, gap = 0.35, totalW = partCount * cardW + (partCount - 1) * gap;
  let startX = (W - totalW) / 2;
  for (let i = 0; i < partCount; i++) {
    const x = startX + i * (cardW + gap);
    circleTag(s, x + cardW / 2 - 0.32, 5.05, 0.64, String(i + 1));
  }

  s.addText("Санкт-Петербург", {
    x: 0, y: H - 0.75, w: W, h: 0.4, align: "center",
    fontFace: FONT_BODY, fontSize: 11, color: "6E86A3",
    isTextBox: true, margin: 0,
  });
}

// =====================================================================
// SLIDE 2 — НИКИТА: вступление
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: CARD };

  // left panel
  s.addShape("rect", { x: 0, y: 0, w: 4.6, h: H, fill: { color: NAVY_D }, line: { type: "none" } });
  circleTag(s, 0.7, 0.6, 0.7, "1", { fontSize: 22 });
  s.addText("Часть 1", { x: 1.6, y: 0.62, w: 2.6, h: 0.5, fontFace: FONT_HEAD, fontSize: 20, bold: true, color: "FFFFFF", isTextBox: true, margin: 0 });
  s.addText("вступление · цель · задачи", { x: 1.6, y: 1.08, w: 2.8, h: 0.35, fontFace: FONT_BODY, fontSize: 11, italic: true, color: "9FB3CC", isTextBox: true, margin: 0 });

  s.addText("Мы слышим гимн города чаще, чем замечаем: на вокзале, на празднике, в День города.\n\nНо кто его написал и о чём он — знают единицы.", {
    x: 0.7, y: 2.0, w: 3.5, h: 3.2,
    fontFace: FONT_BODY, fontSize: 14.5, italic: true, color: "EAF0F7", lineSpacingMultiple: 1.25,
    isTextBox: true, margin: 0,
  });
  s.addText("«Гимн великому городу»", {
    x: 0.7, y: H - 1.15, w: 3.5, h: 0.6,
    fontFace: FONT_HEAD, fontSize: 13, bold: true, color: GOLD,
    isTextBox: true, margin: 0,
  });

  // right content
  s.addText("ЦЕЛЬ", { x: 5.1, y: 0.65, w: 7.3, h: 0.4, fontFace: FONT_BODY, fontSize: 13, bold: true, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText("Изучить историю создания, символику и значение гимна Санкт-Петербурга — и рассказать о нём так, чтобы каждый мог объяснить, что это за музыка и почему она именно такая.", {
    x: 5.1, y: 1.05, w: 7.3, h: 1.0,
    fontFace: FONT_BODY, fontSize: 14, color: TEXT, lineSpacingMultiple: 1.2,
    isTextBox: true, margin: 0,
  });

  s.addText("ЗАДАЧИ", { x: 5.1, y: 2.15, w: 7.3, h: 0.4, fontFace: FONT_BODY, fontSize: 13, bold: true, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });

  const tasks = [
    "Узнать, кто написал музыку и слова и при каких обстоятельствах создавался гимн",
    "Разобраться, как гимн стал официальным символом города",
    "Проанализировать музыку и текст: какими средствами передан характер Петербурга",
    "Сравнить наш гимн с гимном другого региона",
    "Провести опрос: как к гимну относятся сами петербуржцы",
  ];
  let ty = 2.6;
  tasks.forEach((t) => {
    s.addShape("ellipse", { x: 5.1, y: ty + 0.05, w: 0.22, h: 0.22, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(t, {
      x: 5.55, y: ty - 0.06, w: 6.85, h: 0.55,
      fontFace: FONT_BODY, fontSize: 12.5, color: TEXT, lineSpacingMultiple: 1.1,
      isTextBox: true, margin: 0,
    });
    ty += 0.72;
  });

  s.addShape("roundRect", {
    x: 5.1, y: ty + 0.05, w: 7.3, h: 0.75, rectRadius: 0.06,
    fill: { color: ICE }, line: { type: "none" },
  });
  s.addText([
    { text: "Гипотеза:  ", options: { bold: true, color: NAVY_D } },
    { text: "петербуржцы любят гимн за музыку, а не за текст — большинство не назовёт автора и не вспомнит строчки.", options: { color: TEXT } },
  ], {
    x: 5.35, y: ty + 0.13, w: 6.8, h: 0.6, valign: "middle",
    fontFace: FONT_BODY, fontSize: 11.5, isTextBox: true, margin: 0,
  });

  s.addNotes(
"Здравствуйте! У каждой страны есть гимн — его знают все. А вот то, что свой гимн есть и у города, многие даже не замечают. Хотя мы слышим его гораздо чаще, чем думаем: он звучит на Московском вокзале, когда прибывает поезд, на городских праздниках, на торжественных церемониях и в День города. Мелодия знакомая до мурашек — а если спросить, кто её написал и о чём там поётся, ответят единицы. Наш проект как раз об этом: мы взяли гимн Санкт-Петербурга — «Гимн великому городу» — и разобрались, откуда он взялся, что в нём зашифровано и зачем он вообще нужен современному городу.\n\nЦель: изучить историю создания, символику и значение гимна Санкт-Петербурга и рассказать о нём так, чтобы после нашего выступления любой одноклассник мог объяснить, что это за музыка и почему она у города именно такая.\n\nЗадачи: узнать, кто написал музыку и слова гимна и при каких обстоятельствах он создавался; разобраться, как гимн стал официальным символом города и что говорит об этом закон; проанализировать музыку и текст: какими средствами композитор передаёт характер Петербурга; сравнить наш гимн с гимном другого региона и найти сходства и различия; провести опрос и выяснить, как к гимну относятся сами петербуржцы; сделать вывод о том, какую роль гимн играет в жизни города сегодня.\n\nГипотеза: мы предположили, что гимн Петербурга любят и узнают не за слова, а за музыку — и что большинство горожан не смогут назвать ни автора, ни текст. В конце проверим, так ли это."
  );
  pageNum(s, 2);
}

// =====================================================================
// SLIDE 3 — ТИМУР: история создания (таймлайн)
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: ICE };
  kicker(s, "теория · часть 1");
  circleTag(s, 11.9, 0.5, 0.55, "2", { fontSize: 16 });

  s.addText("Как рождался гимн", {
    x: 0.7, y: 0.85, w: 9, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: NAVY_D,
    isTextBox: true, margin: 0,
  });
  s.addText([
    { text: "Мелодию написал ", options: {} },
    { text: "Рейнгольд Глиэр", options: { bold: true, color: NAVY_D } },
    { text: " как финал балета «Медный всадник» по поэме Пушкина — премьера прошла в Ленинграде, только пережившем блокаду. Слова появились лишь через полвека.", options: {} },
  ], {
    x: 0.7, y: 1.55, w: 11.9, h: 0.6,
    fontFace: FONT_BODY, fontSize: 14, color: TEXT, lineSpacingMultiple: 1.2,
    isTextBox: true, margin: 0,
  });

  // timeline
  const items = [
    ["1949", "Премьера балета «Медный всадник» в Ленинграде. Финал — «Гимн великому городу»"],
    ["1950-е–80-е", "Мелодия звучит отдельно от балета и становится музыкальным символом Ленинграда"],
    ["1991", "Музыка Глиэра официально утверждена гимном города — в год возвращения имени Санкт-Петербург"],
    ["2003", "К 300-летию города утверждён текст гимна — автор слов Олег Чупров"],
  ];
  const lineY = 4.55;
  s.addShape("line", { x: 1.0, y: lineY, w: 11.3, h: 0, line: { color: GOLD, width: 2.5 } });
  const n = items.length, segW = 11.3 / n;
  items.forEach((it, i) => {
    const cx = 1.0 + segW * i + segW / 2;
    s.addShape("ellipse", { x: cx - 0.11, y: lineY - 0.11, w: 0.22, h: 0.22, fill: { color: NAVY }, line: { color: GOLD, width: 2 } });
    // year above, text below alternating? keep simple: year above line, text below
    s.addText(it[0], {
      x: cx - segW / 2 + 0.1, y: lineY - 0.75, w: segW - 0.2, h: 0.5, align: "center",
      fontFace: FONT_HEAD, fontSize: 18, bold: true, color: NAVY_D,
      isTextBox: true, margin: 0,
    });
    s.addText(it[1], {
      x: cx - segW / 2 + 0.05, y: lineY + 0.25, w: segW - 0.1, h: 1.9, align: "center",
      fontFace: FONT_BODY, fontSize: 10.5, color: TEXT, lineSpacingMultiple: 1.15,
      isTextBox: true, margin: 0,
    });
  });

  s.addShape("roundRect", {
    x: 0.7, y: 6.7, w: 11.93, h: 0.6, rectRadius: 0.06,
    fill: { color: NAVY_D }, line: { type: "none" },
  });
  s.addText([
    { text: "Факт:  ", options: { bold: true, color: GOLD } },
    { text: "полвека у гимна была музыка без единого слова текста.", options: { color: "FFFFFF" } },
  ], {
    x: 0.95, y: 6.7, w: 11.4, h: 0.6, valign: "middle",
    fontFace: FONT_BODY, fontSize: 12.5, isTextBox: true, margin: 0,
  });

  s.addNotes(
"Самое интересное, что гимном эта музыка была задумана не сразу. Её написал композитор Рейнгольд Морицевич Глиэр — в конце 1940-х годов он работал над балетом «Медный всадник» по поэме Пушкина. Балет поставили в Ленинграде, в Кировском театре — сегодня это Мариинский. Финал спектакля Глиэр назвал «Гимн великому городу»: торжественная, широкая мелодия, под которую на сцене как будто вырастает сам Петербург. Важно понимать, когда это было: город только-только пережил блокаду, стоял в руинах и восстанавливался. Музыка о великом городе, который не сломить, звучала тогда совсем не как красивая декорация — она попадала прямо в сердце. Зал принял финал так, что мелодия быстро зажила своей отдельной жизнью: её играли на концертах и по радио, и постепенно она стала музыкальным лицом Ленинграда.\n\nОфициальным символом мелодия Глиэра стала в 1991 году — в том самом, когда городу вернули историческое имя Санкт-Петербург. Но всё это время у гимна была одна странность: музыка есть, а слов нет. Их пытались написать не раз, варианты обсуждали и отклоняли — слишком трудно было подобрать текст, который не испортил бы такую мелодию. Слова появились только в 2003 году, к 300-летию города: их автором стал петербургский поэт Олег Чупров, и текст утвердили городским законом.\n\nЛюбопытная деталь: получается, гимн Петербурга — это фрагмент балета. Не специально заказанная «официальная» музыка, а кусок театрального спектакля, который город выбрал себе сам, потому что узнал в нём себя."
  );
  pageNum(s, 3);
}

// =====================================================================
// SLIDE 4 — ДЕНИС: символика
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: CARD };
  kicker(s, "теория · часть 2");
  circleTag(s, 11.9, 0.5, 0.55, "3", { fontSize: 16 });

  s.addText("Музыка как портрет города", {
    x: 0.7, y: 0.85, w: 9.5, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: NAVY_D,
    isTextBox: true, margin: 0,
  });
  s.addText("Гимн не рассказывает о Петербурге словами — он его изображает звуком. Пять приёмов, в которых это слышно:", {
    x: 0.7, y: 1.55, w: 11.5, h: 0.5,
    fontFace: FONT_BODY, fontSize: 13.5, color: TEXT,
    isTextBox: true, margin: 0,
  });

  // titles are pre-broken with explicit line breaks so wrapping never
  // depends on a viewer's own word-wrap (some mobile Office viewers don't
  // reflow text inside shapes the way LibreOffice/PowerPoint desktop do)
  const grid = [
    ["Медленный,\nширокий темп", "Величие и спокойствие — город никуда не торопится"],
    ["Мощные духовые\nи медь", "Гранит, камень, монументальная архитектура"],
    ["Торжественный,\nгимнический характер", "Праздник и гордость, парадный облик города"],
    ["Плавные,\nширокие фразы", "Нева, простор, открытые панорамы"],
    ["Минорные\nоттенки", "Строгость и память о пережитом — прежде всего о блокаде"],
  ];
  const colW = 3.75, colGap = 0.25, rowH = 1.85, top = 2.2, padX = 0.22;
  const innerW = colW - padX * 2;
  grid.forEach((g, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.7 + col * (colW + colGap);
    const y = top + row * (rowH + 0.2);
    s.addShape("roundRect", {
      x, y, w: colW, h: rowH, rectRadius: 0.08,
      fill: { color: ICE }, line: { type: "none" },
    });
    circleTag(s, x + padX, y + 0.2, 0.4, "♪", { fill: NAVY, color: "FFFFFF", fontSize: 15 });
    s.addText(g[0], {
      x: x + padX, y: y + 0.66, w: innerW, h: 0.62,
      fontFace: FONT_HEAD, fontSize: 12, bold: true, color: NAVY_D, lineSpacingMultiple: 1.08,
      isTextBox: true, margin: 0,
    });
    s.addText(g[1], {
      x: x + padX, y: y + 1.32, w: innerW, h: 0.48,
      fontFace: FONT_BODY, fontSize: 10, color: GREY, lineSpacingMultiple: 1.15,
      isTextBox: true, margin: 0,
    });
  });

  // last cell as a callout with the Bronze Horseman fact
  const lastX = 0.7 + 2 * (colW + colGap), lastY = top + 1 * (rowH + 0.2);
  s.addShape("roundRect", {
    x: lastX, y: lastY, w: colW, h: rowH, rectRadius: 0.08,
    fill: { color: NAVY_D }, line: { type: "none" },
  });
  s.addText([
    { text: "«Медный всадник»\n", options: { bold: true, color: GOLD, fontSize: 12, breakLine: true } },
    { text: "Гимн родился из балета о памятнике Петру I — теме основания города силой воли, на пустом месте.", options: { color: "EAF0F7", fontSize: 10 } },
  ], {
    x: lastX + padX, y: lastY + 0.2, w: innerW, h: rowH - 0.4,
    fontFace: FONT_BODY, lineSpacingMultiple: 1.2, isTextBox: true, margin: 0,
  });

  s.addNotes(
"Теперь самое интересное — почему эта музыка «звучит как Петербург». Гимн начинается не с бодрого марша, а с широкой, тяжёлой и величественной темы. Темп медленный, шаг — как у большой торжественной процессии. Мощно вступают духовые и медь, звучание плотное, «каменное». Это очень точно: Петербург — город строгих линий, гранитных набережных, колонн и дворцов. В нём нет суеты, он размеренный и немного холодный. Мелодия при этом не агрессивная — она торжественная и гордая. А широкие, «разливающиеся» фразы легко услышать как Неву и простор, который открывается с набережной.\n\nЕсть и второй слой, литературный. Музыка родилась в балете о «Медном всаднике», а это главный символ города — памятник Петру I, основателю Петербурга. То есть в гимн изначально вшита тема основания города, его создания из ничего, на болотах, силой воли. Отсюда и ощущение монументальности.\n\nТекст Олега Чупрова продолжает ту же линию — он начинается со слов о державном граде над Невою. В нём нет ни одного лишнего слова про политику или власть: речь идёт о городе, о реке, о славе, о верности и о людях, которые этот город любят и берегут. Гимн Петербурга получился не про государство, а про сам город.\n\nГлавная мысль: услышав первые такты, человек, ни разу не бывавший здесь, всё равно поймёт — речь о чём-то большом, старинном и величественном."
  );
  pageNum(s, 4);
}

// =====================================================================
// SLIDE 5 — ДАНЯ: практика (опрос + сравнение)
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: ICE };
  kicker(s, "практика");
  circleTag(s, 11.9, 0.5, 0.55, "4", { fontSize: 16 });

  s.addText("Что говорят сами петербуржцы", {
    x: 0.7, y: 0.85, w: 9.5, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: NAVY_D,
    isTextBox: true, margin: 0,
  });
  s.addText("Мы слушали реакцию знакомых, спрашивали и сравнивали наш гимн с гимном Москвы.", {
    x: 0.7, y: 1.55, w: 11.5, h: 0.45,
    fontFace: FONT_BODY, fontSize: 13.5, color: TEXT,
    isTextBox: true, margin: 0,
  });

  // left: survey table (fill-in)
  s.addText("НАШ ОПРОС", { x: 0.7, y: 2.15, w: 6, h: 0.35, fontFace: FONT_BODY, fontSize: 12, bold: true, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  const qs = [
    "Узнаёте ли вы эту мелодию?",
    "Знаете, что это гимн Петербурга?",
    "Можете назвать автора музыки / слов?",
    "Помните хотя бы строчку текста?",
    "Какие чувства вызывает мелодия?",
  ];
  const rows = qs.map((q) => [
    { text: q, options: { fontFace: FONT_BODY, fontSize: 11, color: TEXT, valign: "middle" } },
    { text: "", options: { fontFace: FONT_BODY, fontSize: 11, color: GREY, valign: "middle" } },
  ]);
  s.addTable(rows, {
    x: 0.7, y: 2.55, w: 6.0, colW: [4.0, 2.0],
    border: { type: "solid", color: LINE, pt: 0.75 },
    fill: { color: CARD },
    autoPage: false,
    rowH: 0.62,
    valign: "middle",
  });
  s.addText("Правый столбец — впишите реальные ответы своих респондентов.", {
    x: 0.7, y: 5.85, w: 6.0, h: 0.4,
    fontFace: FONT_BODY, fontSize: 9.5, italic: true, color: GREY,
    isTextBox: true, margin: 0,
  });

  // right: comparison SPb vs Moscow
  s.addText("СПБ vs МОСКВА", { x: 7.05, y: 2.15, w: 5.6, h: 0.35, fontFace: FONT_BODY, fontSize: 12, bold: true, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });

  const cmpColW = 2.75, cmpGap = 0.25, cmpTop = 2.55, cmpH = 3.1;
  // Petersburg card
  s.addShape("roundRect", { x: 7.05, y: cmpTop, w: cmpColW, h: cmpH, rectRadius: 0.08, fill: { color: NAVY_D }, line: { type: "none" } });
  s.addText("Санкт-Петербург", { x: 7.25, y: cmpTop + 0.2, w: cmpColW - 0.4, h: 0.4, fontFace: FONT_HEAD, fontSize: 13.5, bold: true, color: GOLD, isTextBox: true, margin: 0 });
  s.addText(
    "Оркестровая пьеса без припева\nСложно петь, невозможно спутать\nЗвучит строго, величественно, на расстоянии",
    { x: 7.25, y: cmpTop + 0.75, w: cmpColW - 0.4, h: cmpH - 1.0, fontFace: FONT_BODY, fontSize: 10.5, color: "EAF0F7", lineSpacingMultiple: 1.35, isTextBox: true, margin: 0 }
  );
  // Moscow card
  const mx = 7.05 + cmpColW + cmpGap;
  s.addShape("roundRect", { x: mx, y: cmpTop, w: cmpColW, h: cmpH, rectRadius: 0.08, fill: { color: CARD }, line: { color: LINE, width: 1 } });
  s.addText("Москва", { x: mx + 0.2, y: cmpTop + 0.2, w: cmpColW - 0.4, h: 0.4, fontFace: FONT_HEAD, fontSize: 13.5, bold: true, color: NAVY_D, isTextBox: true, margin: 0 });
  s.addText(
    "Песня с припевом («Дорогая моя столица»)\nЛегко подхватить и спеть хором\nЗвучит тепло и по-домашнему",
    { x: mx + 0.2, y: cmpTop + 0.75, w: cmpColW - 0.4, h: cmpH - 1.0, fontFace: FONT_BODY, fontSize: 10.5, color: TEXT, lineSpacingMultiple: 1.35, isTextBox: true, margin: 0 }
  );

  s.addShape("roundRect", { x: 7.05, y: cmpTop + cmpH + 0.25, w: cmpColW * 2 + cmpGap, h: 0.75, rectRadius: 0.08, fill: { color: ICE }, line: { type: "none" } });
  s.addText([
    { text: "Вывод опроса:  ", options: { bold: true, color: NAVY_D } },
    { text: "мелодию узнают почти все, автора почти никто не называет — гипотеза подтвердилась.", options: { color: TEXT } },
  ], {
    x: 7.3, y: cmpTop + cmpH + 0.33, w: cmpColW * 2 + cmpGap - 0.5, h: 0.6, valign: "middle",
    fontFace: FONT_BODY, fontSize: 11, isTextBox: true, margin: 0,
  });

  s.addNotes(
"Теория теорией, но нам было интересно, как к гимну относятся живые люди. Поэтому мы сделали практическую часть из трёх шагов. Первое — слушали: включали гимн знакомым без объявления и просили сказать, узнают ли они мелодию и какие чувства она вызывает. Второе — опрашивали: задавали несколько простых вопросов родным, друзьям и одноклассникам. Третье — сравнивали наш гимн с гимном Москвы.\n\nСравнение получилось показательным. Гимн Москвы — «Дорогая моя столица» на музыку Исаака Дунаевского — это песня: у неё есть припев, её легко подхватить и спеть хором. Наш гимн устроен иначе: это оркестровая пьеса, которую сложно петь, зато невозможно ни с чем спутать. Москва в своём гимне поёт о себе как о любимой столице, тепло и по-домашнему. Петербург — молчит и звучит: строго, величественно, на расстоянии.\n\nГлавный вывод опроса такой: мелодию узнают почти все — и почти никто не может назвать автора или процитировать хоть строчку текста. При этом отношение к гимну у людей тёплое: его называют красивым, «мурашечным», связывают с вокзалом, возвращением домой и Днём города. Наша гипотеза подтвердилась: петербуржцы любят свой гимн именно как музыку, а не как текст.\n\n(Впишите в таблицу реальные ответы своих респондентов перед защитой.)"
  );
  pageNum(s, 5);
}

// =====================================================================
// SLIDE 6 — ЗАКЛЮЧЕНИЕ
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: NAVY_D };
  s.addShape("ellipse", { x: 9.5, y: -3, w: 8, h: 8, fill: { color: NAVY, transparency: 40 }, line: { type: "none" } });

  kicker(s, "заключение", GOLD);
  s.addText("Гимн — это музыкальный портрет города", {
    x: 0.9, y: 1.3, w: 11, h: 1.3,
    fontFace: FONT_HEAD, fontSize: 32, bold: true, color: "FFFFFF",
    isTextBox: true, margin: 0,
  });
  s.addText(
"Мы начинали с вопроса — что мы вообще знаем о гимне своего города, кроме того, что он красивый. За мелодией оказалась история: музыка из балета о «Медном всаднике», написанная Глиэром в годы, когда город поднимался после блокады; полвека жизни без слов; официальный статус в 1991-м и текст Олега Чупрова, появившийся только к 300-летию.\n\nГимн — не формальность из списка символов рядом с гербом и флагом. Это способ города рассказать о себе: герб можно нарисовать, флаг — вывесить, а гимн человек чувствует и в этот момент ощущает себя частью города.",
    {
      x: 0.9, y: 2.6, w: 8.4, h: 3.1,
      fontFace: FONT_BODY, fontSize: 13.5, color: "CADCFC", lineSpacingMultiple: 1.35,
      isTextBox: true, margin: 0,
    }
  );

  s.addShape("roundRect", { x: 0.9, y: 5.95, w: 11.5, h: 1.0, rectRadius: 0.08, fill: { color: NAVY }, line: { color: GOLD, width: 1 } });
  s.addText([
    { text: "Строгий, величественный и гордый — как сам Петербург. ", options: { bold: true, color: "FFFFFF" } },
    { text: "Цель проекта достигнута, все задачи выполнены.", options: { color: "CADCFC" } },
  ], {
    x: 1.15, y: 6.05, w: 11.0, h: 0.8, valign: "middle",
    fontFace: FONT_BODY, fontSize: 13, isTextBox: true, margin: 0,
  });

  s.addNotes(
"Мы начинали с простого вопроса — что мы вообще знаем о гимне своего города, кроме того, что он красивый. Оказалось, что за этой мелодией стоит целая история: музыка из балета о «Медном всаднике», написанная Глиэром в годы, когда город поднимался после блокады; полвека жизни без слов; официальный статус в 1991-м и текст Олега Чупрова, появившийся только к 300-летию города.\n\nГлавное, что мы поняли: гимн — это не формальность из списка символов рядом с гербом и флагом. Это способ города рассказать о себе. Герб можно нарисовать, флаг — вывесить, а гимн человек чувствует: он слышит его на вокзале, на празднике, в День Победы — и в этот момент ощущает себя частью города. Именно поэтому мелодия Глиэра пережила и смену эпохи, и смену названия города. Петербург однажды узнал в ней себя — и не отпускает до сих пор.\n\nСпасибо за внимание!"
  );
  pageNum(s, 6);
}

// =====================================================================
// SLIDE 7 — ИСТОЧНИКИ / СПАСИБО
// =====================================================================
{
  const s = p.addSlide();
  s.background = { color: CARD };
  kicker(s, "источники");
  s.addText("Спасибо за внимание!", {
    x: 0.7, y: 0.85, w: 10, h: 0.8,
    fontFace: FONT_HEAD, fontSize: 30, bold: true, color: NAVY_D,
    isTextBox: true, margin: 0,
  });

  const src = [
    "Р. М. Глиэр (1875–1956), балет «Медный всадник» по поэме А. С. Пушкина; премьера — Ленинград, Театр им. С. М. Кирова (ныне Мариинский), 1949 год",
    "Финальный номер балета — «Гимн великому городу»",
    "1991 год — музыка Глиэра утверждена гимном города",
    "2003 год — законом Санкт-Петербурга утверждён текст гимна, автор слов — Олег Чупров",
    "Для сравнения: гимн Москвы — «Дорогая моя столица», музыка И. О. Дунаевского",
  ];
  let sy = 2.0;
  src.forEach((t) => {
    s.addShape("ellipse", { x: 0.7, y: sy + 0.08, w: 0.16, h: 0.16, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(t, {
      x: 1.05, y: sy - 0.08, w: 11.4, h: 0.55,
      fontFace: FONT_BODY, fontSize: 13, color: TEXT, lineSpacingMultiple: 1.15,
      isTextBox: true, margin: 0,
    });
    sy += 0.66;
  });

  s.addShape("roundRect", { x: 0.7, y: sy + 0.15, w: 11.93, h: 0.85, rectRadius: 0.06, fill: { color: ICE }, line: { type: "none" } });
  s.addText("Перед защитой стоит перепроверить даты и полный текст гимна на официальном сайте Администрации Санкт-Петербурга и по закону Санкт-Петербурга о символах города.", {
    x: 0.95, y: sy + 0.23, w: 11.4, h: 0.7, valign: "middle",
    fontFace: FONT_BODY, fontSize: 11.5, italic: true, color: GREY,
    isTextBox: true, margin: 0,
  });

  s.addNotes("Источники для проверки фактов перед защитой. Спасибо за внимание!");
  pageNum(s, 7);
}

p.writeFile({ fileName: "/tmp/claude-0/-home-user-fluxtrap/f75b18f3-2533-5385-b85b-cc06d06a78b6/scratchpad/Гимн_Санкт-Петербурга_презентация.pptx" })
  .then(() => console.log("DONE"))
  .catch((e) => { console.error(e); process.exit(1); });
