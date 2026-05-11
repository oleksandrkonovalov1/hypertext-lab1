#!/usr/bin/env node

import {
  Document, Packer, Paragraph, TextRun, Header, ImageRun,
  AlignmentType, PageNumber, HeadingLevel, PageBreak,
  LevelFormat, BorderStyle,
} from "docx";
import { writeFileSync, readFileSync } from "fs";

const MM_TO_DXA = 56.693;
const PT_TO_HALF_PT = 2;
const FONT = "Times New Roman";
const FONT_CODE = "Courier New";
const BODY_SIZE = 14 * PT_TO_HALF_PT;
const CODE_SIZE = 10 * PT_TO_HALF_PT;
const LINE_SPACING_15 = 360;
const INDENT = Math.round(12.5 * MM_TO_DXA);

const margins = {
  top: Math.round(20 * MM_TO_DXA),
  bottom: Math.round(20 * MM_TO_DXA),
  left: Math.round(30 * MM_TO_DXA),
  right: Math.round(15 * MM_TO_DXA),
};

// --- Helpers ---

function tr(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: BODY_SIZE, ...opts });
}

function trBold(text, opts = {}) {
  return tr(text, { bold: true, ...opts });
}

function centered(runs, spacing = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto", ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function empty() {
  return centered(tr(""));
}

function body(text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: INDENT },
    alignment: AlignmentType.JUSTIFIED,
    children: [tr(text)],
  });
}

function bodyRuns(runs) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: INDENT },
    alignment: AlignmentType.JUSTIFIED,
    children: runs,
  });
}

function bodyNoIndent(text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
    children: [tr(text)],
  });
}

function heading1(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: `${number} ${title}`.toUpperCase(), font: FONT, size: BODY_SIZE, bold: true })],
  });
}

function heading2(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: INDENT },
    children: [new TextRun({ text: `${number} ${title}`, font: FONT, size: BODY_SIZE, bold: true })],
  });
}

function codeLine(text) {
  return new Paragraph({
    spacing: { after: 0, line: 240, lineRule: "auto" },
    children: [new TextRun({ text, font: FONT_CODE, size: CODE_SIZE })],
  });
}

function codeBlock(code) {
  return code.split("\n").map(line => codeLine(line));
}

function listingCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: INDENT },
    children: [tr(`Лістинг ${number} – ${title}`)],
  });
}

function figureCaption(number, title) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [tr(`Рисунок ${number} – ${title}`)],
  });
}

function imageFromFile(path, widthPx, heightPx) {
  const data = readFileSync(path);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 0 },
    children: [new ImageRun({
      type: "png",
      data,
      transformation: { width: widthPx, height: heightPx },
      altText: { title: path, description: path, name: path },
    })],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// --- Read source files ---
const BASE = "/Users/oleksandrkonovalov/Projects/nure/courses/hypertext-and-hypermedia/labs/lab_01";
const SRC = `${BASE}/src`;
const indexHtml = readFileSync(`${SRC}/index.html`, "utf-8");
const plantsHtml = readFileSync(`${SRC}/plants.html`, "utf-8");
const contactHtml = readFileSync(`${SRC}/contact.html`, "utf-8");
const styleCss = readFileSync(`${SRC}/style.css`, "utf-8");

// --- TITLE PAGE ---

const titlePage = [
  centered(tr("Міністерство освіти і науки України")),
  centered(tr("Харківський національний університет радіоелектроніки")),
  empty(),
  centered(tr("Кафедра програмної інженерії")),
  empty(), empty(), empty(), empty(),
  centered(trBold("ЗВІТ")),
  centered(tr("з лабораторної роботи № 1")),
  centered(tr("з дисципліни «Гіпертекст та гіпермедіа»")),
  centered(tr("на тему: «HTML5 та CSS»")),
  empty(),
  centered(tr("Варіант 9")),
  empty(), empty(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [tr("Виконав: ст. гр. ПЗПІ-25-6")],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [tr("Коновалов О. О.")],
  }),
  empty(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [tr("Перевірив: Зибіна К. В.")],
  }),
  empty(), empty(), empty(), empty(), empty(), empty(),
  centered(tr("Харків — 2026")),
];

// --- SECTION 1: МЕТА РОБОТИ ---

const section1 = [
  heading1("1", "Мета роботи"),
  body("Метою лабораторної роботи є вивчення стандарту HTML5 та CSS, способів виконання базових операцій, особливостей створення вебсторінок з використанням каскадних таблиць стилів."),
];

// --- SECTION 2: ЗАВДАННЯ ---

const section2 = [
  heading1("2", "Завдання"),
  body("Відповідно до варіанту 9 необхідно розробити вебсайт на тему «Ботанічний довідник». Завдання розподілені на три блоки:"),
  empty(),
  bodyRuns([trBold("Блок 1 "), tr("(базовий — оцінка 3):")]),
  body("1. Розробити вебсайт із трьома взаємопов’язаними сторінками з використанням елементів розмітки <header>, <main>, <footer>, <aside>, <article>, <section>."),
  body("2. Використати основні теги, кольори (тексту, фону тексту), абзаци, заголовки (H1–H6), картинки, посилання, списки (нумерований, маркований, вкладений), мітки, засоби форматування тексту, таблиці, фон сторінки (колір, картинка), підвал."),
  body("3. Включити в сайт всі відомі мета-теги."),
  body("4. Доповнити сайт єдиним стилем для всіх сторінок засобами CSS."),
  body("5. Розглянути поняття: способи впровадження CSS, селектор, клас, ID-селектори, шрифти, кольори, фон, вирівнювання, списки, кордони і рамки, курсори."),
  body("6. Нарисувати структуру вебсайту, що вийшов."),
  empty(),
  bodyRuns([trBold("Блок 2 "), tr("(розширений — оцінка 4, потрібні 3 з 4 завдань):")]),
  body("1. Доповнити сторінки елементом <nav> з горизонтальним випадаючим меню."),
  body("2. Реалізувати на сторінці фрейм <iframe>."),
  body("3. Відобразити аудіо і відео з елементами управління."),
  body("4. Створити форму з використанням елементів datetime, email, tel, url та атрибутів валідації (required, maxlength, pattern, min, max, step)."),
  empty(),
  bodyRuns([trBold("Блок 3 "), tr("(просунутий — оцінка 5, потрібне 1 з 2 завдань):")]),
  body("1. Ознайомитись із CSS-властивостями для роботи з градієнтом, тінню, прозорістю та застосувати їх."),
  body("2. Ознайомитись та застосувати універсальний селектор *, дочірній селектор та селектори псевдокласів (не менше, ніж для 3-х різних елементів)."),
];

// --- SECTION 3: ХІД РОБОТИ ---

const section3 = [
  heading1("3", "Хід роботи"),

  heading2("3.1", "Структура сайту"),
  body("Вебсайт «Ботанічний довідник» складається з трьох взаємопов’язаних HTML-сторінок, об’єднаних єдиним файлом стилів style.css. Кожна сторінка містить спільні елементи: шапку (<header>), навігаційне меню (<nav>), основний вміст (<main> з <article>), бічну панель (<aside>) та підвал (<footer>)."),
  empty(),
  body("Карта сайту:"),
  ...codeBlock(`Головна (index.html)
├── Про довідник (#about)
├── Класифікація рослин (#classification)
├── Рослина місяця (#featured)
└── Ботанічні терміни (#terms)

Рослини (plants.html)
├── Квіткові рослини (#flowering)
├── Дерева (#trees)
├── Лікарські трави (#herbs)
├── Порівняльна таблиця (#table)
└── Карта природних зон (#iframe-section)

Контакти (contact.html)
├── Медіа: аудіо + відео (#media)
└── Форма зворотного зв'язку (#contact-form)`),
  empty(),

  heading2("3.2", "Блок 1 — HTML5 та базовий CSS"),
  body("Створено три сторінки з використанням семантичних тегів HTML5: <header> для шапки сайту з назвою та підзаголовком, <nav> для навігаційного меню, <main> для основного контенту, <article> для статей, <section> для тематичних розділів, <aside> для бічної панелі з фактами та посиланнями, <footer> для підвалу з копірайтом."),
  body("На всіх сторінках використано заголовки від H1 до H6, абзаци (<p>), форматування тексту (<strong>, <em>, <u>, <mark>, <span>), нумеровані та марковані списки (включно з вкладеними), таблиці з <caption>, <thead>, <tbody>, зображення з посиланнями на Wikipedia."),
  body("Включено мета-теги: charset, viewport, description, keywords, author, robots, theme-color, Content-Language, X-UA-Compatible, og:title, og:description, og:type."),
  body("Єдиний файл style.css підключено до всіх трьох сторінок через <link rel=\"stylesheet\">. У CSS продемонстровано три способи задання стилів: зовнішній файл (style.css), внутрішні стилі (<style> в <head>) та вбудовані стилі (атрибут style=\"\"). Використано селектори тегів, класів (.bg-text, .plant-card), ідентифікаторів (#featured, #top), налаштовано шрифти, кольори, фон сторінки (колір + SVG-патерн), вирівнювання, рамки, курсори."),
  empty(),

  heading2("3.3", "Блок 2 — Навігація, медіа, форми"),
  body("Реалізовано горизонтальне випадаюче меню навігації за допомогою <nav> та CSS-hover. Меню є sticky (position: sticky) і залишається видимим при прокручуванні сторінки. При наведенні на пункт «Рослини» з’являється підменю з посиланнями на розділи."),
  body("На сторінці plants.html додано <iframe> з картою OpenStreetMap, що відображає природні зони України."),
  body("На сторінці contact.html реалізовано аудіоплеєр (<audio controls>) та відеоплеєр (<video controls>) з атрибутом poster для зображення-заставки. Обидва медіаелементи мають стандартні браузерні елементи управління."),
  body("Створено форму зворотного зв’язку з різними типами полів: text (з required та maxlength), email (з required), tel (з pattern для валідації номера), url, datetime-local (з min), number (з min, max, step), select, textarea (з required та maxlength), password (з pattern для складності пароля)."),
  empty(),

  heading2("3.4", "Блок 3 — Просунутий CSS"),
  body("Застосовано CSS-градієнти (linear-gradient) до шапки, підвалу, кнопки відправки форми та секції «Рослина місяця». Додано тіні (box-shadow) до карток рослин, бічної панелі, кнопок та відеоплеєра. Реалізовано прозорість (opacity) для зображень рослин при наведенні курсора."),
  body("Використано універсальний селектор * для скидання margin, padding та встановлення box-sizing: border-box. Застосовано дочірні селектори (.nav-menu > li > a, footer > p) та нащадкові селектори (.form-group input, article p > strong). Реалізовано псевдокласи для більш ніж 3-х елементів: :hover (посилання, картки, рядки таблиць), :visited та :active (посилання), :focus (поля форми), :invalid (поля з помилковим значенням), :nth-child (рядки таблиць), :first-child (елементи списків)."),
];

// --- SECTION 4: РЕЗУЛЬТАТИ ---

const section4 = [
  heading1("4", "Результати"),
  body("Результати виконання лабораторної роботи представлені у вигляді скріншотів вебсайту «Ботанічний довідник» у браузері Google Chrome."),
  empty(),

  imageFromFile(`${BASE}/screenshots/index-full.png`, 450, 850),
  figureCaption("4.1", "Головна сторінка (index.html)"),

  pageBreak(),
  imageFromFile(`${BASE}/screenshots/plants-full.png`, 450, 850),
  figureCaption("4.2", "Сторінка каталогу рослин (plants.html)"),

  pageBreak(),
  imageFromFile(`${BASE}/screenshots/contact-full.png`, 450, 850),
  figureCaption("4.3", "Сторінка контактів та медіа (contact.html)"),

  pageBreak(),
  imageFromFile(`${BASE}/screenshots/nav-dropdown.png`, 550, 300),
  figureCaption("4.4", "Горизонтальне випадаюче меню навігації"),
];

// --- SECTION 5: ВИСНОВКИ ---

const section5 = [
  heading1("5", "Висновки"),
  body("У ході виконання лабораторної роботи було вивчено стандарт HTML5 та каскадні таблиці стилів CSS. Розроблено вебсайт «Ботанічний довідник», що складається з трьох взаємопов’язаних сторінок з використанням семантичної розмітки, мета-тегів, таблиць, списків, форм з валідацією, медіаелементів та інтерактивної карти."),
  body("Опановано різні типи CSS-селекторів (тегів, класів, ідентифікаторів, дочірні, псевдокласи), а також CSS-ефекти: градієнти, тіні та прозорість. Сайт відповідає вимогам усіх трьох блоків завдань."),
];

// --- SECTION 6: КОНТРОЛЬНІ ПИТАННЯ ---

const section6 = [
  heading1("6", "Відповіді на контрольні питання"),

  bodyRuns([trBold("1. Що таке браузер?")]),
  body("Браузер — програмний застосунок, призначений для доступу до інформації в Інтернеті. Він інтерпретує HTML-код сторінки, знаходить теги розмітки та відображає їх у вигляді стильового оформлення змісту."),

  bodyRuns([trBold("3. Опишіть загальний синтаксис мови HTML.")]),
  body("Синтаксис HTML базується на тегах: <тег атрибут=\"значення\"> (одиночний) або <тег атрибут=\"значення\">...вміст...</тег> (парний). Теги бувають одиночні (self-closing) та парні (контейнери). Атрибути — обов’язкові та необов’язкові пари ключ/значення всередині відкриваючого тегу."),

  bodyRuns([trBold("5. Для чого потрібен тег <head>?")]),
  body("Тег <head> призначений для зберігання службових елементів документа: мета-тегів, заголовку <title>, підключення стилів <link>, скриптів <script>. Вміст <head> не відображається на сторінці."),

  bodyRuns([trBold("7. Які властивості елементу <div>?")]),
  body("<div> — універсальний блоковий контейнер. Завжди відокремлюється від інших елементів порожнім рядком, не несе семантичного навантаження (є роздільником), дозволяє застосовувати CSS-атрибути, пов’язані з межею блоків."),

  bodyRuns([trBold("9. Які існують способи використання стилів у документі?")]),
  body("1) Зовнішній файл стилів (підключається через <link rel=\"stylesheet\">). 2) Внутрішні стилі в <style> всередині <head>. 3) Вбудовані стилі через атрибут style=\"\" безпосередньо в тезі."),

  bodyRuns([trBold("11. Що таке ідентифікатор у CSS?")]),
  body("Ідентифікатор (ID-селектор) визначає унікальне ім’я елемента, яке використовується для зміни його стилю. Синтаксис: #ім’я { властивість: значення; }. В HTML задається атрибутом id=\"\". Кожен ID унікальний у межах документа."),

  bodyRuns([trBold("13. За допомогою яких тегів можна розділити сторінку на частини?")]),
  body("<header> (шапка), <nav> (навігація), <main> (основний вміст), <article> (стаття), <section> (розділ), <aside> (бічна панель), <footer> (підвал), <div> (універсальний блок)."),

  bodyRuns([trBold("15. Назвіть основні атрибути валідації форми.")]),
  body("required — обов’язкове поле; maxlength — максимальна довжина тексту; min / max — діапазон числових/дата-значень; step — крок зміни значення; pattern — регулярний вираз для перевірки формату."),

  bodyRuns([trBold("17. Для чого потрібен тег <video>? Назвіть атрибути цього тегу.")]),
  body("<video> додає, відтворює і управляє відеороликом на вебсторінці. Основні атрибути: src (шлях до файлу), controls (елементи управління), autoplay, loop, muted, poster (зображення-заставка), width, height. Вкладений тег <source> дозволяє вказати кілька форматів."),

  bodyRuns([trBold("19. Які є головні етапи створення сайту?")]),
  body("1) Визначення мети та цільової аудиторії. 2) Планування структури та карти сайту. 3) Розробка дизайну (wireframe, макет). 4) Верстка HTML/CSS. 5) Програмування функціоналу (JS, backend). 6) Наповнення контентом. 7) Тестування у різних браузерах. 8) Публікація (хостинг, домен)."),
];

// --- APPENDIX А ---

function appendixHeading(letter, title) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [new TextRun({ text: `ДОДАТОК ${letter}`, font: FONT, size: BODY_SIZE, bold: true })],
  });
}

function appendixSubtitle(title) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [tr(title)],
  });
}

const appendixA = [
  pageBreak(),
  appendixHeading("А"),
  appendixSubtitle("Вихідний код"),
  empty(),

  listingCaption("А.1", "index.html"),
  ...codeBlock(indexHtml),
  empty(),

  pageBreak(),
  listingCaption("А.2", "plants.html"),
  ...codeBlock(plantsHtml),
  empty(),

  pageBreak(),
  listingCaption("А.3", "contact.html"),
  ...codeBlock(contactHtml),
  empty(),

  pageBreak(),
  listingCaption("А.4", "style.css"),
  ...codeBlock(styleCss),
];

// --- DOCUMENT ---

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: BODY_SIZE, language: { value: "uk-UA" } },
        paragraph: { spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" } },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: { spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: { spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" }, outlineLevel: 1 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: {
            top: Math.round(20 * MM_TO_DXA),
            bottom: Math.round(20 * MM_TO_DXA),
            left: Math.round(20 * MM_TO_DXA),
            right: Math.round(20 * MM_TO_DXA),
            header: 708, footer: 708,
          },
        },
      },
      children: [...titlePage],
    },
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { ...margins, header: 708, footer: 708 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: BODY_SIZE })],
          })],
        }),
      },
      children: [
        ...section1,
        ...section2,
        ...section3,
        ...section4,
        ...section5,
        ...section6,
        ...appendixA,
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
const outPath = `${BASE}/Звіт_ЛР1_Коновалов_ПЗПІ-25-6.docx`;
writeFileSync(outPath, buffer);
console.log(`Created: ${outPath} (${(buffer.length / 1024).toFixed(0)} KB)`);
