import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const inputPath = path.join(root, "База Лехи из телефонной книги.csv");
const outputDir = path.join(root, "outputs");
const outputPath = path.join(outputDir, "lekha_contacts_structured_mvp.xlsx");

const TEST_PHONE = "+7 888 888-88-88";

const rawText = await fs.readFile(inputPath, "utf8");
const rawRows = rawText
  .replace(/^\uFEFF/, "")
  .split(/\r?\n/)
  .map((line) => line.trim().replace(/^"|"$/g, "").trim())
  .filter(Boolean);

const keywordMap = [
  ["event-агентство", ["агентств", "агентство", "ивент", "эвент", "event", "forevent", "wedding", "свадебн", "праздничн", "организатор", "ивентор", "эвентер"]],
  ["частный заказчик", ["частн", "заказчик", "невеста", "юбилей", "свадьба", "корпоратив"]],
  ["слушатель", ["слушатель", "купил", "купила", "купили", "видел", "видела"]],
  ["ресторан/площадка", ["ресторан", "бар", "клуб", "площадка", "остериа", "дом", "зал", "сдд", "спикизи"]],
  ["партнер/подрядчик", ["ведущий", "фотограф", "звукач", "звукорежиссер", "артдиректор", "арт директор", "директор", "управляющ"]],
  ["не писать", ["не писать", "не работает", "нет ватсап", "агрессия"]],
];

const stopWords = ["не писать", "не работает", "нет ватсап", "агрессия"];
const noPrefixPattern = /^(нет)\b/i;

function classify(raw) {
  const lower = raw.toLowerCase();
  const tags = keywordMap
    .filter(([, words]) => words.some((word) => lower.includes(word)))
    .map(([tag]) => tag);

  const doNotContact = stopWords.some((word) => lower.includes(word)) || noPrefixPattern.test(raw);

  let type = "неразобранный";
  if (doNotContact) type = "не писать";
  else if (/^(АГЕНТСТВО|ИВЕНТОР|ЭВЕНТЕР|ИВЕНТ|EVENT)\b/i.test(raw) || tags.includes("event-агентство")) type = "event-агентство";
  else if (/^(ЧАСТНЫЙ ЗАКАЗЧИК|ЗАКАЗЧИК|НЕВЕСТА)\b/i.test(raw) || tags.includes("частный заказчик")) type = "частный заказчик";
  else if (/^СЛУШАТЕЛЬ\b/i.test(raw) || tags.includes("слушатель")) type = "слушатель";
  else if (/^(РЕСТОРАН|КОНЦЕРТНАЯ ПЛОЩАДКА)\b/i.test(raw) || tags.includes("ресторан/площадка")) type = "ресторан/площадка";
  else if (/^(ВЕДУЩИЙ|АРТ ДИРЕКТОР)\b/i.test(raw) || tags.includes("партнер/подрядчик")) type = "партнер/подрядчик";

  const priority = type === "event-агентство" ? "A" : type === "частный заказчик" ? "B" : type === "не писать" ? "STOP" : "C";
  return { type, tags: [...new Set(tags)], doNotContact, priority };
}

function extractName(raw) {
  const cleaned = raw
    .replace(/^НЕ ПИСАТЬ\s+/i, "")
    .replace(/^НЕТ\s+/i, "")
    .replace(/^ЧАСТНЫЙ ЗАКАЗЧИК\s+/i, "")
    .replace(/^ЗАКАЗЧИК\s+/i, "")
    .replace(/^АГЕНТСТВО\s+/i, "")
    .replace(/^СЛУШАТЕЛЬ\s+/i, "")
    .replace(/^ВЕДУЩИЙ\s+/i, "")
    .replace(/^АРТ ДИРЕКТОР\s+/i, "")
    .replace(/^РЕСТОРАН\s+/i, "")
    .trim();

  const stopWords = /^(?:\d|[0-9./-]+|корпоратив|велкам|квартет|трио|свадьба|юбилей|работали|видел|видела|купил|купила|купили|от|через)$/i;
  const parts = cleaned.split(/\s+/).filter(Boolean);
  const nameParts = [];
  for (const part of parts) {
    if (nameParts.length >= 4) break;
    if (stopWords.test(part)) break;
    if (/[₽]/.test(part)) break;
    nameParts.push(part);
  }
  return nameParts.join(" ") || cleaned.slice(0, 60);
}

function compactNote(raw, type) {
  return raw
    .replace(/^НЕ ПИСАТЬ\s+/i, "")
    .replace(/^НЕТ\s+/i, "")
    .replace(/^ЧАСТНЫЙ ЗАКАЗЧИК\s+/i, "")
    .replace(/^ЗАКАЗЧИК\s+/i, "")
    .replace(/^АГЕНТСТВО\s+/i, "")
    .replace(/^СЛУШАТЕЛЬ\s+/i, "")
    .replace(/^ВЕДУЩИЙ\s+/i, "")
    .replace(/^АРТ ДИРЕКТОР\s+/i, "")
    .replace(/^РЕСТОРАН\s+/i, "")
    .replace(/\s+/g, " ")
    .trim() || type;
}

function nextAction(type, doNotContact) {
  if (doNotContact) return "не писать, проверить вручную";
  if (type === "event-агентство") return "мягкое касание: напомнить о себе, можно отправить видео";
  if (type === "частный заказчик") return "проверить контекст, писать только при уместном поводе";
  if (type === "слушатель") return "мягкий контент/приглашение, не продавать";
  if (type === "ресторан/площадка") return "поддерживать контакт, уточнить релевантность";
  if (type === "партнер/подрядчик") return "поддерживать партнерский контакт";
  return "разобрать вручную";
}

const headers = [
  "id",
  "display_name",
  "phone",
  "telegram_status",
  "contact_type",
  "priority",
  "do_not_contact",
  "tags",
  "cadence_days",
  "next_action",
  "last_touch_at",
  "next_touch_at",
  "clean_note",
  "raw_phonebook_line",
  "manual_review",
];

const contactRows = rawRows.map((raw, index) => {
  const info = classify(raw);
  return [
    index + 1,
    extractName(raw),
    TEST_PHONE,
    "не проверен",
    info.type,
    info.priority,
    info.doNotContact ? "да" : "нет",
    info.tags.join(", "),
    info.doNotContact ? "" : 30,
    nextAction(info.type, info.doNotContact),
    "",
    "",
    compactNote(raw, info.type),
    raw,
    "проверить",
  ];
});

const counts = new Map();
for (const row of contactRows) {
  counts.set(row[4], (counts.get(row[4]) || 0) + 1);
}

const workbook = Workbook.create();
const contacts = workbook.worksheets.add("Contacts");
const summary = workbook.worksheets.add("Summary");
const rules = workbook.worksheets.add("Agent rules");

function colName(n) {
  let s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}

const contactMatrix = [headers, ...contactRows];
contacts.getRange(`A1:${colName(headers.length)}${contactMatrix.length}`).values = contactMatrix;
contacts.freezePanes.freezeRows(1);

const summaryRows = [
  ["Метрика", "Значение"],
  ["Всего строк в выгрузке", contactRows.length],
  ["Тестовый номер", TEST_PHONE],
  ["Канал MVP", "Telegram"],
  ["Отправка без утверждения", "запрещена"],
  ["Лимит касания", "не чаще 1 раза в месяц на контакт"],
  ["Первый интерфейс", "Telegram-бот + песочница"],
  ["", ""],
  ["Тип контакта", "Количество"],
  ...[...counts.entries()].sort((a, b) => b[1] - a[1]),
];
summary.getRange(`A1:B${summaryRows.length}`).values = summaryRows;

const ruleRows = [
  ["Правило", "Описание"],
  ["Без автоотправки", "Сообщения не уходят, пока Лёха не утвердит текст."],
  ["Без прямой продажи", "Агент напоминает о себе и поддерживает контакт, но не закрывает сделку."],
  ["Без мата и хамства", "Запрещены мат, оскорбления, агрессия и токсичный юмор."],
  ["Горячий лид", "Дата, цена, состав, райдер или интерес к бэнду — сразу уведомить Лёху."],
  ["Голосовые правки", "Лёха может голосом объяснить, что изменить в тексте или правиле."],
  ["Песочница", "Спорные формулировки тестируются до реального контакта."],
  ["Материалы", "Видео с концертов хранить как Telegram-файлы/материалы для отправки."],
];
rules.getRange(`A1:B${ruleRows.length}`).values = ruleRows;

for (const [sheet, widths] of [
  [contacts, [8, 28, 18, 16, 18, 10, 14, 32, 12, 42, 14, 14, 54, 64, 16]],
  [summary, [32, 24]],
  [rules, [24, 70]],
]) {
  widths.forEach((width, index) => {
    sheet.getRange(`${colName(index + 1)}:${colName(index + 1)}`).format.columnWidth = width;
  });
  sheet.getUsedRange().format.wrapText = true;
}

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(outputPath);
