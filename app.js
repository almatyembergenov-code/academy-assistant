const reasonCodes = [
  {code:"U01", reason:"Больничный", type:"Уважительная"},
  {code:"U02", reason:"Отпуск", type:"Уважительная"},
  {code:"U03", reason:"Командировка", type:"Уважительная"},
  {code:"U04", reason:"Обучение", type:"Уважительная"},
  {code:"U05", reason:"Отгул", type:"Уважительная"},
  {code:"U99", reason:"Другая уважительная", type:"Уважительная"},
  {code:"N01", reason:"Не вышел на связь", type:"Неуважительная"},
  {code:"N02", reason:"Забыл / не явился", type:"Неуважительная"},
  {code:"N03", reason:"Отказался проходить", type:"Неуважительная"},
  {code:"N04", reason:"Без объяснения причины", type:"Неуважительная"}
];

const violationCodes = [
  {code:"M01", violation:"Выключил камеру"},
  {code:"M02", violation:"Выключил звук без разрешения"},
  {code:"M03", violation:"Получал подсказки"},
  {code:"M04", violation:"Разговаривал с коллегами"},
  {code:"M05", violation:"Фото/скриншот вопросов"},
  {code:"M06", violation:"Некорректно указал ФИО/город/подразделение"},
  {code:"M07", violation:"Проходил вне графика / без видеоконтроля"},
  {code:"M08", violation:"Неоднократные замечания"},
  {code:"M09", violation:"Иное нарушение"}
];

let notStartedRows = [];
let trainerRows = [];
let finalRows = [];
let historyRows = [];
let summaryRows = [];
let manualCheckRows = [];

function renderDictionaries() {
  renderTable("reasonCodesTable", reasonCodes, ["code","reason","type"]);
  renderTable("violationCodesTable", violationCodes, ["code","violation"]);
}

function normalizeHeader(h) {
  return String(h || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function pick(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const found = keys.find(k => normalizeHeader(k) === normalizeHeader(c));
    if (found !== undefined) return row[found];
  }
  const normalized = keys.map(k => [k, normalizeHeader(k)]);
  for (const c of candidates) {
    const nc = normalizeHeader(c);
    const found = normalized.find(([k,n]) => n.includes(nc) || nc.includes(n));
    if (found) return row[found[0]];
  }
  return "";
}

async function readWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, {type:"array"});
  const first = wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[first], {defval:""});
}

function statusText(id, text) {
  document.getElementById(id).textContent = text;
}

function renderTable(id, rows, columns = null) {
  const table = document.getElementById(id);
  if (!rows || rows.length === 0) {
    table.innerHTML = "<tr><td>Нет данных</td></tr>";
    return;
  }
  const cols = columns || Object.keys(rows[0]);
  table.innerHTML = "<thead><tr>" + cols.map(c => `<th>${escapeHtml(c)}</th>`).join("") + "</tr></thead>" +
    "<tbody>" + rows.slice(0, 300).map(r => "<tr>" + cols.map(c => `<td>${escapeHtml(r[c] ?? "")}</td>`).join("") + "</tr>").join("") + "</tbody>";
}

function escapeHtml(v) {
  return String(v ?? "").replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

async function prepareTrainerFile() {
  const file = document.getElementById("notStartedFile").files[0];
  if (!file) return alert("Сначала загрузи файл “Не начат”.");

  const rows = await readWorkbookFromFile(file);
  notStartedRows = rows.map(r => ({
    "ФИО": pick(r, ["Имя пользователя", "ФИО", "Пользователь", "Сотрудник"]),
    "Город": pick(r, ["Город", "Регион", "Подразделение"]),
    "Подразделение": pick(r, ["Подразделение", "Группа", "Канал"]),
    "Должность": pick(r, ["Должность"]),
    "Офис": pick(r, ["Офис"]),
    "Код причины": "",
    "Комментарий тренера": ""
  })).filter(r => r["ФИО"]);

  statusText("notStartedInfo", `Загружено строк: ${rows.length}. Подготовлено для тренеров: ${notStartedRows.length}.`);
  renderTable("notStartedPreview", notStartedRows);
}

function downloadTrainerTemplate() {
  if (!notStartedRows.length) return alert("Сначала подготовь файл для тренеров.");
  downloadXlsx("Файл_для_тренеров_причины_пропусков.xlsx", {
    "Причины пропусков": notStartedRows,
    "Справочник причин": reasonCodes
  });
}

async function mergeTrainerFiles() {
  const files = Array.from(document.getElementById("trainerFiles").files || []);
  if (!files.length) return alert("Загрузи хотя бы один файл от тренера.");

  trainerRows = [];
  for (const file of files) {
    const rows = await readWorkbookFromFile(file);
    for (const r of rows) {
      const code = String(pick(r, ["Код причины", "Причина", "Код"]) || "").trim().toUpperCase();
      const dict = reasonCodes.find(x => x.code === code);
      trainerRows.push({
        "Источник файла": file.name,
        "ФИО": pick(r, ["ФИО", "Имя пользователя", "Пользователь", "Сотрудник"]),
        "Город": pick(r, ["Город", "Регион"]),
        "Подразделение": pick(r, ["Подразделение", "Канал"]),
        "Код причины": code,
        "Причина": dict ? dict.reason : "",
        "Тип причины": dict ? dict.type : "Требует проверки",
        "Комментарий тренера": pick(r, ["Комментарий тренера", "Комментарий", "Примечание"])
      });
    }
  }
  trainerRows = trainerRows.filter(r => r["ФИО"]);
  statusText("trainerInfo", `Объединено файлов: ${files.length}. Строк с ФИО: ${trainerRows.length}.`);
  renderTable("trainerPreview", trainerRows);
}

async function loadFinalResults() {
  const file = document.getElementById("resultsFile").files[0];
  if (!file) return alert("Загрузи итоговую выгрузку iSpring.");
  const rows = await readWorkbookFromFile(file);
  finalRows = rows.map(r => ({
    "ФИО": pick(r, ["Имя пользователя", "ФИО", "Пользователь", "Сотрудник"]),
    "Город": pick(r, ["Город", "Регион", "Подразделение"]),
    "Подразделение": pick(r, ["Подразделение", "Группа", "Канал"]),
    "Статус": pick(r, ["Статус"]),
    "Результат": parsePercent(pick(r, ["Результат", "Оценка", "Баллы", "Процент"])),
    "Офис": pick(r, ["Офис"])
  })).filter(r => r["ФИО"]);
  statusText("resultsInfo", `Загружено итоговых строк: ${finalRows.length}.`);
}

async function loadHistory() {
  const files = Array.from(document.getElementById("historyFiles").files || []);
  if (!files.length) return alert("Загрузи 1–3 файла истории.");
  historyRows = [];
  for (const file of files) {
    const rows = await readWorkbookFromFile(file);
    rows.forEach(r => historyRows.push({
      "Период/файл": file.name,
      "ФИО": pick(r, ["Имя пользователя", "ФИО", "Пользователь", "Сотрудник"]),
      "Статус": pick(r, ["Статус"]),
      "Результат": parsePercent(pick(r, ["Результат", "Оценка", "Баллы", "Процент"])),
      "Комментарий": pick(r, ["Комментарий", "Примечание", "Комментарий тренера"]),
      "Тип результата": pick(r, ["Тип результата"])
    }));
  }
  historyRows = historyRows.filter(r => r["ФИО"]);
  statusText("historyInfo", `Загружено исторических строк: ${historyRows.length} из ${files.length} файлов.`);
}

function parsePercent(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v <= 1 ? Math.round(v * 100) : Math.round(v);
  const s = String(v).replace("%","").replace(",",".").trim();
  const n = Number(s);
  if (Number.isNaN(n)) return null;
  return n <= 1 ? Math.round(n * 100) : Math.round(n);
}

function calculateSummary() {
  if (!finalRows.length && !trainerRows.length) return alert("Нужны итоговые результаты или файлы от тренеров.");

  const byNameReason = new Map();
  trainerRows.forEach(r => {
    if (r["ФИО"]) byNameReason.set(normalizeName(r["ФИО"]), r);
  });

  const all = finalRows.length ? finalRows.map(r => ({...r})) : notStartedRows.map(r => ({...r, "Статус":"Не начат", "Результат": null}));

  manualCheckRows = [];

  for (const row of all) {
    const reason = byNameReason.get(normalizeName(row["ФИО"]));
    row["Код причины"] = reason?.["Код причины"] || "";
    row["Причина"] = reason?.["Причина"] || "";
    row["Тип причины"] = reason?.["Тип причины"] || "";

    if (row["Тип причины"] === "Неуважительная") {
      row["Итоговый результат"] = 0;
      row["Тип результата"] = "0% за неуважительный пропуск";
      row["Итоговый комментарий"] = "Неявка без уважительной причины. Результат обнулен.";
    } else if (row["Тип причины"] === "Уважительная") {
      const hist = historyRows.filter(h => normalizeName(h["ФИО"]) === normalizeName(row["ФИО"]));
      const valid = hist.map(h => h["Результат"]).filter(x => typeof x === "number" && x > 0);
      if (valid.length >= 3) {
        const avg = Math.round(valid.slice(0,3).reduce((a,b)=>a+b,0) / 3);
        row["Итоговый результат"] = avg;
        row["Тип результата"] = "Средний по 3 последним";
        row["Итоговый комментарий"] = `Уважительная причина. Выставлен средний результат по 3 последним тестированиям: ${avg}%.`;
      } else if (hist.length === 0) {
        row["Итоговый результат"] = 80;
        row["Тип результата"] = "80% первое тестирование";
        row["Итоговый комментарий"] = "Первое тестирование / история не найдена. Установлен минимальный проходной результат 80%.";
      } else {
        row["Итоговый результат"] = "";
        row["Тип результата"] = "Требуется ручная проверка";
        row["Итоговый комментарий"] = "Недостаточно корректных исторических результатов. Проверь комментарии в iSpring.";
        manualCheckRows.push(row);
      }
    } else {
      row["Итоговый результат"] = row["Результат"];
      row["Тип результата"] = "Фактический результат";
      row["Итоговый комментарий"] = "";
    }
  }

  const avg = average(all.map(r => r["Итоговый результат"]).filter(x => typeof x === "number"));
  const participated = all.filter(r => String(r["Статус"]).toLowerCase().includes("прой") || typeof r["Результат"] === "number").length;
  const absent = all.filter(r => r["Код причины"]).length;
  const violations = all.filter(r => String(r["Код нарушения"] || "").startsWith("M")).length;

  document.getElementById("kpiAvg").textContent = avg === null ? "—" : `${avg}%`;
  document.getElementById("kpiParticipated").textContent = participated;
  document.getElementById("kpiAbsent").textContent = absent;
  document.getElementById("kpiViolations").textContent = violations;

  summaryRows = buildGroupSummary(all);
  renderTable("summaryTable", summaryRows);
  renderTable("manualCheckTable", manualCheckRows.length ? manualCheckRows : [{"Статус":"Нет строк для ручной проверки"}]);

  window.calculatedRows = all;
}

function buildGroupSummary(rows) {
  const groups = new Map();
  for (const r of rows) {
    const key = `${r["Город"] || "Не указан"} | ${r["Подразделение"] || "Не указано"}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }
  return Array.from(groups.entries()).map(([key, arr]) => {
    const [city, unit] = key.split(" | ");
    const scores = arr.map(r => r["Итоговый результат"]).filter(x => typeof x === "number");
    return {
      "Город/регион": city,
      "Подразделение": unit,
      "Количество": arr.length,
      "Средний результат": average(scores) === null ? "" : `${average(scores)}%`,
      "80% и выше": scores.filter(x => x >= 80).length,
      "Ниже 80%": scores.filter(x => x < 80).length,
      "Уважительные пропуски": arr.filter(x => x["Тип причины"] === "Уважительная").length,
      "Неуважительные пропуски": arr.filter(x => x["Тип причины"] === "Неуважительная").length
    };
  });
}

function average(nums) {
  if (!nums.length) return null;
  return Math.round(nums.reduce((a,b)=>a+b,0) / nums.length);
}

function normalizeName(name) {
  return String(name || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function downloadFinalWorkbook() {
  const rows = window.calculatedRows || [];
  if (!rows.length) return alert("Сначала рассчитай MVP-свод.");
  downloadXlsx("B2C_Test_Control_Итоговый_MVP.xlsx", {
    "Итоговые данные": rows,
    "Свод": summaryRows,
    "Ручная проверка": manualCheckRows,
    "Коды причин": reasonCodes,
    "Коды нарушений": violationCodes
  });
}

function downloadXlsx(filename, sheets) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0,31));
  });
  XLSX.writeFile(wb, filename);
}

renderDictionaries();
