import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Upload, Lock, Trophy, Users, ChevronRight, Trash2, AlertCircle, CheckCircle2, Shield, History, Search, X, Table, Pencil, Check, Sparkles } from "lucide-react";
import { storageGet, storageSet, storageDelete } from "./storage";

const LOGO_DATA_URI = "data:image/png;base64,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";

// Фирменные цвета "Водокачки"
const BRAND_BLUE = "#1C378F";
const BRAND_RED = "#E8394E";

const ADMIN_CODE = "vodokachka2026"; // поменяйте на свой код доступа
const MIN_RATING = 25; // минимальный % за турнир — ниже этого не опускаемся

const FORMATS = [
  { id: "solo", label: "Соло" },
  { id: "pair", label: "Пара" },
  { id: "retro", label: "Ретро" },
];

const BLOCK_TITLE_TO_CATEGORY = {
  "ЛИГА ЧЕМПИОНОВ — ВЕРХНЯЯ СЕТКА": "top",
  "ЛИГА ЧЕМПИОНОВ — НИЖНЯЯ СЕТКА": "mid",
  "ЛИГА ЕВРОПЫ": "low",
  "ЛИГА КОНФЕРЕНЦИЙ": "intertoto",
  // Более ранние турниры современного формата называли сетки проще, без приставки
  // "ЛИГА ЧЕМПИОНОВ — " (например, ЧВ27) — просто дополнительные варианты названий той же
  // структуры, старые не трогаем.
  "ВЕРХНЯЯ СЕТКА": "top",
  "НИЖНЯЯ СЕТКА": "mid",
  "КУБОК ИНТЕРТОТО": "intertoto",
};

const STAGE_PREFIXES = [
  { prefix: "ЛЧ ВС", category: "top" },
  { prefix: "ЛЧ НС", category: "mid" },
  { prefix: "ЛЕ", category: "low" },
  { prefix: "ЛК", category: "intertoto" },
];

function sheetToRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

// Как sheetToRows, но "чистит" объединённые ячейки: SheetJS иногда размножает значение
// объединённой ячейки на все строки/столбцы, которые она перекрывает (а не только на
// главную, левую верхнюю) — из-за этого декоративная строка под матчем (где выведено
// название клуба, а голы объединены с реальной строкой матча выше) может ошибочно
// выглядеть как ещё один сыгранный матч. Используется только в разборе старых форматов
// для довнесения статистики — на "живой" импорт турниров не влияет.
function sheetToRowsNoMergeDupes(ws) {
  const rows = sheetToRows(ws);
  if (!ws["!merges"] || !ws["!ref"]) return rows;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  ws["!merges"].forEach((merge) => {
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        if (r === merge.s.r && c === merge.s.c) continue; // главная ячейка — не трогаем
        const arrR = r - range.s.r;
        const arrC = c - range.s.c;
        if (rows[arrR] && arrC >= 0) rows[arrR][arrC] = null;
      }
    }
  });
  return rows;
}

function findHeaderRow(rows, required) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    if (required.every((h) => row.includes(h))) return i;
  }
  return -1;
}

function indexHeaders(row) {
  const col = {};
  row.forEach((h, i) => { if (h) col[h] = i; });
  return col;
}

// Название сетки в файле часто написано КАПСОМ ("ЛИГА ЧЕМПИОНОВ — ВЕРХНЯЯ СЕТКА") —
// для показа приводим к normal case ("Лига чемпионов — Верхняя сетка"), не трогая
// произвольные названия, которые уже не капсом (например, вписанные админом вручную).
function prettyBracketName(raw) {
  if (!raw) return raw;
  const str = String(raw);
  if (str !== str.toUpperCase() || str === str.toLowerCase()) return str;
  return str
    .split(" — ")
    .map((seg) => {
      const lower = seg.toLowerCase();
      return lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join(" — ");
}

function stageCategory(stage) {
  const s = String(stage);
  for (const { prefix, category } of STAGE_PREFIXES) {
    if (s.startsWith(prefix)) return category;
  }
  return null;
}

function newRecord(name) {
  return {
    name,
    played: 0,
    groupWins: 0,
    groupDraws: 0,
    groupLosses: 0,
    groupGoalsFor: 0,
    groupGoalsAgainst: 0,
    playoffWins: 0,
    playoffLosses: 0,
    playoffGoalsFor: 0,
    playoffGoalsAgainst: 0,
    intertotoWins: 0,
    poMatches: 0,
    semi: { top: false, mid: false, low: false },
    final: { top: false, mid: false, low: false },
    reachedLCH: false,
    bracketStage: {}, // название сетки (как в файле, дословно) -> "semi" | "final" | "champion"
  };
}

function readSettings(wb) {
  const ws = wb.Sheets["0_Настройки"];
  if (!ws) return { G: null, vsSize: null };
  const rows = sheetToRows(ws);
  let G = null, vsSize = null;
  // Ищем ячейку-подпись в любой колонке (не полагаемся на номер столбца —
  // он "плывёт", если колонка A на листе полностью пустая) и берём значение
  // из следующей ячейки той же строки.
  rows.forEach((row) => {
    row.forEach((cell, idx) => {
      const label = String(cell || "").trim();
      if (!label) return;
      if (label === "кол-во матчей") {
        const val = row[idx + 1];
        if (val !== null && val !== undefined && val !== "") G = Number(val);
      }
      if (label.includes("Верхней сетке")) {
        const val = row[idx + 1];
        if (val !== null && val !== undefined && val !== "") vsSize = Number(val);
      }
    });
  });
  return { G, vsSize };
}

function computeGroupStats(wb, players) {
  const ws = wb.Sheets["3_Таблица"];
  if (!ws) throw new Error('Не найден лист "3_Таблица"');
  const rows = sheetToRows(ws);
  const headerIdx = findHeaderRow(rows, ["Участник"]);
  if (headerIdx === -1) throw new Error('На листе "3_Таблица" не найдена строка заголовков');
  const col = indexHeaders(rows[headerIdx]);
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const name = normalizeName(row[col["Участник"]]);
    if (!name) continue;
    const p = players[name] || (players[name] = newRecord(name));
    p.groupWins = Number(row[col["В"]]) || 0;
    p.groupDraws = Number(row[col["Н"]]) || 0;
    p.groupLosses = Number(row[col["П"]]) || 0;
    p.groupGoalsFor = Number(row[col["МЗ"]]) || 0;
    p.groupGoalsAgainst = Number(row[col["МП"]]) || 0;
  }
}

function computePlayoffStats(wb, players) {
  const ws = wb.Sheets["4_Плей-офф расписание"];
  if (!ws) throw new Error('Не найден лист "4_Плей-офф расписание"');
  const rows = sheetToRows(ws);
  const headerIdx = findHeaderRow(rows, ["Этап", "Победитель"]);
  if (headerIdx === -1) throw new Error('На листе "4_Плей-офф расписание" не найдена строка заголовков');
  const col = indexHeaders(rows[headerIdx]);
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const stage = row[col["Этап"]];
    const status = row[col["Статус"]];
    if (!stage) continue;
    const played = status && String(status).includes("Сыгран");
    const p1 = normalizeName(row[col["Участник 1"]]);
    const p2 = normalizeName(row[col["Участник 2"]]);
    const winner = normalizeName(row[col["Победитель"]]);
    const g1 = Number(row[col["Г1"]]);
    const g2 = Number(row[col["Г2"]]);
    const hasScore = !Number.isNaN(g1) && !Number.isNaN(g2);
    if (played) {
      if (p1) {
        const rec = players[p1] || (players[p1] = newRecord(p1));
        rec.poMatches++;
        if (hasScore) { rec.playoffGoalsFor += g1; rec.playoffGoalsAgainst += g2; }
      }
      if (p2) {
        const rec = players[p2] || (players[p2] = newRecord(p2));
        rec.poMatches++;
        if (hasScore) { rec.playoffGoalsFor += g2; rec.playoffGoalsAgainst += g1; }
      }
    }
    if (played && winner) {
      const category = stageCategory(stage);
      const p = players[winner] || (players[winner] = newRecord(winner));
      // Любая победа в плей-офф считается победой (3 очка), кроме интертото — у него свой вес.
      // Категория (top/mid/low) нужна только для бонуса за полуфинал/финал (см. computeBracket) —
      // не для самого факта победы: некоторые этапы (например "3М" — матч за 3-е место)
      // не попадают ни в один известный бонусный блок, но победа в них всё равно победа.
      if (category === "intertoto") p.intertotoWins++;
      else p.playoffWins++;
      const loser = winner === p1 ? p2 : winner === p2 ? p1 : null;
      if (loser) (players[loser] || (players[loser] = newRecord(loser))).playoffLosses++;
    }
  }
}

function computeBracket(wb, players, bracketOrder) {
  const ws = wb.Sheets["5_Сетка плей-офф"];
  if (!ws) throw new Error('Не найден лист "5_Сетка плей-офф"');
  const data = sheetToRows(ws);
  let i = 0;
  while (i < data.length) {
    const rawTitle = String((data[i] || [])[1] || "").trim();
    const category = BLOCK_TITLE_TO_CATEGORY[rawTitle];
    if (category) {
      const title = prettyBracketName(rawTitle);
      if (bracketOrder && !bracketOrder.includes(title)) bracketOrder.push(title);
      const headerIdx = i + 1;
      const headerRow = data[headerIdx] || [];
      const populated = [];
      headerRow.forEach((v, idx) => { if (v && idx > 0) populated.push(idx); });
      populated.sort((a, b) => a - b);
      const finalCol = populated.length >= 1 ? populated[populated.length - 1] : null;
      const semiCol = populated.length >= 2 ? populated[populated.length - 2] : null;

      let r = headerIdx + 1;
      const allNames = new Set();
      const semiNames = new Set();
      const finalNames = new Set();
      let champion = null;
      let pastThirdPlace = false; // ниже подписи "N место" — уже не настоящий финал, а матч за 3-е место
      let thirdPlaceResolved = false; // нашли победителя матча за 3-е место — дальше в этой же колонке
      const thirdPlaceMatch = []; // { name, isWinner } — если в сетке есть отдельный матч за 3-е место
      while (r < data.length) {
        const row = data[r] || [];
        const rowTitle = String(row[1] || "").trim();
        if (BLOCK_TITLE_TO_CATEGORY[rowTitle]) break;
        row.forEach((v, idx) => {
          if (idx > 0 && v) {
            const raw = String(v).trim();
            if (/^\d+\s*место$/i.test(raw)) return; // подпись "3 место" — не имя игрока
            const cleanName = normalizeName(raw.replace("🏆", "").trim());
            if (cleanName) allNames.add(cleanName);
          }
        });
        if (semiCol !== null && row[semiCol]) semiNames.add(normalizeName(String(row[semiCol]).trim()));
        if (finalCol !== null && row[finalCol]) {
          const raw = String(row[finalCol]).trim();
          if (/^\d+\s*место$/i.test(raw)) {
            pastThirdPlace = true; // всё ниже в этой колонке — уже матч за 3-е место
          } else if (!pastThirdPlace) {
            if (raw.startsWith("🏆")) champion = normalizeName(raw.replace("🏆", "").trim());
            else finalNames.add(normalizeName(raw));
          } else {
            // Матч за 3-е место: победитель (с 🏆) получает титул "3 место",
            // проигравший — "4 место" (а не просто "Полуфиналист"). Сразу после того как
            // нашли победителя — останавливаемся: в паре файлов ниже в той же колонке без
            // предупреждения начинается СОВСЕМ ДРУГОЙ маленький турнир (тот же столбец
            // используется повторно), и без этой остановки его участники ошибочно
            // приписываются к текущей сетке.
            if (raw.startsWith("🏆")) {
              thirdPlaceMatch.push({ name: normalizeName(raw.replace("🏆", "").trim()), isWinner: true });
              thirdPlaceResolved = true;
            } else {
              thirdPlaceMatch.push({ name: normalizeName(raw), isWinner: false });
            }
          }
        }
        if (thirdPlaceResolved) { r++; break; }
        r++;
      }
      if (champion) finalNames.add(champion);

      allNames.forEach((name) => {
        const p = players[name] || (players[name] = newRecord(name));
        if (category === "top" || category === "mid") p.reachedLCH = true;
      });
      semiNames.forEach((name) => {
        const p = players[name] || (players[name] = newRecord(name));
        if (category !== "intertoto") p.semi[category] = true;
        if (!p.bracketStage[title]) p.bracketStage[title] = "semi";
      });
      finalNames.forEach((name) => {
        const p = players[name] || (players[name] = newRecord(name));
        if (category !== "intertoto") p.final[category] = true;
        p.bracketStage[title] = "final";
      });
      if (champion) {
        const p = players[champion] || (players[champion] = newRecord(champion));
        p.bracketStage[title] = "champion";
      }
      // Матч за 3-е место — определяет титул точнее, чем общий "Полуфиналист", поэтому
      // переопределяем bracketStage напрямую (не через обычное "только если выше рангом",
      // а безусловно — это уточнение, а не понижение).
      thirdPlaceMatch.forEach(({ name, isWinner }) => {
        const p = players[name] || (players[name] = newRecord(name));
        p.bracketStage[title] = isWinner ? "third" : "fourth";
      });

      i = r;
      continue;
    }
    i++;
  }
}

function computeO(p) {
  let o = 4;
  o += p.groupWins * 2;
  o += p.groupDraws * 1;
  o += p.playoffWins * 3;
  o += p.intertotoWins * 1.25;
  if (p.semi.top) o += 2;
  if (p.final.top) o += 2;
  if (p.semi.mid) o += 1.5;
  if (p.final.mid) o += 1.5;
  if (p.semi.low) o += 1;
  if (p.final.low) o += 1;
  if (p.reachedLCH) o += 4;
  return o;
}

// ---------- ИСТОРИЧЕСКИЕ ФОРМАТЫ (для довнесения статистики в уже импортированные турниры) ----------

// Формат A: листы "Ввод" (состав игрок/команда + матчи группы), "Плей-офф" (все сетки,
// Общие помощники для "командных" исторических форматов: собрать все матчи со счётом
// из листа (по названиям столбцов, независимо от того, сколько раз заголовок повторяется
// на листе — секции/раунды пропускаются сами, там просто нет чисел в нужных колонках),
// и свернуть их в статистику по каждой команде.
function collectScoredMatches(rows, col, headers) {
  const h = headers || { t1: "Команда 1", g1: "Гол 1", g2: "Гол 2", t2: "Команда 2" };
  const matches = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const t1 = row[col[h.t1]];
    const g1 = row[col[h.g1]];
    const g2 = row[col[h.g2]];
    const t2 = row[col[h.t2]];
    if (t1 && t2 && typeof g1 === "number" && typeof g2 === "number") {
      matches.push([String(t1).trim(), g1, g2, String(t2).trim()]);
    }
  }
  return matches;
}

function tallyTeamStats(matches) {
  const teamStats = {};
  const ensure = (t) => (teamStats[t] = teamStats[t] || { played: 0, wins: 0, draws: 0, losses: 0, gf: 0, ga: 0 });
  matches.forEach(([t1, g1, g2, t2]) => {
    const s1 = ensure(t1), s2 = ensure(t2);
    s1.played++; s2.played++;
    s1.gf += g1; s1.ga += g2;
    s2.gf += g2; s2.ga += g1;
    if (g1 > g2) { s1.wins++; s2.losses++; }
    else if (g1 < g2) { s2.wins++; s1.losses++; }
    else { s1.draws++; s2.draws++; }
  });
  return teamStats;
}

function buildLegacyRows(teamStats, playerOfTeam) {
  const rows = Object.keys(teamStats).map((team) => {
    const s = teamStats[team];
    const name = playerOfTeam[team] || null;
    return {
      name: name || `⚠️ ${team}`,
      team,
      unresolved: !name,
      played: s.played, wins: s.wins, draws: s.draws, losses: s.losses,
      goalsFor: s.gf, goalsAgainst: s.ga,
    };
  });
  rows.sort((a, b) => a.name.localeCompare(b.name, "ru"));
  return rows;
}

// Формат A: листы "Ввод" (состав игрок/команда + матчи группы), "Плей-офф" (все сетки,
// каждый раунд — 2 ноги). Игроки закреплены за клубами, а не фигурируют в матчах напрямую.
// Мы НЕ считаем очки/рейтинг для этого формата — только сырые В/Н/П/голы по каждому матчу
// (обе ноги считаются отдельными матчами), которые потом накладываются на уже существующий
// турнир (у него % уже посчитан через обычный импорт истории).
// Разбирает структуру плей-офф форматов A и B: секции по имени ("Верхняя сетка"/"Средняя
// сетка"/"Нижняя сетка"/"Финал Интертото" и т.п.) с раундами, подписанными текстом
// ("1/4 финала"/"1/2 финала"/"Финал") — определяем стадию по префиксу подписи раунда,
// а не по положению колонок (как в современном формате), потому что тут нет визуальной
// сетки-таблицы, только построчное расписание. Одна и та же секция может встречаться
// несколько раз подряд (например, "Верхняя сетка" отдельным блоком под 1/4 финала и
// отдельным блоком под 1/2+финал) — это ничего не ломает, просто добавляется в тот же бренд.
function parseTextRoundBracketStages(ws, playerOfTeam) {
  const rowsPo = sheetToRows(ws);
  // "Раунд"/название секции лежат в столбце B исходника — если колонка A на листе пустая,
  // SheetJS обрезает её и все позиции "уезжают" на один влево (та же ловушка, что уже
  // встречалась с листом "0_Настройки" в самом начале). Считаем реальное смещение через !ref.
  const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { c: 0 } };
  const labelCol = 1 - range.s.c;

  const bracketOrder = [];
  const stageByTeam = {};
  const ensureTeam = (t) => (stageByTeam[t] = stageByTeam[t] || {});
  const setStage = (obj, key, stage) => {
    const rank = STAGE_RANK[stage];
    if (!obj[key] || STAGE_RANK[obj[key]] < rank) obj[key] = stage;
  };

  let i = 0;
  while (i < rowsPo.length) {
    const row = rowsPo[i] || [];
    const label = String(row[labelCol] || "").trim();
    const nextRow = rowsPo[i + 1] || [];
    const isSectionHeader = label && String(nextRow[labelCol] || "").trim() === "Раунд";
    if (isSectionHeader) {
      const sectionName = label;
      if (!bracketOrder.includes(sectionName)) bracketOrder.push(sectionName);
      const col = indexHeaders(nextRow);
      let r = i + 2;
      while (r < rowsPo.length) {
        const dataRow = rowsPo[r] || [];
        const roundLabel = String(dataRow[labelCol] || "").trim();
        const ahead = rowsPo[r + 1] || [];
        if (roundLabel && String(ahead[labelCol] || "").trim() === "Раунд") break; // это уже заголовок следующей секции
        if (!roundLabel) { r++; continue; }
        const t1 = dataRow[col["Команда 1"]];
        const t2 = dataRow[col["Команда 2"]];
        const g1 = dataRow[col["Гол 1"]];
        const g2 = dataRow[col["Гол 2"]];
        const winner = dataRow[col["Победитель"]];
        if (typeof g1 === "number" && typeof g2 === "number" && t1 && t2) {
          let stage = null;
          if (roundLabel.startsWith("Финал")) stage = "final";
          else if (roundLabel.startsWith("1/2 финала")) stage = "semi";
          if (stage) {
            [t1, t2].forEach((t) => setStage(ensureTeam(String(t).trim()), sectionName, stage));
            if (stage === "final" && winner) setStage(ensureTeam(String(winner).trim()), sectionName, "champion");
          }
        }
        r++;
      }
      i = r;
      continue;
    }
    i++;
  }

  const playerStage = {};
  Object.entries(stageByTeam).forEach(([team, stages]) => {
    const player = playerOfTeam[team];
    if (!player) return;
    const ps = (playerStage[player] = playerStage[player] || {});
    Object.entries(stages).forEach(([bracket, stage]) => setStage(ps, bracket, stage));
  });

  // "Финал средней и нижней сеток" — решающий матч между чемпионами средней и нижней
  // сетки, почётнее, чем просто "чемпион средней/нижней сетки" по отдельности — поднимаем
  // его в списке сразу после верхней сетки, чтобы тай-брейк по порядку сеток (при равной
  // стадии "Победитель") отдавал приоритет именно ему.
  const specialIdx = bracketOrder.findIndex((b) => b.toLowerCase().includes("финал средней и нижней"));
  if (specialIdx > 1) {
    const [special] = bracketOrder.splice(specialIdx, 1);
    bracketOrder.splice(1, 0, special);
  }

  // На некоторых файлах в той же позиции, где обычно название секции, случайно
  // оказывается декоративная подпись раунда ("1/4 финала" и т.п.) — это не ломает сами
  // титулы (внутри неё нет матчей 1/2 финала или финала, значит и стадия не назначится),
  // но чтобы не засорять список сеток для админки — оставляем только те, где реально
  // кто-то дошёл хотя бы до полуфинала.
  const usedBrackets = new Set();
  Object.values(stageByTeam).forEach((stages) => Object.keys(stages).forEach((b) => usedBrackets.add(b)));
  const cleanBracketOrder = bracketOrder.filter((b) => usedBrackets.has(b));

  return { playerStage, bracketOrder: cleanBracketOrder };
}

function parseLegacyFormatA(wb) {
  const wsIn = wb.Sheets["Ввод"];
  const wsPo = wb.Sheets["Плей-офф"];
  if (!wsIn) throw new Error('Не найден лист "Ввод"');
  if (!wsPo) throw new Error('Не найден лист "Плей-офф"');

  const rowsIn = sheetToRows(wsIn);

  const matchHeaderIdx = findHeaderRow(rowsIn, ["Команда 1", "Гол 1", "Гол 2", "Команда 2"]);
  if (matchHeaderIdx === -1) throw new Error('На листе "Ввод" не найдена строка заголовков матчей');
  const matchCol = indexHeaders(rowsIn[matchHeaderIdx]);

  let playerHeaderIdx = -1;
  for (let i = 0; i < rowsIn.length; i++) {
    if (rowsIn[i][0] === "№" && rowsIn[i][1] === "Команда") { playerHeaderIdx = i; break; }
  }
  if (playerHeaderIdx === -1) throw new Error('На листе "Ввод" не найдена строка заголовков состава (№/Команда)');

  const playerOfTeam = {};
  const skippedTeams = [];
  for (let r = playerHeaderIdx + 1; r < rowsIn.length; r++) {
    const row = rowsIn[r];
    const name = row[0];
    const team = row[1] ? String(row[1]).trim() : row[1];
    if (!team) continue;
    if (!name || String(name).trim() === "-" || String(name).trim() === "") {
      skippedTeams.push(team);
      continue;
    }
    playerOfTeam[team] = normalizeName(String(name).trim());
  }

  let matches = collectScoredMatches(rowsIn, matchCol);

  const rowsPo = sheetToRows(wsPo);
  const poHeaderIdx = findHeaderRow(rowsPo, ["Команда 1", "Гол 1", "Гол 2", "Команда 2"]);
  if (poHeaderIdx === -1) throw new Error('На листе "Плей-офф" не найдена строка заголовков матчей');
  const poCol = indexHeaders(rowsPo[poHeaderIdx]);
  matches = matches.concat(collectScoredMatches(rowsPo, poCol));

  const teamStats = tallyTeamStats(matches);
  const rows = buildLegacyRows(teamStats, playerOfTeam);

  const { playerStage, bracketOrder } = parseTextRoundBracketStages(wsPo, playerOfTeam);
  rows.forEach((r) => {
    r.titles = pickBestTitle(playerStage[r.name], bracketOrder);
  });

  return { rows, skippedTeams, bracketNames: bracketOrder };
}

// Формат B: листы "Команды" (состав, игроки через "+"), "Расписание" (матчи группы,
// один круг), "Плей-офф" (сетки, по 1 матчу на раунд — структура столбцов как в формате A).
function parseLegacyFormatB(wb) {
  const wsTeams = wb.Sheets["Команды"];
  const wsSchedule = wb.Sheets["Расписание"];
  const wsPo = wb.Sheets["Плей-офф"];
  if (!wsTeams) throw new Error('Не найден лист "Команды"');
  if (!wsSchedule) throw new Error('Не найден лист "Расписание"');
  if (!wsPo) throw new Error('Не найден лист "Плей-офф"');

  const rowsTeams = sheetToRows(wsTeams);
  const teamsHeaderIdx = findHeaderRow(rowsTeams, ["Участники", "Команда"]);
  if (teamsHeaderIdx === -1) throw new Error('На листе "Команды" не найдена строка заголовков (Участники/Команда)');
  const teamsCol = indexHeaders(rowsTeams[teamsHeaderIdx]);

  const playerOfTeam = {};
  const skippedTeams = [];
  for (let r = teamsHeaderIdx + 1; r < rowsTeams.length; r++) {
    const row = rowsTeams[r];
    const participants = row[teamsCol["Участники"]];
    const team = row[teamsCol["Команда"]] ? String(row[teamsCol["Команда"]]).trim() : row[teamsCol["Команда"]];
    if (!team) continue;
    if (!participants || String(participants).trim() === "-" || String(participants).trim() === "") {
      skippedTeams.push(team);
      continue;
    }
    // "Илья+Женя" -> "Илья/Женя" — приводим к тому же разделителю, что и пары на сайте
    // (полные имена всё равно впишет админ вручную в превью, это лишь черновой вид).
    playerOfTeam[team] = normalizeName(String(participants).trim()).replace(/\s*\+\s*/g, "/");
  }

  const rowsSchedule = sheetToRows(wsSchedule);
  const scheduleHeaderIdx = findHeaderRow(rowsSchedule, ["Команда 1", "Гол 1", "Гол 2", "Команда 2"]);
  if (scheduleHeaderIdx === -1) throw new Error('На листе "Расписание" не найдена строка заголовков матчей');
  const scheduleCol = indexHeaders(rowsSchedule[scheduleHeaderIdx]);
  let matches = collectScoredMatches(rowsSchedule, scheduleCol);

  const rowsPo = sheetToRows(wsPo);
  const poHeaderIdx = findHeaderRow(rowsPo, ["Команда 1", "Гол 1", "Гол 2", "Команда 2"]);
  if (poHeaderIdx === -1) throw new Error('На листе "Плей-офф" не найдена строка заголовков матчей');
  const poCol = indexHeaders(rowsPo[poHeaderIdx]);
  matches = matches.concat(collectScoredMatches(rowsPo, poCol));

  const teamStats = tallyTeamStats(matches);
  const rows = buildLegacyRows(teamStats, playerOfTeam);

  const { playerStage, bracketOrder } = parseTextRoundBracketStages(wsPo, playerOfTeam);
  rows.forEach((r) => {
    r.titles = pickBestTitle(playerStage[r.name], bracketOrder);
  });

  return { rows, skippedTeams, bracketNames: bracketOrder };
}

// Формат C: листы "Расписание группы" + "Расписание плей-офф". Игроки фигурируют в матчах
// НАПРЯМУЮ (столбцы "Игрок 1"/"Игрок 2"), без привязки через клуб/команду — самый простой
// случай, привязка "игрок -> игрок" тождественная.
// Разбирает структуру плей-офф Формата C (там, где она есть): секции по имени
// ("Верхняя сетка"/"Средняя сетка"/"Нижняя сетка" и т.п.), но раунды здесь пронумерованы
// ("1 раунд", "2 раунд"...), а не названы словами — и одна и та же секция может
// встречаться повторно в разных местах листа (каскад: вылетел из верхней — попал в
// среднюю, и т.д.). Стадию определяем от МАКСИМАЛЬНОГО номера раунда, который вообще
// встретился для этой секции по всему листу: последний раунд = финал, предпоследний =
// полуфинал. Работает только там, где есть строка заголовков "Раунд"/"Игрок 1"/...; на
// листах без заголовков (запасной разбор по фиксированным колонкам) титулы не считаем —
// там просто нет опоры для определения раундов/секций.
function parseNumberedRoundBracketStages(sheetsWithRows) {
  const requiredHeaders = ["Раунд", "Игрок 1", "Гол 1", "Гол 2", "Игрок 2"];
  const matchesBySection = {};

  sheetsWithRows.forEach(({ rows }) => {
    const blocks = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i] || [];
      if (requiredHeaders.every((h) => row.includes(h))) {
        const col = indexHeaders(row);
        const roundCol = col["Раунд"];
        const sectionName = String((rows[i - 1] || [])[roundCol] || "").trim();
        if (sectionName) blocks.push({ sectionName, headerRowIdx: i, col, roundCol });
      }
    }
    blocks.forEach((block, bi) => {
      const endIdx = bi + 1 < blocks.length ? blocks[bi + 1].headerRowIdx - 1 : rows.length;
      for (let r = block.headerRowIdx + 1; r < endIdx; r++) {
        const row = rows[r] || [];
        const roundLabel = String(row[block.roundCol] || "").trim();
        const m = roundLabel.match(/^(\d+)\s*раунд/i);
        if (!m) continue;
        const roundNum = parseInt(m[1], 10);
        const t1 = row[block.col["Игрок 1"]];
        const t2 = row[block.col["Игрок 2"]];
        const g1 = row[block.col["Гол 1"]];
        const g2 = row[block.col["Гол 2"]];
        const winner = block.col["Победитель"] !== undefined ? row[block.col["Победитель"]] : null;
        if (typeof g1 === "number" && typeof g2 === "number" && t1 && t2) {
          (matchesBySection[block.sectionName] = matchesBySection[block.sectionName] || []).push({
            roundNum,
            t1: normalizeName(String(t1).trim()),
            t2: normalizeName(String(t2).trim()),
            winner: winner ? normalizeName(String(winner).trim()) : null,
          });
        }
      }
    });
  });

  const maxRoundBySection = {};
  Object.entries(matchesBySection).forEach(([section, ms]) => {
    maxRoundBySection[section] = Math.max(...ms.map((m) => m.roundNum));
  });

  const setStage = (obj, key, stage) => {
    const rank = STAGE_RANK[stage];
    if (!obj[key] || STAGE_RANK[obj[key]] < rank) obj[key] = stage;
  };

  const bracketOrder = [];
  const playerStage = {};
  Object.entries(matchesBySection).forEach(([section, ms]) => {
    if (!bracketOrder.includes(section)) bracketOrder.push(section);
    const maxRound = maxRoundBySection[section];
    ms.forEach(({ roundNum, t1, t2, winner }) => {
      let stage = null;
      if (roundNum === maxRound) stage = "final";
      else if (roundNum === maxRound - 1) stage = "semi";
      if (stage) {
        [t1, t2].forEach((name) => setStage((playerStage[name] = playerStage[name] || {}), section, stage));
        if (stage === "final" && winner) setStage((playerStage[winner] = playerStage[winner] || {}), section, "champion");
      }
    });
  });

  return { playerStage, bracketOrder };
}

// Разбирает структуру плей-офф там, где подписи раунда/сетки идут ОТДЕЛЬНОЙ СТРОКОЙ прямо
// над матчем (в том же столбце, что и "Игрок 1"), в свободном тексте — например,
// "Верхняя сетка. 1 раунд (1 м.+8 м.)" или "Верхняя сетка. ФИНАЛ". Смесь: ранние раунды
// пронумерованы, а финал часто подписан явным словом "ФИНАЛ"/"полуфинал" — учитываем оба
// варианта. Устойчиво к: опечаткам в названии сетки внутри одного файла ("Вехняя сетка"
// вместо "Верхняя сетка"), отсутствию строки заголовков вообще (тогда берём фиксированные
// колонки C/D/E/F/G с поправкой на возможную обрезанную пустую колонку A), и победителю,
// который из-за объединённых ячеек иногда попадает в строку-подпись раунда, а не в саму
// строку с результатом.
function parseInlineLabelBracketStages(sheetsWithRows, playerHeaders) {
  const bracketOrder = [];
  const playerStage = {};
  const setStage = (obj, key, stage) => {
    const rank = STAGE_RANK[stage];
    if (!obj[key] || STAGE_RANK[obj[key]] < rank) obj[key] = stage;
  };
  const cleanName = (v) => (v ? normalizeName(String(v).split("\n")[0].trim()) : null);

  const levenshtein = (a, b) => {
    const dp = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
    for (let j = 0; j <= b.length; j++) dp[0][j] = j;
    for (let i = 1; i <= a.length; i++) {
      for (let j = 1; j <= b.length; j++) {
        dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
    return dp[a.length][b.length];
  };
  // Список заведомо правильных названий сеток для этого стиля файлов — если встреченное
  // название похоже на одно из них (но написано не точь-в-точь, например "Вехняя сетка"
  // вместо "Верхняя сетка" — реальная опечатка в одном из исходных файлов), поправляем на
  // правильный вариант сразу здесь, а не гадаем потом при показе.
  const KNOWN_BRACKET_NAMES = ["Верхняя сетка", "Средняя сетка", "Нижняя сетка", "Кубок Интертото", "Финал Интертото"];
  const canonicalSection = (name) => {
    const exact = KNOWN_BRACKET_NAMES.find((k) => k.toLowerCase() === name.toLowerCase());
    if (exact) return exact;
    const known = KNOWN_BRACKET_NAMES.find((k) => levenshtein(k.toLowerCase(), name.toLowerCase()) <= 2);
    if (known) return known;
    const existing = bracketOrder.find((b) => levenshtein(b, name) <= 2);
    return existing || name;
  };

  const matchesBySection = {};

  const scanRows = (rows, t1Col, t2Col, g1Col, g2Col, winnerCol, startIdx) => {
    let currentSection = null;
    let labelRow = null;
    for (let i = startIdx; i < rows.length; i++) {
      const row = rows[i] || [];
      const cell = row[t1Col];
      const g1 = row[g1Col];
      const isMatchRow = cell && typeof g1 === "number";
      if (!isMatchRow) {
        if (cell && typeof cell === "string" && cell.includes(".")) {
          // "Фин. ниж. сетки+фин. кубка интер." — сокращённая запись бонусного матча между
          // чемпионами двух разных сеток. Модель "одна секция = одна сетка" тут не подходит
          // — пропускаем такие строки целиком, не пытаясь угадать.
          if (cell.includes("+")) { currentSection = null; labelRow = null; continue; }
          const dotIdx = cell.indexOf(".");
          const rawSection = cell.slice(0, dotIdx).trim();
          const sectionName = canonicalSection(rawSection);
          const rest = cell.slice(dotIdx + 1).trim();
          const restLower = rest.toLowerCase();
          let explicitStage = null;
          if (restLower.includes("полуфинал")) explicitStage = "semi";
          else if (restLower.includes("финал")) explicitStage = "final";
          const roundMatch = rest.match(/^(\d+)\s*раунд/i);
          const roundNum = roundMatch ? parseInt(roundMatch[1], 10) : null;
          currentSection = { name: sectionName, explicitStage, roundNum };
          labelRow = row; // на случай, если победитель окажется приклеен сюда (объединённая ячейка)
        }
        continue;
      }
      if (!currentSection) continue;
      const t1 = cleanName(row[t1Col]);
      const t2 = cleanName(row[t2Col]);
      const g2 = row[g2Col];
      if (t1 && t2 && typeof g1 === "number" && typeof g2 === "number") {
        if (!bracketOrder.includes(currentSection.name)) bracketOrder.push(currentSection.name);
        let winner = winnerCol !== undefined ? cleanName(row[winnerCol]) : null;
        if (!winner && labelRow && winnerCol !== undefined) winner = cleanName(labelRow[winnerCol]);
        (matchesBySection[currentSection.name] = matchesBySection[currentSection.name] || []).push({
          roundNum: currentSection.roundNum, explicitStage: currentSection.explicitStage, t1, t2, winner,
        });
      }
    }
  };

  sheetsWithRows.forEach(({ ws, rows }) => {
    const headerIdx = rows.findIndex((row) => row && row.includes(playerHeaders.t1) && row.includes(playerHeaders.t2));
    if (headerIdx !== -1) {
      const col = indexHeaders(rows[headerIdx]);
      const t1Col = col[playerHeaders.t1];
      const t2Col = col[playerHeaders.t2];
      const g1Col = col["Гол 1"] !== undefined ? col["Гол 1"] : t1Col + 1;
      const g2Col = col["Гол 2"] !== undefined ? col["Гол 2"] : t1Col + 2;
      const winnerCol = col["Победитель"];
      scanRows(rows, t1Col, t2Col, g1Col, g2Col, winnerCol, headerIdx + 1);
      return;
    }
    // Нет строки заголовков вообще — пробуем фиксированные колонки C/D/E/F/G (с поправкой
    // на возможную обрезанную пустую колонку A, как и в остальных запасных разборах).
    if (!ws || !ws["!ref"]) return;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const offset = range.s.c;
    const t1Col = 2 - offset, g1Col = 3 - offset, g2Col = 4 - offset, t2Col = 5 - offset, winnerCol = 6 - offset;
    scanRows(rows, t1Col, t2Col, g1Col, g2Col, winnerCol, 0);
  });

  Object.entries(matchesBySection).forEach(([section, ms]) => {
    const hasExplicitFinal = ms.some((m) => m.explicitStage === "final");
    const nums = ms.filter((m) => m.roundNum != null).map((m) => m.roundNum);
    const maxNumberedRound = nums.length ? Math.max(...nums) : null;
    ms.forEach((m) => {
      let stage = null;
      if (m.explicitStage === "final") stage = "final";
      else if (m.explicitStage === "semi") stage = "semi";
      else if (m.roundNum != null && maxNumberedRound != null) {
        if (hasExplicitFinal) {
          if (m.roundNum === maxNumberedRound) stage = "semi";
        } else {
          if (m.roundNum === maxNumberedRound) stage = "final";
          else if (m.roundNum === maxNumberedRound - 1) stage = "semi";
        }
      }
      if (!stage) return;
      setStage((playerStage[m.t1] = playerStage[m.t1] || {}), section, stage);
      setStage((playerStage[m.t2] = playerStage[m.t2] || {}), section, stage);
      if (stage === "final" && m.winner) setStage((playerStage[m.winner] = playerStage[m.winner] || {}), section, "champion");
    });
  });

  return { playerStage, bracketOrder };
}

function parseLegacyFormatC(wb) {
  const headers = { t1: "Игрок 1", g1: "Гол 1", g2: "Гол 2", t2: "Игрок 2" };
  const required = ["Игрок 1", "Гол 1", "Гол 2", "Игрок 2"];

  // Берём только листы, где по НАЗВАНИЮ ожидаем реальный список матчей ("Расписание..."/
  // "Распис..." — название сокращают по-разному в разных файлах, поэтому ищем по подстроке
  // "расп"). НЕ берём "Личные встречи", "Итог..." и подобные — на разных файлах это либо
  // просто повторяет те же самые матчи ещё раз в другом виде (было задвоение), либо просто
  // не совпадает по названию — подстрока "расп" сама по себе уже отсекает такие листы.
  // Исключение — если для плей-офф вообще нет отдельного "расп"-листа (как в паре файлов,
  // где итоги плей-офф лежат только на листе "Итог плей-офф") — тогда он единственный
  // источник, и его нужно подключить как запасной вариант.
  const raspSheetNames = wb.SheetNames.filter((name) => name.toLowerCase().includes("расп"));
  const raspHasPlayoff = raspSheetNames.some((name) => name.toLowerCase().includes("плей"));
  let sourceSheetNames = raspSheetNames;
  if (!raspHasPlayoff && wb.SheetNames.includes("Итог плей-офф")) {
    sourceSheetNames = sourceSheetNames.concat(["Итог плей-офф"]);
  }
  if (sourceSheetNames.length === 0) {
    throw new Error('Не найдено листов с расписанием матчей (название содержит "расп") или "Итог плей-офф"');
  }

  let matches = [];
  let sheetsUsed = 0;
  const sheetsWithRows = [];
  sourceSheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    const rows = sheetToRowsNoMergeDupes(ws);
    sheetsWithRows.push({ ws, rows });
    const headerIdx = findHeaderRow(rows, required);
    if (headerIdx !== -1) {
      const col = indexHeaders(rows[headerIdx]);
      matches = matches.concat(collectScoredMatches(rows, col, headers));
      sheetsUsed++;
      return;
    }
    // Запасной вариант: на некоторых листах плей-офф вообще нет строки заголовков —
    // данные лежат в фиксированных столбцах (C/D/E/F = Игрок1/Гол1/Гол2/Игрок2, блоками
    // "раунд / матч / команды"). Названия таких листов слишком разные, чтобы перечислять
    // все варианты ("плей-офф", "ЛЧ", "ЛЕ", "кубок"...) — поэтому не завязываемся на
    // название, а страхуемся фильтром по содержимому: настоящее имя игрока/пары не может
    // быть однобуквенным или голым числом (а вспомогательные листы жеребьёвки иногда дают
    // именно такие "матчи" по случайному совпадению чисел в нужных позициях).
    //
    // ВАЖНО: столбцы C/D/E/F считаем по НАСТОЯЩЕМУ адресу листа (через !ref), а не по
    // индексу в массиве после sheet_to_json — если колонка A на листе пустая, SheetJS
    // обрезает её и все номера столбцов "уезжают" на один влево (та же ловушка, что была
    // с листом "0_Настройки" в самом начале).
    const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { c: 0 } };
    const offset = range.s.c; // 0, если лист начинается с колонки A; >0, если она пустая и обрезана
    const ABS = { t1: 2, g1: 3, g2: 4, t2: 5 }; // столбцы C, D, E, F (0 = A)
    const relCol = { t1: ABS.t1 - offset, g1: ABS.g1 - offset, g2: ABS.g2 - offset, t2: ABS.t2 - offset };
    const fixed = collectScoredMatches(rows, relCol, { t1: "t1", g1: "g1", g2: "g2", t2: "t2" }).filter(
      ([t1, , , t2]) => t1.length > 1 && t2.length > 1 && !/^\d+$/.test(t1) && !/^\d+$/.test(t2)
    );
    if (fixed.length > 0) {
      matches = matches.concat(fixed);
      sheetsUsed++;
    }
  });
  if (sheetsUsed === 0) {
    throw new Error('На листах со списком матчей не найдена строка заголовков "Игрок 1"/"Гол 1"/"Гол 2"/"Игрок 2"');
  }

  const playerStats = tallyTeamStats(matches);
  const identity = {};
  Object.keys(playerStats).forEach((name) => { identity[name] = normalizeName(name); });
  const rows = buildLegacyRows(playerStats, identity);

  // Два разных стиля подписи раундов встречаются в разных файлах ("Раунд"-колонка с
  // пронумерованными раундами / отдельная строка-подпись с текстом прямо над матчем) —
  // пробуем оба, для конкретного файла реально сработает только один, это безопасно.
  const r1 = parseNumberedRoundBracketStages(sheetsWithRows);
  const r2 = parseInlineLabelBracketStages(sheetsWithRows, { t1: "Игрок 1", t2: "Игрок 2" });
  const bracketOrder = [...r1.bracketOrder, ...r2.bracketOrder.filter((b) => !r1.bracketOrder.includes(b))];
  const playerStage = {};
  [r1.playerStage, r2.playerStage].forEach((ps) => {
    Object.entries(ps).forEach(([name, stages]) => {
      const target = (playerStage[name] = playerStage[name] || {});
      Object.entries(stages).forEach(([bracket, stage]) => {
        const rank = STAGE_RANK[stage];
        if (!target[bracket] || STAGE_RANK[target[bracket]] < rank) target[bracket] = stage;
      });
    });
  });
  rows.forEach((r) => {
    r.titles = pickBestTitle(playerStage[r.name], bracketOrder);
  });

  return { rows, skippedTeams: [], bracketNames: bracketOrder };
}

// Формат D: листы, где есть столбцы "Пара 1"/"Пара 2" (сами заголовки голов называются
// по-разному в разных листах одного и того же файла — то единым "Счет" на 2 ячейки, то
// "Г1"/"Г2" — поэтому берём их не по названию, а по положению: сразу после "Пара 1" идут
// голы первой и второй пары, затем сама "Пара 2"). Имя пары иногда записано вместе с
// декоративным названием "команды" через перенос строки ("Оськин/Иванов\nСборная АПЛ") —
// берём только первую строчку.
// Формат D: в подписи матча стадия написана словами прямо в тексте ("Верхняя сетка. 2
// раунд (полуфинал)", "Верхняя сетка. ФИНАЛ"), плюс отдельный итоговый блок с прямым
// указанием чемпиона ("🏆 Чемпион Верхней сетки") — не нужно ничего вычислять, только
// прочитать текст.
function parseFormatDBracketStages(ws) {
  const rows = sheetToRowsNoMergeDupes(ws);
  const headerIdx = rows.findIndex((row) => row && row.includes("Пара 1") && row.includes("Пара 2"));
  if (headerIdx === -1) return { playerStage: {}, bracketOrder: [] };
  const pairCol = rows[headerIdx].indexOf("Пара 1");
  const g1Col = pairCol + 1;
  const g2Col = pairCol + 2;
  const pair2Col = pairCol + 3;
  const winnerCol = rows[headerIdx].indexOf("Победитель"); // -1, если такого столбца нет — тогда просто не определяем чемпиона

  const bracketOrder = [];
  const playerStage = {};
  const setStage = (obj, key, stage) => {
    const rank = STAGE_RANK[stage];
    if (!obj[key] || STAGE_RANK[obj[key]] < rank) obj[key] = stage;
  };
  const cleanName = (v) => (v ? normalizeName(String(v).split("\n")[0].trim()) : null);

  // "🏆 Чемпион <сетки>" намеренно НЕ разбираем: название сетки там стоит в родительном
  // падеже ("Чемпион Верхней сетки"), а в подписях матчей — в именительном ("Верхняя
  // сетка. ..."), из-за чего чемпион и финалист расходились по разным ключам и не
  // сравнивались между собой. Надёжнее взять победителя прямо из результата матча,
  // подписанного как "ФИНАЛ" — там название сетки гарантированно совпадает по форме.
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const label = row[0] ? String(row[0]).trim() : "";
    if (!label || label.startsWith("🏆")) continue;

    const dotIdx = label.indexOf(".");
    if (dotIdx === -1) continue; // строки вроде "1 тур" — не про конкретную сетку
    const sectionName = label.slice(0, dotIdx).trim();
    const rest = label.slice(dotIdx + 1).trim().toLowerCase();
    let stage = null;
    if (rest.includes("полуфинал")) stage = "semi";
    else if (rest.includes("финал")) stage = "final"; // "ФИНАЛ" без приставки "полу"
    if (!stage) continue;

    const dataRow = rows[i + 1] || [];
    const t1 = cleanName(dataRow[pairCol]);
    const t2 = cleanName(dataRow[pair2Col]);
    const g1 = dataRow[g1Col];
    const g2 = dataRow[g2Col];
    if (t1 && t2 && typeof g1 === "number" && typeof g2 === "number") {
      if (!bracketOrder.includes(sectionName)) bracketOrder.push(sectionName);
      [t1, t2].forEach((name) => setStage((playerStage[name] = playerStage[name] || {}), sectionName, stage));
      if (stage === "final" && winnerCol !== -1) {
        const winner = cleanName(dataRow[winnerCol]);
        if (winner) setStage((playerStage[winner] = playerStage[winner] || {}), sectionName, "champion");
      }
    }
  }

  return { playerStage, bracketOrder };
}

function parseLegacyFormatD(wb) {
  const sourceSheetNames = wb.SheetNames.filter((name) => {
    const rows = sheetToRowsNoMergeDupes(wb.Sheets[name]);
    return rows.some((row) => row && row.includes("Пара 1") && row.includes("Пара 2"));
  });
  if (sourceSheetNames.length === 0) {
    throw new Error('Не найдено листов со столбцами "Пара 1"/"Пара 2"');
  }

  let matches = [];
  let sheetsUsed = 0;
  sourceSheetNames.forEach((sheetName) => {
    const rows = sheetToRowsNoMergeDupes(wb.Sheets[sheetName]);
    const headerIdx = rows.findIndex((row) => row && row.includes("Пара 1") && row.includes("Пара 2"));
    if (headerIdx === -1) return;
    const pairCol = rows[headerIdx].indexOf("Пара 1");
    const g1Col = pairCol + 1, g2Col = pairCol + 2, pair2Col = pairCol + 3;
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      let t1 = row[pairCol];
      const g1 = row[g1Col];
      const g2 = row[g2Col];
      let t2 = row[pair2Col];
      if (t1 && t2 && typeof g1 === "number" && typeof g2 === "number") {
        t1 = String(t1).split("\n")[0].trim();
        t2 = String(t2).split("\n")[0].trim();
        matches.push([t1, g1, g2, t2]);
      }
    }
    sheetsUsed++;
  });
  if (sheetsUsed === 0) {
    throw new Error('На листах не найдена строка заголовков "Пара 1"/"Пара 2"');
  }

  const playerStats = tallyTeamStats(matches);
  const identity = {};
  Object.keys(playerStats).forEach((name) => { identity[name] = normalizeName(name); });
  const rows = buildLegacyRows(playerStats, identity);

  const poSheetName = sourceSheetNames.find((n) => n.toLowerCase().includes("плей"));
  if (poSheetName) {
    const { playerStage, bracketOrder } = parseFormatDBracketStages(wb.Sheets[poSheetName]);
    rows.forEach((r) => {
      r.titles = pickBestTitle(playerStage[r.name], bracketOrder);
    });
    return { rows, skippedTeams: [], bracketNames: bracketOrder };
  }
  return { rows, skippedTeams: [] };
}

// Современный формат ("0_Настройки"/"1_Участники"/"2_Расписание"/"3_Таблица"/
// "4_Плей-офф расписание"/"5_Сетка плей-офф") — тот же самый парсер, что используется в
// обычном "Импорте турнира" (parseTournamentFile), он уже считает и статистику, и %.
// Здесь берём из него только В/Н/П/голы для довнесения в уже существующую запись турнира —
// % оставляем прежним (не трогаем), как и для остальных исторических форматов, чтобы не
// создавать дублирующую запись турнира через обычный импорт.
function parseModernFormatForBackfill(wb) {
  const parsed = parseTournamentFile(wb);
  const rows = parsed.rows.map((r) => ({
    name: r.name,
    played: r.played,
    wins: r.wins,
    draws: r.draws,
    losses: r.losses,
    goalsFor: r.goalsFor,
    goalsAgainst: r.goalsAgainst,
    titles: r.titles,
  }));
  return { rows, skippedTeams: [], bracketNames: parsed.bracketNames };
}

// Компактный редактор титулов — плашки под именем игрока (не отдельные столбцы,
// чтобы не заставлять листать таблицу вправо). Если известен список сеток этого
// турнира (bracketOptions) — выбор сетки выпадающим списком; если нет (старые форматы,
// где структура сеток не разбирается) — свободный текст.
function TitlesEditor({ titles, bracketOptions, onChange }) {
  const [showAdd, setShowAdd] = useState(false);
  const [bracket, setBracket] = useState("");
  const [stage, setStage] = useState("Полуфиналист");
  const [newColor, setNewColor] = useState(null); // null = авто (по названию сетки)
  const [colorPickerIdx, setColorPickerIdx] = useState(null);
  const [manualBracket, setManualBracket] = useState(false); // вручную вписываем сетку вместо выбора из списка

  const removeTitle = (idx) => onChange(titles.filter((_, i) => i !== idx));
  const setTitleColor = (idx, color) => {
    onChange(titles.map((t, i) => (i === idx ? { ...t, color: color || undefined } : t)));
    setColorPickerIdx(null);
  };
  const addTitle = () => {
    if (!bracket.trim()) return;
    const entry = { bracket: bracket.trim(), stage };
    if (newColor) entry.color = newColor;
    onChange([...titles, entry]);
    setBracket("");
    setNewColor(null);
    setManualBracket(false);
    setShowAdd(false);
  };

  const ColorDots = ({ current, onPick }) => (
    <span className="inline-flex items-center gap-0.5">
      <button type="button" onClick={() => onPick(null)} title="Авто (по названию сетки)" className={`w-3 h-3 rounded-full bg-white border ${!current ? "border-slate-800 border-2" : "border-slate-300"}`} />
      {Object.values(TITLE_COLOR_PRESETS).map((p) => (
        <button key={p.key} type="button" onClick={() => onPick(p.key)} title={p.label} className={`w-3 h-3 rounded-full ${p.dot} ${current === p.key ? "ring-2 ring-offset-1 ring-slate-800" : ""}`} />
      ))}
    </span>
  );

  return (
    <div className="mt-1 flex flex-wrap items-center gap-1">
      {titles.map((t, i) => {
        const c = bracketColorClasses(t.bracket, t.color);
        return (
          <span key={i} className={`relative inline-flex items-center gap-1 ${c.bg} border ${c.border} ${c.text} text-[10px] px-1.5 py-0.5 rounded-full whitespace-nowrap`}>
            <button
              type="button"
              onClick={() => setColorPickerIdx(colorPickerIdx === i ? null : i)}
              title="Изменить цвет заливки"
              className={`w-2 h-2 rounded-full ${c.dot} shrink-0 ring-1 ring-black/10`}
            />
            {t.stage} ({t.bracket})
            <button onClick={() => removeTitle(i)} className={`${c.icon} hover:text-red-600 shrink-0`}>
              <X size={9} />
            </button>
            {colorPickerIdx === i && (
              <span className="absolute top-full left-0 mt-1 z-20 bg-white border border-slate-300 rounded-lg shadow-md px-1.5 py-1">
                <ColorDots current={t.color || null} onPick={(color) => setTitleColor(i, color)} />
              </span>
            )}
          </span>
        );
      })}
      {showAdd ? (
        <span className="inline-flex items-center gap-1 flex-wrap">
          {bracketOptions && bracketOptions.length > 0 && !manualBracket ? (
            <select
              value={bracket}
              onChange={(e) => {
                if (e.target.value === "__manual__") { setManualBracket(true); setBracket(""); }
                else setBracket(e.target.value);
              }}
              className="text-[10px] bg-white border border-slate-300 rounded px-1 py-0.5 max-w-[130px]"
            >
              <option value="">сетка…</option>
              {bracketOptions.map((b) => <option key={b} value={b}>{b}</option>)}
              <option value="__manual__">— другое (вписать) —</option>
            </select>
          ) : (
            <span className="inline-flex items-center gap-1">
              <input
                value={bracket}
                onChange={(e) => setBracket(e.target.value)}
                placeholder="название сетки"
                className="text-[10px] bg-white border border-slate-300 rounded px-1 py-0.5 w-28"
              />
              {bracketOptions && bracketOptions.length > 0 && (
                <button type="button" onClick={() => { setManualBracket(false); setBracket(""); }} className="text-[9px] text-blue-600 hover:underline shrink-0">список</button>
              )}
            </span>
          )}
          <select value={stage} onChange={(e) => setStage(e.target.value)} className="text-[10px] bg-white border border-slate-300 rounded px-1 py-0.5">
            <option value="Полуфиналист">Полуфиналист</option>
            <option value="Финалист">Финалист</option>
            <option value="Победитель">Победитель</option>
          </select>
          <ColorDots current={newColor} onPick={setNewColor} />
          <button onClick={addTitle} className="text-green-600 hover:text-green-700 shrink-0"><Check size={12} /></button>
          <button onClick={() => { setShowAdd(false); setManualBracket(false); }} className="text-slate-400 hover:text-slate-600 shrink-0"><X size={12} /></button>
        </span>
      ) : (
        <button onClick={() => setShowAdd(true)} className="text-[10px] text-blue-600 hover:underline shrink-0">+ титул</button>
      )}
    </div>
  );
}

const LEGACY_FORMATS = [
  { id: "formatA", label: "Формат A — команды/клубы, 2 матча в раунде (листы «Ввод», «Плей-офф») — например, ЧВ2", parser: parseLegacyFormatA },
  { id: "formatB", label: "Формат B — пары через «+», 1 матч в раунде (листы «Команды», «Расписание», «Плей-офф») — например, ЧВ3", parser: parseLegacyFormatB },
  { id: "formatC", label: "Формат C — игроки напрямую в матчах (листы «Расписание группы», «Расписание плей-офф») — большинство турниров ЧВ4–ЧВ17, ЧВ19–ЧВ26, SE1, SE2", parser: parseLegacyFormatC },
  { id: "formatD", label: "Формат D — столбцы «Пара 1»/«Пара 2» (листы «Расписание», «Плей-офф») — например, ЧВ18", parser: parseLegacyFormatD },
  { id: "modern", label: "Современный формат (0_Настройки/1_Участники/... — как в обычном импорте турнира) — начиная с ЧВ27", parser: parseModernFormatForBackfill },
];

// Общие для ВСЕХ парсеров (современного и исторических): стадии титулов и выбор
// единственного самого почётного титула за турнир — сначала по стадии (Победитель >
// Финалист > Полуфиналист), при равенстве — по порядку сетки в файле (раньше = почётнее).
const STAGE_LABEL = { semi: "Полуфиналист", fourth: "4 место", third: "3 место", final: "Финалист", champion: "Победитель" };

// Именованная палитра для плашек титулов — используется и для автоопределения по названию
// сетки, и как список вариантов при ручном выборе цвета в редакторе.
const TITLE_COLOR_PRESETS = {
  blue: { key: "blue", label: "Синий", bg: "bg-blue-100", border: "border-blue-300", text: "text-blue-800", icon: "text-blue-600", dot: "bg-blue-500" },
  purple: { key: "purple", label: "Фиолетовый", bg: "bg-purple-100", border: "border-purple-300", text: "text-purple-800", icon: "text-purple-600", dot: "bg-purple-500" },
  amber: { key: "amber", label: "Жёлтый", bg: "bg-amber-100", border: "border-amber-300", text: "text-amber-800", icon: "text-amber-600", dot: "bg-amber-500" },
  green: { key: "green", label: "Зелёный", bg: "bg-green-100", border: "border-green-300", text: "text-green-800", icon: "text-green-600", dot: "bg-green-500" },
  red: { key: "red", label: "Красный", bg: "bg-red-100", border: "border-red-300", text: "text-red-800", icon: "text-red-600", dot: "bg-red-500" },
  gray: { key: "gray", label: "Серый", bg: "bg-slate-200", border: "border-slate-300", text: "text-slate-800", icon: "text-slate-600", dot: "bg-slate-500" },
  pink: { key: "pink", label: "Розовый", bg: "bg-pink-100", border: "border-pink-300", text: "text-pink-800", icon: "text-pink-600", dot: "bg-pink-500" },
  teal: { key: "teal", label: "Бирюзовый", bg: "bg-teal-100", border: "border-teal-300", text: "text-teal-800", icon: "text-teal-600", dot: "bg-teal-500" },
};

// Автоопределение цвета по названию сетки — работает и для новых названий ("Лига чемпионов
// — Верхняя/Нижняя сетка", "Лига Европы", "Лига Конференций"), и для старых ("Верхняя/
// Средняя/Нижняя сетка", "Кубок/Финал Интертото"). Важный нюанс: "Лига чемпионов — Нижняя
// сетка" (ЛЧ НС) по названию говорит "нижняя", но по смыслу это СРЕДНЯЯ категория — красим
// как "Средняя сетка", а не как настоящую (старую) "Нижняя сетка". Проверка на ключевые
// слова — нечёткая (допускает 1-2 опечатки, вроде "Вехняя" вместо "Верхняя" — реальный
// случай в паре исходных файлов), чтобы мелкая ошибка в написании не путала цвет.
function autoBracketColorKey(bracketName) {
  const s = String(bracketName || "").toLowerCase();
  const hasChampionsPrefix = s.includes("чемпион");

  const levenshtein = (a, b) => {
    const dp = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
    for (let j = 0; j <= b.length; j++) dp[0][j] = j;
    for (let i = 1; i <= a.length; i++) {
      for (let j = 1; j <= b.length; j++) {
        dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
    return dp[a.length][b.length];
  };
  const hasWordLike = (keyword) =>
    s.includes(keyword) || s.split(/[\s—-]+/).some((w) => w.length > 3 && levenshtein(w, keyword) <= 2);

  if (hasWordLike("верхняя")) return "blue";
  if (hasWordLike("нижняя") && hasChampionsPrefix) return "purple";
  if (hasWordLike("средняя")) return "purple";
  if (hasWordLike("нижняя") || s.includes("европ")) return "amber";
  if (s.includes("интертот") || s.includes("конференц")) return "green";
  return "amber";
}

// overrideColor — необязательный ручной выбор (ключ из TITLE_COLOR_PRESETS), заданный
// админом при редактировании конкретного титула; если не задан — цвет определяется
// автоматически по названию сетки.
function bracketColorClasses(bracketName, overrideColor) {
  const key = overrideColor && TITLE_COLOR_PRESETS[overrideColor] ? overrideColor : autoBracketColorKey(bracketName);
  return TITLE_COLOR_PRESETS[key];
}
const STAGE_RANK = { semi: 1, fourth: 1, third: 2, final: 3, champion: 4 };

function pickBestTitle(bracketStageMap, bracketOrder) {
  let best = null;
  Object.entries(bracketStageMap || {}).forEach(([bracket, stage]) => {
    const rank = STAGE_RANK[stage];
    const orderIdx = bracketOrder.indexOf(bracket);
    if (!best || rank > best.rank || (rank === best.rank && orderIdx < best.orderIdx)) {
      best = { bracket, stage, rank, orderIdx };
    }
  });
  return best ? [{ bracket: best.bracket, stage: STAGE_LABEL[best.stage] }] : [];
}

function parseTournamentFile(wb) {
  const { G, vsSize } = readSettings(wb);
  if (!G) throw new Error('Не удалось найти "кол-во матчей" на листе "0_Настройки"');
  if (!vsSize) throw new Error('Не удалось найти размер верхней сетки на листе "0_Настройки"');

  const players = {};
  const bracketOrder = []; // порядок сеток как в файле — тай-брейк "почётности" при равной стадии
  computeGroupStats(wb, players);
  computePlayoffStats(wb, players);
  computeBracket(wb, players, bracketOrder);

  const R = Math.log2(vsSize);
  const idealO = 4 + G * 2 + R * 3 + 2 + 2 + 4;
  const idealMatches = G + R;
  const idealNorm = (idealO / idealMatches) * 10;

  const rows = Object.values(players).map((p) => {
    const played = G + p.poMatches;
    const O = computeO(p);
    const norm = played > 0 ? (O / played) * 10 : 0;
    const pct = idealNorm > 0 ? (norm / idealNorm) * 100 : 0;
    const wins = p.groupWins + p.playoffWins + p.intertotoWins;
    const draws = p.groupDraws;
    const losses = p.groupLosses + p.playoffLosses;
    const goalsFor = p.groupGoalsFor + p.playoffGoalsFor;
    const goalsAgainst = p.groupGoalsAgainst + p.playoffGoalsAgainst;
    const titles = pickBestTitle(p.bracketStage, bracketOrder);

    return {
      name: p.name,
      played,
      O: round3(O),
      norm: round3(norm),
      pct: round3(Math.max(pct, MIN_RATING)),
      wins, draws, losses, goalsFor, goalsAgainst,
      titles,
    };
  });
  rows.sort((a, b) => b.pct - a.pct);
  return {
    rows,
    meta: { G, vsSize, R, idealO: round3(idealO), idealNorm: round3(idealNorm) },
    bracketNames: bracketOrder,
  };
}

function round3(x) {
  return Math.round(x * 1000) / 1000;
}

// Срезает хвост вида " (100)" / " (-)" — это рейтинг игрока перед турниром, не часть имени —
// чтобы один и тот же игрок с разным рейтингом в разных турнирах считался одним человеком.
function normalizeName(raw) {
  if (!raw) return raw;
  return String(raw).replace(/\s*\([^)]*\)\s*$/, "").trim();
}

// Для поиска: "ё" и "е" считаются одной и той же буквой (Легенький ⇄ Лёгенький),
// плюс регистр не важен. Не путать с normalizeName выше — та чистит имя из файла,
// эта только приводит строку к виду для сравнения при поиске.
function normalizeForSearch(str) {
  return String(str).toLowerCase().replace(/ё/g, "е");
}

// Приводит имя пары к единому порядку ("Иванов/Петров" и "Петров/Иванов" — одна и та же пара),
// чтобы история и рейтинг не разваливались на два разных "игрока" из-за порядка перечисления.
// На соло-имена (без "/") не влияет.
function canonicalizePairName(name) {
  if (!name || !String(name).includes("/")) return name;
  return String(name)
    .split("/")
    .map((n) => n.trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "ru"))
    .join("/");
}

// ---------- КАТАЛОГ ИГРОКОВ (фамилия -> полное имя, для расшифровки пар) ----------

function surnameOf(fullName) {
  const parts = String(fullName).trim().split(/\s+/);
  return parts[parts.length - 1];
}

// Достаёт отдельные полные имена из имени турнира (соло: "Имя Фамилия", пара: "Имя1 Фамилия1/Имя2 Фамилия2")
function extractFullNames(name) {
  return String(name).split("/").map((n) => n.trim()).filter(Boolean);
}

function buildSurnameIndex(roster) {
  const index = {}; // фамилия (lowercase) -> [полные имена]
  roster.forEach((fullName) => {
    const sn = surnameOf(fullName).toLowerCase();
    if (!index[sn]) index[sn] = [];
    if (!index[sn].includes(fullName)) index[sn].push(fullName);
  });
  return index;
}

// Пытается расшифровать пару "Фамилия1/Фамилия2" в "Имя1 Фамилия1/Имя2 Фамилия2" по каталогу.
// Возвращает { resolvedName, parts: [{ raw, matches, resolved }] } — matches.length!==1 значит нужна ручная правка.
function resolvePairName(name, surnameIndex) {
  const rawParts = name.split("/").map((p) => p.trim());
  const parts = rawParts.map((raw) => {
    const matches = surnameIndex[raw.toLowerCase()] || [];
    return { raw, matches, resolved: matches.length === 1 ? matches[0] : null };
  });
  const allResolved = parts.every((p) => p.resolved);
  const resolvedName = allResolved ? parts.map((p) => p.resolved).join("/") : name;
  return { resolvedName, parts, allResolved };
}

// Парсинг старой вкладки "Итогового рейтинга" (Имя | Кол-во турниров | ЧВ1 | ЧВ2 | ... | Рейтинг за последние 3...)
// Каждый столбец с подзаголовком "Очки" — отдельный турнир, значение в ячейке — уже готовый % за турнир.
function parseHistorySheet(ws) {
  const data = sheetToRows(ws);
  let headerIdx = -1;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    if (row.includes("Имя") && row.includes("Кол-во турниров")) { headerIdx = i; break; }
  }
  if (headerIdx === -1) throw new Error('Не найдена строка заголовков ("Имя", "Кол-во турниров")');
  const headers = data[headerIdx];
  const subheader = data[headerIdx + 1] || [];
  const nameColIdx = headers.indexOf("Имя");

  const tournamentCols = [];
  subheader.forEach((v, idx) => { if (v === "Очки") tournamentCols.push(idx); });

  const tournaments = tournamentCols.map((colIdx) => {
    const rows = [];
    for (let r = headerIdx + 2; r < data.length; r++) {
      const row = data[r] || [];
      const name = canonicalizePairName(normalizeName(row[nameColIdx]));
      const val = row[colIdx];
      if (name && val !== null && val !== undefined && val !== "") {
        rows.push({
          name, played: null, O: null, norm: null,
          wins: null, draws: null, losses: null, goalsFor: null, goalsAgainst: null,
          pct: round3(Math.max(Number(val), MIN_RATING)),
        });
      }
    }
    return { name: String(headers[colIdx]), rows };
  });
  return tournaments;
}

function computeStandings(tournaments, mode, cutoffId) {
  let startIdx = 0;
  if (mode === "public" && cutoffId) {
    const idx = tournaments.findIndex((t) => t.id === cutoffId);
    if (idx >= 0) startIdx = idx;
  }
  const history = {};
  tournaments.forEach((t, idx) => {
    t.rows.forEach((p) => {
      if (!history[p.name]) history[p.name] = [];
      history[p.name].push({ pct: p.pct, idx, tournamentName: t.name });
    });
  });
  const results = [];
  Object.keys(history).forEach((name) => {
    const hist = history[name];
    const last3 = hist.slice(-3);
    const avg = last3.reduce((s, x) => s + x.pct, 0) / last3.length;
    const lastIdx = hist[hist.length - 1].idx;
    if (mode === "public" && lastIdx < startIdx) return;
    results.push({
      name,
      avg: round3(avg),
      tournamentsPlayed: hist.length,
      lastTournaments: last3.map((x) => x.tournamentName),
    });
  });
  results.sort((a, b) => b.avg - a.avg);
  return results;
}

// Тот же расчёт, что и обычный рейтинг (среднее по последним 3 турнирам КАЖДОГО игрока),
// но с отсечкой: остаются только те, кто участвовал хотя бы в одном из последних 3 турниров
// формата по хронологии (иначе игрок вычёркивается, даже если формально высокий средний %).
// Дальше — top-30.
function computeLast3EditionsStandings(tournaments, n = 3) {
  const startIdx = Math.max(0, tournaments.length - n);
  const history = {};
  tournaments.forEach((t, idx) => {
    t.rows.forEach((p) => {
      if (!history[p.name]) history[p.name] = [];
      history[p.name].push({ pct: p.pct, idx, tournamentName: t.name });
    });
  });
  const results = [];
  Object.keys(history).forEach((name) => {
    const hist = history[name];
    const last3 = hist.slice(-3);
    const avg = last3.reduce((s, x) => s + x.pct, 0) / last3.length;
    const lastIdx = hist[hist.length - 1].idx;
    if (lastIdx < startIdx) return; // не играл ни в одном из последних N турниров формата
    results.push({
      name,
      avg: round3(avg),
      tournamentsPlayed: hist.length,
      lastTournaments: last3.map((x) => x.tournamentName),
    });
  });
  results.sort((a, b) => b.avg - a.avg);
  return results.slice(0, 30);
}

function useStorage() {
  const [tournaments, setTournaments] = useState({ solo: [], pair: [], retro: [] });
  const [cutoffs, setCutoffs] = useState({ solo: null, pair: null, retro: null });
  const [roster, setRoster] = useState([]); // список известных полных имён игроков
  const [titlesEnabled, setTitlesEnabledState] = useState(true); // показывать ли титулы на сайте
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    let cancelled = false;
    async function load() {
      const nextT = { solo: [], pair: [], retro: [] };
      const nextC = { solo: null, pair: null, retro: null };
      let nextRoster = [];
      let nextTitlesEnabled = true;
      for (const f of FORMATS) {
        const tRes = await storageGet(`tournaments:${f.id}`);
        if (tRes && tRes.value) nextT[f.id] = tRes.value;
        const cRes = await storageGet(`cutoff:${f.id}`);
        if (cRes && cRes.value) nextC[f.id] = cRes.value.id;
      }
      const rRes = await storageGet("roster");
      if (rRes && rRes.value) nextRoster = rRes.value;
      const tsRes = await storageGet("settings:titlesEnabled");
      if (tsRes && tsRes.value) nextTitlesEnabled = tsRes.value.enabled;
      if (!cancelled) {
        setTournaments(nextT);
        setCutoffs(nextC);
        setRoster(nextRoster);
        setTitlesEnabledState(nextTitlesEnabled);
        setLoaded(true);
      }
    }
    load();
    return () => { cancelled = true; };
  }, []);

  const saveTournaments = useCallback(async (format, list) => {
    setTournaments((prev) => ({ ...prev, [format]: list }));
    await storageSet(`tournaments:${format}`, list);
  }, []);

  const saveCutoff = useCallback(async (format, id) => {
    setCutoffs((prev) => ({ ...prev, [format]: id }));
    await storageSet(`cutoff:${format}`, { id });
  }, []);

  const addToRoster = useCallback(async (names) => {
    const set = new Set(roster);
    let changed = false;
    names.forEach((n) => {
      const trimmed = n && n.trim();
      // В каталоге живут только отдельные полные имена — строка с "/" сюда попасть не должна
      // (это защита от повторения случая, когда парный турнир занесли как соло).
      if (trimmed && !trimmed.includes("/") && !set.has(trimmed)) { set.add(trimmed); changed = true; }
    });
    if (!changed) return roster;
    const merged = Array.from(set).sort();
    setRoster(merged);
    await storageSet("roster", merged);
    return merged;
  }, [roster]);

  const removeFromRoster = useCallback(async (namesToRemove) => {
    const toRemove = new Set(namesToRemove);
    const filtered = roster.filter((n) => !toRemove.has(n));
    setRoster(filtered);
    await storageSet("roster", filtered);
    return filtered;
  }, [roster]);

  const setTitlesEnabled = useCallback(async (enabled) => {
    setTitlesEnabledState(enabled);
    await storageSet("settings:titlesEnabled", { enabled });
  }, []);

  // Полное восстановление из бэкапа (экспорт/импорт JSON) — перезаписывает всё целиком,
  // а не сливает с текущим состоянием.
  const restoreAll = useCallback(async (backup) => {
    const nextT = backup.tournaments || { solo: [], pair: [], retro: [] };
    const nextC = backup.cutoffs || { solo: null, pair: null, retro: null };
    const nextR = backup.roster || [];
    for (const f of FORMATS) {
      await storageSet(`tournaments:${f.id}`, nextT[f.id] || []);
      await storageSet(`cutoff:${f.id}`, { id: nextC[f.id] || null });
    }
    await storageSet("roster", nextR);
    setTournaments(nextT);
    setCutoffs(nextC);
    setRoster(nextR);
  }, []);

  return { tournaments, cutoffs, roster, loaded, saveTournaments, saveCutoff, addToRoster, removeFromRoster, restoreAll, titlesEnabled, setTitlesEnabled };
}

function MedalBadge({ rank }) {
  const styles = {
    1: "bg-amber-400 text-amber-950",
    2: "bg-slate-300 text-slate-800",
    3: "bg-orange-400 text-orange-950",
  };
  if (rank > 3) return <span className="text-slate-500 tabular-nums w-7 inline-block text-right">{rank}</span>;
  return (
    <span className={`w-7 h-7 rounded-full inline-flex items-center justify-center text-sm font-bold ${styles[rank]}`}>
      {rank}
    </span>
  );
}

function RatingTable({ standings, query, onSelectPlayer }) {
  const filtered = query
    ? standings.filter((s) => normalizeForSearch(s.name).includes(normalizeForSearch(query)))
    : standings;

  if (standings.length === 0) {
    return (
      <div className="text-center py-16 text-slate-600">
        <Users className="mx-auto mb-3 opacity-40" size={32} />
        <p>Пока нет данных для этого рейтинга.</p>
      </div>
    );
  }
  if (filtered.length === 0) {
    return (
      <div className="text-center py-12 text-slate-500 text-sm">
        Никто не найден по запросу «{query}».
      </div>
    );
  }
  return (
    <div className="overflow-x-auto rounded-xl border border-slate-300/60">
      <table className="w-full text-sm">
        <thead>
          <tr className="bg-slate-200/80 text-slate-700 text-left uppercase tracking-normal sm:tracking-wide text-xs">
            <th className="py-3 pl-3 pr-1 sm:px-4 font-medium">#</th>
            <th className="py-3 px-1 sm:px-4 font-medium">Игрок</th>
            <th className="py-3 px-1 sm:px-4 font-medium text-right">
              <span className="sm:hidden">Турн.</span>
              <span className="hidden sm:inline">Турниров</span>
            </th>
            <th className="py-3 pl-1 pr-3 sm:px-4 font-medium text-right">Рейтинг</th>
          </tr>
        </thead>
        <tbody>
          {filtered.map((s) => {
            const rank = standings.indexOf(s) + 1;
            return (
              <tr
                key={s.name}
                onClick={() => onSelectPlayer(s.name)}
                className={`border-t border-slate-200 cursor-pointer hover:bg-slate-200/60 transition-colors ${rank % 2 === 0 ? "bg-slate-100/10" : "bg-slate-100/40"} ${query ? "ring-1 ring-red-500/40" : ""}`}
              >
                <td className="py-2.5 pl-3 pr-1 sm:px-4"><MedalBadge rank={rank} /></td>
                <td className="py-2.5 px-1 sm:px-4 text-slate-900 font-medium">{s.name}</td>
                <td className="py-2.5 px-1 sm:px-4 text-right text-slate-600 tabular-nums">{s.tournamentsPlayed}</td>
                <td className="py-2.5 pl-1 pr-3 sm:px-4 text-right font-mono font-semibold tabular-nums" style={{ color: BRAND_RED }}>{Math.round(s.avg)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// Сводит статистику игрока по всем турнирам формата: сумма W-Н-П/голов (там, где эти
// сырые данные есть) + список турниров с их индивидуальными показателями + все титулы.
function computePlayerDetail(tournaments, playerName) {
  const totals = { wins: 0, draws: 0, losses: 0, goalsFor: 0, goalsAgainst: 0, matches: 0, hasStats: false };
  const history = [];
  const titleAchievements = [];
  tournaments.forEach((t) => {
    const row = t.rows.find((r) => r.name === playerName);
    if (!row) return;
    const hasStats = row.wins !== null && row.wins !== undefined;
    if (hasStats) {
      totals.wins += row.wins;
      totals.draws += row.draws;
      totals.losses += row.losses;
      totals.goalsFor += row.goalsFor;
      totals.goalsAgainst += row.goalsAgainst;
      totals.matches += row.played;
      totals.hasStats = true;
    }
    history.push({ tournamentId: t.id, tournamentName: t.name, ...row, hasStats });
    (row.titles || []).forEach((title) => {
      titleAchievements.push({ tournamentName: t.name, bracket: title.bracket, stage: title.stage, color: title.color });
    });
  });
  history.reverse(); // сначала последние турниры
  titleAchievements.reverse();

  // Склеиваем одинаковые титулы в одну плашку ("3× Победитель (Верхняя сетка) — ЧВ24,
  // ЧВ14, ЧВ12") вместо отдельной плашки на каждый турнир. Группируем строго по точному
  // названию сетки — "Верхняя сетка" (старые турниры) и "Лига чемпионов — Верхняя сетка"
  // (новые) НЕ склеиваются, даже если цветом они совпадают, это разные титулы.
  const groupedTitles = [];
  const groupIndex = {};
  titleAchievements.forEach((t) => {
    const key = t.stage + "|||" + t.bracket;
    if (groupIndex[key] === undefined) {
      groupIndex[key] = groupedTitles.length;
      groupedTitles.push({ stage: t.stage, bracket: t.bracket, color: t.color, tournamentNames: [t.tournamentName] });
    } else {
      groupedTitles[groupIndex[key]].tournamentNames.push(t.tournamentName);
    }
  });

  return { totals, history, titleAchievements: groupedTitles };
}

function PlayerDetailModal({ playerName, tournaments, onClose, zIndex = 50, titlesEnabled = true }) {
  const detail = useMemo(() => computePlayerDetail(tournaments, playerName), [tournaments, playerName]);
  const { totals, history, titleAchievements } = detail;
  const [openTournamentId, setOpenTournamentId] = useState(null);
  const openTournament = openTournamentId ? tournaments.find((t) => t.id === openTournamentId) : null;

  // Сбрасываем прокрутку страницы наверх при открытии — если страницу пролистали
  // вниз (например, кликнули по игроку в конце длинного списка), fixed-элемент
  // внутри вложенного iframe иногда "уезжает" вместе со страницей.
  useEffect(() => {
    window.scrollTo(0, 0);
  }, []);

  return (
    <div className="fixed inset-0 bg-black/70 flex items-start justify-center p-2 sm:p-4 overflow-y-auto" style={{ zIndex }} onClick={onClose}>
      <div
        className="bg-slate-100 border border-slate-300 rounded-2xl w-full sm:max-w-lg flex flex-col shrink-0"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="shrink-0 bg-slate-100 border-b border-slate-200 px-5 py-4 flex items-center justify-between rounded-t-2xl">
          <h2 className="text-slate-900 font-semibold text-base">{playerName}</h2>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-700">
            <X size={18} />
          </button>
        </div>

        <div className="px-5 py-4 shrink-0">
          {titlesEnabled && titleAchievements.length > 0 && (
            <div className="mb-4">
              <h3 className="text-xs uppercase tracking-wide text-slate-500 mb-2">Титулы</h3>
              <div className="flex flex-wrap gap-1.5">
                {titleAchievements.map((t, i) => {
                  const c = bracketColorClasses(t.bracket, t.color);
                  const count = t.tournamentNames.length;
                  return (
                    <span
                      key={i}
                      className={`inline-flex items-center gap-1 ${c.bg} border ${c.border} ${c.text} text-xs px-2 py-1 rounded-full`}
                    >
                      <Trophy size={11} className={`${c.icon} shrink-0`} />
                      {count > 1 ? `${count}× ` : ""}{t.stage} ({t.bracket}) — {t.tournamentNames.join(", ")}
                    </span>
                  );
                })}
              </div>
            </div>
          )}

          <h3 className="text-xs uppercase tracking-wide text-slate-500 mb-2">Итого за все турниры</h3>
          {totals.hasStats ? (
            <div className="grid grid-cols-3 gap-2">
              <StatBox label="Матчей" value={totals.matches} />
              <StatBox label="Победы" value={totals.wins} accent="text-green-600" />
              <StatBox label="Ничьи" value={totals.draws} />
              <StatBox label="Поражения" value={totals.losses} accent="text-red-600" />
              <StatBox label="Голы забито" value={totals.goalsFor} />
              <StatBox label="Голы пропущено" value={totals.goalsAgainst} />
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Подробной статистики (В/Н/П, голы) пока нет — эти турниры были перенесены только по итоговому рейтингу.
            </p>
          )}
        </div>

        {/* Скроллится только сам список турниров — у него своя небольшая
            фиксированная высота, которая гарантированно помещается на экране
            целиком, независимо от панели браузера (внешняя модалка при этом
            не растягивается на весь экран и не упирается в неё). */}
        <div className="px-5 pb-5 flex flex-col min-h-0">
          <h3 className="text-xs uppercase tracking-wide text-slate-500 mb-2 shrink-0">
            По турнирам <span className="normal-case text-slate-400">(нажмите на турнир для подробностей)</span>
          </h3>
          <div className="space-y-2 overflow-y-auto max-h-64 pr-1">
            {history.map((h, i) => (
              <button
                key={i}
                onClick={() => setOpenTournamentId(h.tournamentId)}
                className="w-full text-left bg-slate-200/50 hover:bg-slate-200 transition-colors rounded-lg px-3 py-2.5"
              >
                <div className="flex items-center justify-between mb-1">
                  <span className="text-sm text-slate-800 font-medium">{h.tournamentName}</span>
                  <span className="font-mono text-sm font-semibold" style={{ color: BRAND_RED }}>{Math.round(h.pct)}</span>
                </div>
                {h.hasStats ? (
                  <p className="text-xs text-slate-500">
                    {h.played} матчей · {h.wins}В-{h.draws}Н-{h.losses}П · голы {h.goalsFor}:{h.goalsAgainst}
                  </p>
                ) : (
                  <p className="text-xs text-slate-400">только итоговый рейтинг</p>
                )}
              </button>
            ))}
          </div>
        </div>
      </div>

      {openTournament && (
        <TournamentDetailModal
          tournament={openTournament}
          tournaments={tournaments}
          onClose={() => setOpenTournamentId(null)}
          zIndex={zIndex + 10}
          titlesEnabled={titlesEnabled}
        />
      )}
    </div>
  );
}

// Уровень A: список участников турнира со статистикой + титулами, отсортирован по %.
function TournamentDetailModal({ tournament, tournaments, onClose, zIndex = 60, titlesEnabled = true }) {
  const sortedRows = useMemo(() => [...tournament.rows].sort((a, b) => b.pct - a.pct), [tournament]);
  const [selectedPlayer, setSelectedPlayer] = useState(null);

  return (
    <div className="fixed inset-0 bg-black/80 flex items-start justify-center p-2 sm:p-4 overflow-y-auto" style={{ zIndex }} onClick={onClose}>
      <div
        className="bg-slate-100 border border-slate-300 rounded-2xl w-full sm:max-w-xl shrink-0"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="bg-slate-100 border-b border-slate-200 px-5 py-4 flex items-center justify-between rounded-t-2xl">
          <h2 className="text-slate-900 font-semibold text-base truncate pr-3">{tournament.name}</h2>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-700 shrink-0">
            <X size={18} />
          </button>
        </div>
        <div className="p-4">
          <p className="text-xs text-slate-500 mb-2">
            Участников: {sortedRows.length} <span className="text-slate-400">(нажмите на игрока для личной статистики)</span>
          </p>
          <div className="max-h-[65vh] overflow-y-auto rounded-lg border border-slate-300">
            <table className="w-full text-xs">
              <thead className="bg-slate-200 text-slate-600 sticky top-0">
                <tr>
                  <th className="text-left py-2 px-2 w-6">#</th>
                  <th className="text-left py-2 px-2">Игрок</th>
                  <th className="text-right py-2 px-2">Игр</th>
                  <th className="text-right py-2 px-2">В-Н-П</th>
                  <th className="text-right py-2 px-2">Голы</th>
                  <th className="text-right py-2 px-2">%</th>
                </tr>
              </thead>
              <tbody>
                {sortedRows.map((r, i) => (
                  <tr
                    key={r.name}
                    onClick={() => setSelectedPlayer(r.name)}
                    className="border-t border-slate-200 cursor-pointer hover:bg-slate-200/60 transition-colors"
                  >
                    <td className="py-1.5 px-2 text-slate-400 tabular-nums">{i + 1}</td>
                    <td className="py-1.5 px-2 text-slate-900 font-medium">
                      {r.name}
                      {titlesEnabled && (r.titles || []).length > 0 && (
                        <div className="flex flex-wrap gap-1 mt-0.5">
                          {r.titles.map((t, ti) => {
                            const c = bracketColorClasses(t.bracket, t.color);
                            return (
                              <span key={ti} className={`inline-flex items-center gap-0.5 ${c.bg} border ${c.border} ${c.text} text-[10px] px-1.5 py-0.5 rounded-full`}>
                                <Trophy size={9} className={c.icon} />{t.stage}
                              </span>
                            );
                          })}
                        </div>
                      )}
                    </td>
                    <td className="py-1.5 px-2 text-right tabular-nums text-slate-600">{r.played ?? "—"}</td>
                    <td className="py-1.5 px-2 text-right tabular-nums text-slate-600 whitespace-nowrap">
                      {r.wins !== null && r.wins !== undefined ? `${r.wins}-${r.draws}-${r.losses}` : "—"}
                    </td>
                    <td className="py-1.5 px-2 text-right tabular-nums text-slate-600 whitespace-nowrap">
                      {r.goalsFor !== null && r.goalsFor !== undefined ? `${r.goalsFor}:${r.goalsAgainst}` : "—"}
                    </td>
                    <td className="py-1.5 px-2 text-right font-mono font-semibold tabular-nums" style={{ color: BRAND_RED }}>
                      {Math.round(r.pct)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {selectedPlayer && (
        <PlayerDetailModal
          playerName={selectedPlayer}
          tournaments={tournaments}
          onClose={() => setSelectedPlayer(null)}
          zIndex={zIndex + 10}
          titlesEnabled={titlesEnabled}
        />
      )}
    </div>
  );
}

function StatBox({ label, value, accent }) {
  return (
    <div className="bg-slate-200/60 rounded-lg px-3 py-2 text-center">
      <div className={`text-lg font-mono font-semibold tabular-nums ${accent || "text-slate-900"}`}>{value}</div>
      <div className="text-[10px] uppercase tracking-wide text-slate-500">{label}</div>
    </div>
  );
}

// Синхронизирует горизонтальный скролл двух контейнеров — используется, чтобы у широкой
// таблицы была полоса прокрутки не только внизу, но и сверху (видна сразу, без скролла вниз).
function useSyncedHorizontalScroll(deps) {
  const topRef = useRef(null);
  const bottomRef = useRef(null);
  const [scrollWidth, setScrollWidth] = useState(0);
  const syncing = useRef(false);

  useEffect(() => {
    const el = bottomRef.current;
    if (!el) return;
    const update = () => setScrollWidth(el.scrollWidth);
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  const onTopScroll = () => {
    if (syncing.current) { syncing.current = false; return; }
    syncing.current = true;
    if (bottomRef.current && topRef.current) bottomRef.current.scrollLeft = topRef.current.scrollLeft;
  };
  const onBottomScroll = () => {
    if (syncing.current) { syncing.current = false; return; }
    syncing.current = true;
    if (topRef.current && bottomRef.current) topRef.current.scrollLeft = bottomRef.current.scrollLeft;
  };

  return { topRef, bottomRef, scrollWidth, onTopScroll, onBottomScroll };
}

// Разбивает "ЧВ37 (26.08.2026)" на строку 1 "ЧВ37" и строку 2 "26.08.2026" —
// для компактного двухстрочного заголовка колонки в шахматке.
function splitTournamentName(name) {
  const m = String(name).match(/^(.*?)\s*\(([^)]*)\)\s*$/);
  if (m) return { main: m[1], sub: m[2] };
  return { main: name, sub: null };
}

// Для пары "Имя1 Фамилия1/Имя2 Фамилия2" — делит на 2 строки, чтобы колонка с именем
// в шахматке не занимала ширину суммы обоих имён, а только самого длинного из двух.
function PlayerCellName({ name }) {
  if (!name.includes("/")) return <span className="break-words">{name}</span>;
  const [a, b] = name.split("/");
  return (
    <div className="leading-tight py-0.5">
      <div className="break-words">{a}</div>
      <div className="break-words">{b}</div>
    </div>
  );
}

function MatrixView({ tournaments, cutoffs }) {
  const [format, setFormat] = useState("solo");
  const [mode, setMode] = useState("public");
  const list = tournaments[format] || [];
  const cutoffId = cutoffs[format];

  // Строки/сортировка — те же игроки и тот же порядок, что и в соответствующем режиме
  // основного рейтинга ("Актуальный" — с учётом отсечки, "За всё время" — без неё).
  const standings = useMemo(() => computeStandings(list, mode, cutoffId), [list, mode, cutoffId]);

  const cellData = useMemo(() => {
    const map = {};
    list.forEach((t, idx) => {
      t.rows.forEach((p) => {
        if (!map[p.name]) map[p.name] = {};
        map[p.name][idx] = p.pct;
      });
    });
    return map;
  }, [list]);

  const { topRef, bottomRef, scrollWidth, onTopScroll, onBottomScroll } = useSyncedHorizontalScroll([list, standings]);

  return (
    <div>
      <div className="flex flex-wrap gap-2 mb-2">
        <button
          onClick={() => setMode("public")}
          disabled={!cutoffId}
          style={mode === "public" ? { backgroundColor: BRAND_BLUE } : undefined}
          className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors disabled:opacity-30 disabled:cursor-not-allowed ${
            mode === "public" ? "text-white" : "bg-slate-200 text-slate-700 hover:bg-slate-300"
          }`}
        >
          Актуальный рейтинг
        </button>
        <button
          onClick={() => setMode("all")}
          style={mode === "all" ? { backgroundColor: BRAND_BLUE } : undefined}
          className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
            mode === "all" ? "text-white" : "bg-slate-200 text-slate-700 hover:bg-slate-300"
          }`}
        >
          Рейтинг за всё время
        </button>
      </div>

      <div className="flex flex-wrap gap-2 mb-6">
        {FORMATS.map((f) => (
          <button
            key={f.id}
            onClick={() => setFormat(f.id)}
            style={format === f.id ? { backgroundColor: BRAND_BLUE } : undefined}
            className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
              format === f.id
                ? "text-white"
                : "bg-slate-200 text-slate-700 hover:bg-slate-300"
            }`}
          >
            {f.label}
          </button>
        ))}
      </div>

      <p className="text-xs text-slate-500 mb-4">
        Рейтинг с деталировкой по каждому турниру
      </p>

      {list.length === 0 || standings.length === 0 ? (
        <div className="text-center py-16 text-slate-600">
          <Users className="mx-auto mb-3 opacity-40" size={32} />
          <p>Пока нет данных для этого формата.</p>
        </div>
      ) : (
        <>
          <div ref={topRef} onScroll={onTopScroll} className="overflow-x-auto overflow-y-hidden mb-1" style={{ height: 14 }}>
            <div style={{ width: scrollWidth, height: 1 }} />
          </div>
          <div ref={bottomRef} onScroll={onBottomScroll} className="overflow-x-auto rounded-xl border border-slate-300/60">
            <table className="text-sm border-collapse w-full">
              <thead>
                <tr className="bg-slate-200/80 text-slate-700 text-xs">
                  <th className="sticky left-0 bg-slate-200 z-10 px-3 py-2 text-left font-medium whitespace-nowrap">
                    <span className="inline-block w-5 text-slate-500">#</span>Игрок
                  </th>
                  <th className="px-3 py-2 text-right font-medium whitespace-nowrap border-l border-slate-300">Турниров</th>
                  {list.map((t) => {
                    const { main, sub } = splitTournamentName(t.name);
                    return (
                      <th key={t.id} className="px-2 py-2 text-right font-medium border-l border-slate-300">
                        <div className="whitespace-nowrap leading-tight">{main}</div>
                        {sub && <div className="whitespace-nowrap leading-tight font-normal text-slate-500 text-[10px]">{sub}</div>}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {standings.map((s, i) => (
                  <tr key={s.name} className={`border-t border-slate-200 ${i % 2 === 0 ? "bg-slate-100/10" : "bg-slate-100/40"}`}>
                    <td className="sticky left-0 bg-white px-3 py-1.5 text-slate-900 font-medium max-w-[165px]">
                      <div className="flex items-start">
                        <span className="inline-block w-5 shrink-0 text-slate-400 tabular-nums">{i + 1}</span>
                        <PlayerCellName name={s.name} />
                      </div>
                    </td>
                    <td className="px-3 py-1.5 text-right text-slate-600 tabular-nums border-l border-slate-200">{s.tournamentsPlayed}</td>
                    {list.map((t, idx) => {
                      const v = cellData[s.name]?.[idx];
                      return (
                        <td key={t.id} className="px-2 py-1.5 text-right tabular-nums text-slate-700 border-l border-slate-200">
                          {v !== undefined ? Math.round(v) : ""}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

// Список для "Рейтинг для Чемпионшипа": соло-игроки с рейтингом ≤50 + отдельно те, кого
// вообще нет в соло-рейтинге (никогда не играли соло), но у них есть пара с рейтингом ≤50 —
// такие тоже считаются "слабыми" и добавляются в список. Ретро не учитывается вовсе.
function ChampionshipList({ tournaments, cutoffs, query, onSelectPlayer }) {
  // Чемпионшип считается по рейтингу "за всё время" (без учёта отсечки "актуального"
  // рейтинга) — и для соло, и для пары.
  const soloStandings = useMemo(
    () => computeStandings(tournaments.solo || [], "all", null),
    [tournaments.solo]
  );
  const pairStandings = useMemo(
    () => computeStandings(tournaments.pair || [], "all", null),
    [tournaments.pair]
  );

  const rows = useMemo(() => {
    const soloNames = new Set(soloStandings.map((s) => s.name));
    // В имени пары человек иногда записан только фамилией ("Отрошко/Кутепов"), а в соло —
    // полным именем ("Сергей Кутепов") — точное сравнение строк это пропускало. Сверяемся
    // ещё и по фамилии (тот же справочник, что и для расшифровки пар при импорте).
    const soloSurnameIndex = buildSurnameIndex(soloStandings.map((s) => s.name));
    const hasSoloRecord = (person) => {
      if (soloNames.has(person)) return true;
      const matches = soloSurnameIndex[surnameOf(person).toLowerCase()];
      return !!matches && matches.length === 1;
    };

    const result = [];
    soloStandings.forEach((s) => {
      if (s.avg <= 50) result.push({ name: s.name, rating: s.avg, source: "solo", pairName: null });
    });

    const seenFallback = new Set();
    pairStandings.forEach((s) => {
      if (s.avg > 50) return;
      s.name.split("/").map((p) => p.trim()).forEach((person) => {
        if (hasSoloRecord(person)) return; // если игрок есть в соло (точно или по фамилии) — судим только по соло
        if (seenFallback.has(person)) return;
        seenFallback.add(person);
        result.push({ name: person, rating: s.avg, source: "pair", pairName: s.name });
      });
    });

    result.sort((a, b) => b.rating - a.rating);
    return result;
  }, [soloStandings, pairStandings]);

  const filtered = query
    ? rows.filter((r) => normalizeForSearch(r.name).includes(normalizeForSearch(query)))
    : rows;

  return (
    <div>
      <p className="text-xs text-slate-500 mb-4">
        Игроки с рейтингом 50 и ниже (за всё время, без учёта отсечки актуального
        рейтинга): показан соло-рейтинг — а если игрок никогда не играл соло, то рейтинг
        той пары, за которую он выступает (ретро не учитывается).
      </p>
      <div className="overflow-hidden rounded-xl border border-slate-300">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-200/80 text-slate-700 text-left uppercase tracking-normal sm:tracking-wide text-xs">
              <th className="py-3 pl-3 pr-1 w-8 font-medium">#</th>
              <th className="py-3 px-1 sm:px-4 font-medium">Игрок</th>
              <th className="py-3 pl-1 pr-3 sm:px-4 font-medium text-right">Рейтинг</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((r, i) => (
              <tr
                key={r.name}
                onClick={() => onSelectPlayer({ name: r.source === "pair" ? r.pairName : r.name, format: r.source === "pair" ? "pair" : "solo" })}
                className={`border-t border-slate-200 cursor-pointer hover:bg-slate-200/60 transition-colors ${i % 2 === 0 ? "bg-slate-100/10" : "bg-slate-100/40"}`}
              >
                <td className="py-2.5 pl-3 pr-1 text-slate-400 tabular-nums">{i + 1}</td>
                <td className="py-2.5 px-1 sm:px-4 text-slate-900 font-medium">{r.name}</td>
                <td className="py-2.5 pl-1 pr-3 sm:px-4 text-right font-mono font-semibold tabular-nums" style={{ color: BRAND_RED }}>
                  {Math.round(r.rating)}
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={3} className="py-6 text-center text-slate-400 text-sm">Никого не найдено</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PublicView({ tournaments, cutoffs, isAdmin, titlesEnabled }) {
  const [format, setFormat] = useState("solo");
  const [mode, setMode] = useState("public");
  const [query, setQuery] = useState("");
  const [selectedPlayer, setSelectedPlayer] = useState(null); // string (обычный режим) | { name, format } (Чемпионшип)
  const [showChampionship, setShowChampionship] = useState(false);

  const list = tournaments[format] || [];
  const standings = useMemo(() => {
    if (mode === "last3") return computeLast3EditionsStandings(list, 3);
    return computeStandings(list, mode, cutoffs[format]);
  }, [list, mode, cutoffs, format]);
  const cutoffTournament = list.find((t) => t.id === cutoffs[format]);

  // selectedPlayer бывает либо простой строкой (обычный режим — искать в текущем `list`),
  // либо { name, format } (режим "Чемпионшип" — там игрок может быть найден через пару,
  // и открывать его нужно в списке турниров того формата, где его реально нашли).
  const modalPlayerName = typeof selectedPlayer === "string" ? selectedPlayer : selectedPlayer?.name;
  const modalTournaments = typeof selectedPlayer === "string" ? list : tournaments[selectedPlayer?.format] || [];

  return (
    <div>
      {/* Верхний ряд: вид рейтинга — виден всем посетителям */}
      <div className="flex flex-wrap gap-2 mb-2">
        <button
          onClick={() => { setMode("public"); setShowChampionship(false); }}
          disabled={!cutoffs[format]}
          style={!showChampionship && mode === "public" ? { backgroundColor: BRAND_BLUE } : undefined}
          className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors disabled:opacity-30 disabled:cursor-not-allowed ${
            !showChampionship && mode === "public" ? "text-white" : "bg-slate-200 text-slate-700 hover:bg-slate-300"
          }`}
        >
          Актуальный рейтинг
        </button>
        <button
          onClick={() => { setMode("all"); setShowChampionship(false); }}
          style={!showChampionship && mode === "all" ? { backgroundColor: BRAND_BLUE } : undefined}
          className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
            !showChampionship && mode === "all" ? "text-white" : "bg-slate-200 text-slate-700 hover:bg-slate-300"
          }`}
        >
          Рейтинг за всё время
        </button>
        <button
          onClick={() => setShowChampionship(true)}
          style={showChampionship ? { backgroundColor: BRAND_BLUE } : undefined}
          className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
            showChampionship ? "text-white" : "bg-slate-200 text-slate-700 hover:bg-slate-300"
          }`}
        >
          Рейтинг для Чемпионшипа
        </button>
      </div>

      {/* Нижний ряд: формат — скрыт в режиме "Чемпионшип", там формат не выбирается отдельно */}
      {!showChampionship && (
        <div className="flex flex-wrap gap-2 mb-3">
          {FORMATS.map((f) => (
            <button
              key={f.id}
              onClick={() => setFormat(f.id)}
              style={format === f.id ? { backgroundColor: BRAND_BLUE } : undefined}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                format === f.id ? "text-white" : "bg-slate-200 text-slate-700 hover:bg-slate-300"
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>
      )}

      {!showChampionship && isAdmin && (
        <div className="mb-5">
          <button
            onClick={() => setMode("last3")}
            className={`px-3 py-1.5 rounded-md text-sm ${mode === "last3" ? "bg-slate-300 text-white" : "text-slate-600 hover:text-slate-800"}`}
          >
            Топ-30 за последние 3 турнира (только вы)
          </button>
        </div>
      )}

      <div className="relative mb-4">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" size={16} />
        <input
          type="text"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Найти себя по имени..."
          className="w-full bg-slate-100 border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-red-500"
        />
      </div>

      {showChampionship ? (
        <ChampionshipList
          tournaments={tournaments}
          cutoffs={cutoffs}
          query={query}
          onSelectPlayer={setSelectedPlayer}
        />
      ) : (
        <>
          <p className="text-xs text-slate-500 mb-1">
            {mode === "last3"
              ? `Топ-30 по среднему % за личные последние 3 турнира каждого — но только те, кто играл хотя бы в одном из последних 3 турниров формата «${FORMATS.find(f=>f.id===format)?.label}».`
              : "Очки рейтинга — это процент от максимально возможного количества набранных очков за один турнир. Рейтинг считается за 3 последних турнирах для каждого участника, чтобы отражать его актуальную форму."}
          </p>
          {mode === "public" && cutoffTournament && (
            <p className="text-xs text-slate-400 mb-4">
              В рейтинге учитываются только участники начиная с турнира «{cutoffTournament.name}».
            </p>
          )}
          {mode !== "public" && <div className="mb-4" />}

          <RatingTable standings={standings} query={query} onSelectPlayer={setSelectedPlayer} />
        </>
      )}

      {selectedPlayer && (
        <PlayerDetailModal
          playerName={modalPlayerName}
          tournaments={modalTournaments}
          onClose={() => setSelectedPlayer(null)}
          titlesEnabled={titlesEnabled}
        />
      )}
    </div>
  );
}

function HistoryImport({ tournaments, saveTournaments, addToRoster }) {
  const [sheetNames, setSheetNames] = useState([]);
  const [workbook, setWorkbook] = useState(null);
  const [mapping, setMapping] = useState({}); // sheetName -> format | ""
  const [replaceFlags, setReplaceFlags] = useState({ solo: true, pair: true, retro: true });
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [busy, setBusy] = useState(false);

  const handleFile = async (e) => {
    const f = e.target.files?.[0];
    setError(null);
    setSuccess(null);
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      const guess = {};
      wb.SheetNames.forEach((n) => {
        const low = n.toLowerCase();
        if (low.includes("соло") || low.includes("solo")) guess[n] = "solo";
        else if (low.includes("пар") || low.includes("pair")) guess[n] = "pair";
        else if (low.includes("ретро") || low.includes("retro")) guess[n] = "retro";
        else guess[n] = "";
      });
      setMapping(guess);
    } catch (err) {
      setError(err.message || String(err));
    }
  };

  const doImport = async () => {
    if (!workbook) return;
    setBusy(true);
    setError(null);
    setSuccess(null);
    try {
      const importedFormats = [];
      const allNames = [];
      for (const format of FORMATS.map((f) => f.id)) {
        const sheetName = Object.keys(mapping).find((n) => mapping[n] === format);
        if (!sheetName) continue;
        const ws = workbook.Sheets[sheetName];
        const parsedTournaments = parseHistorySheet(ws);
        const records = parsedTournaments.map((t) => ({
          id: `hist-${format}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          name: t.name,
          importedAt: new Date().toISOString(),
          rows: t.rows,
          meta: { imported: true, source: sheetName },
        }));
        records.forEach((r) => r.rows.forEach((row) => allNames.push(...extractFullNames(row.name))));
        const base = replaceFlags[format] ? [] : (tournaments[format] || []);
        await saveTournaments(format, [...base, ...records]);
        importedFormats.push(`${FORMATS.find((f) => f.id === format).label} (${records.length})`);
      }
      if (allNames.length > 0) await addToRoster(allNames);
      if (importedFormats.length === 0) {
        setError("Не выбрано ни одной вкладки для импорта — укажите формат для хотя бы одной.");
      } else {
        setSuccess(`Импортировано: ${importedFormats.join(", ")}. Каталог игроков пополнен (${new Set(allNames).size} имён).`);
      }
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="bg-slate-200/60 border border-slate-300 rounded-xl p-5">
      <h3 className="text-slate-800 font-semibold mb-1 flex items-center gap-2">
        <History size={18} /> Импорт истории рейтинга (мастер-файл)
      </h3>
      <p className="text-xs text-slate-500 mb-4">
        Загрузите старый файл "Рейтинг по сыгранным матчам.xlsx" со сводными вкладками — каждый столбец турнира
        станет отдельным турниром с уже готовым % (без пересчёта по сырым данным).
      </p>
      <input
        type="file"
        accept=".xlsx"
        onChange={handleFile}
        className="block w-full text-sm text-slate-700 mb-4 file:mr-3 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-slate-300 file:text-slate-800 file:text-sm hover:file:bg-slate-400"
      />

      {sheetNames.length > 0 && (
        <div className="space-y-2 mb-4">
          <p className="text-xs text-slate-600">Сопоставьте вкладки файла с форматами:</p>
          {sheetNames.map((n) => (
            <div key={n} className="flex items-center justify-between gap-3 bg-slate-100/50 rounded-md px-3 py-2">
              <span className="text-sm text-slate-700 truncate">{n}</span>
              <select
                value={mapping[n] || ""}
                onChange={(e) => setMapping((m) => ({ ...m, [n]: e.target.value }))}
                className="bg-slate-100 border border-slate-300 rounded-md text-xs text-slate-800 px-2 py-1 shrink-0"
              >
                <option value="">— не импортировать —</option>
                {FORMATS.map((f) => (
                  <option key={f.id} value={f.id}>{f.label}</option>
                ))}
              </select>
            </div>
          ))}

          <div className="flex flex-wrap gap-3 pt-2">
            {FORMATS.map((f) => (
              <label key={f.id} className="flex items-center gap-1.5 text-xs text-slate-600">
                <input
                  type="checkbox"
                  checked={replaceFlags[f.id]}
                  onChange={(e) => setReplaceFlags((r) => ({ ...r, [f.id]: e.target.checked }))}
                />
                Заменить текущие турниры «{f.label}» ({(tournaments[f.id] || []).length} сейчас)
              </label>
            ))}
          </div>

          <button
            onClick={doImport}
            disabled={busy}
            style={{ backgroundColor: BRAND_RED }}
            className="mt-2 hover:brightness-110 disabled:opacity-50 text-white font-medium text-sm px-4 py-2 rounded-md"
          >
            Импортировать выбранные вкладки
          </button>
        </div>
      )}

      {error && (
        <div className="flex items-start gap-2 bg-red-50 border border-red-300 text-red-700 text-sm rounded-md p-3">
          <AlertCircle size={16} className="mt-0.5 shrink-0" />
          <span>{error}</span>
        </div>
      )}
      {success && (
        <div className="flex items-start gap-2 bg-emerald-50 border border-emerald-300 text-emerald-700 text-sm rounded-md p-3">
          <CheckCircle2 size={16} className="mt-0.5 shrink-0" />
          <span>{success}</span>
        </div>
      )}
    </div>
  );
}

// Полное редактирование содержимого уже сохранённого турнира — имена, В/Н/П, голы,
// итоговый % за турнир — всё в одной таблице, открывается по клику на название турнира
// в списке. Можно и удалять/добавлять игроков вручную.
function TournamentEditModal({ tournament, onClose, onSave }) {
  const [rows, setRows] = useState(() => tournament.rows.map((r) => ({ ...r })));
  const [busy, setBusy] = useState(false);

  useEffect(() => { window.scrollTo(0, 0); }, []);

  const updateField = (idx, field, value) => {
    setRows((prev) => prev.map((r, i) => (i === idx ? { ...r, [field]: value } : r)));
  };
  const removeRow = (idx) => setRows((prev) => prev.filter((_, i) => i !== idx));
  const addRow = () => setRows((prev) => [
    ...prev,
    { name: "", played: 0, wins: 0, draws: 0, losses: 0, goalsFor: 0, goalsAgainst: 0, pct: 0, O: null, norm: null },
  ]);

  const save = async () => {
    setBusy(true);
    const cleaned = rows
      .filter((r) => r.name && String(r.name).trim())
      .map((r) => {
        const numOrNull = (v) => (v === null || v === "" || v === undefined ? null : Number(v));
        return {
          ...r,
          name: String(r.name).trim(),
          played: Number(r.played) || 0,
          wins: numOrNull(r.wins),
          draws: numOrNull(r.draws),
          losses: numOrNull(r.losses),
          goalsFor: numOrNull(r.goalsFor),
          goalsAgainst: numOrNull(r.goalsAgainst),
          pct: Number(r.pct) || 0,
        };
      });
    await onSave(cleaned);
    setBusy(false);
    onClose();
  };

  const numField = (idx, field, width = "w-12") => (
    <input
      type="number"
      value={rows[idx][field] === null || rows[idx][field] === undefined ? "" : rows[idx][field]}
      onChange={(e) => updateField(idx, field, e.target.value === "" ? null : Number(e.target.value))}
      className={`${width} bg-white border border-slate-300 rounded px-1 py-1 text-xs text-right`}
    />
  );

  return (
    <div className="fixed inset-0 bg-black/70 flex items-start justify-center z-50 p-2 sm:p-4 overflow-y-auto" onClick={onClose}>
      <div
        className="bg-slate-100 border border-slate-300 rounded-2xl w-full sm:max-w-3xl shrink-0"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="bg-slate-100 border-b border-slate-200 px-5 py-4 flex items-center justify-between rounded-t-2xl">
          <h2 className="text-slate-900 font-semibold text-base truncate pr-3">Редактирование: {tournament.name}</h2>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-700 shrink-0">
            <X size={18} />
          </button>
        </div>

        <div className="p-4">
          <div className="max-h-[60vh] overflow-y-auto rounded-lg border border-slate-300">
            <table className="w-full text-xs">
              <thead className="bg-slate-200 text-slate-600 sticky top-0">
                <tr>
                  <th className="w-6"></th>
                  <th className="text-left py-2 px-2">Имя</th>
                  <th className="text-right py-2 px-1">Игр</th>
                  <th className="text-right py-2 px-1">В</th>
                  <th className="text-right py-2 px-1">Н</th>
                  <th className="text-right py-2 px-1">П</th>
                  <th className="text-right py-2 px-1">ГЗ</th>
                  <th className="text-right py-2 px-1">ГП</th>
                  <th className="text-right py-2 px-2">%</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={i} className="border-t border-slate-200">
                    <td className="text-center">
                      <button onClick={() => removeRow(i)} className="text-slate-400 hover:text-red-500" title="Удалить игрока">
                        <X size={12} />
                      </button>
                    </td>
                    <td className="py-1 px-2">
                      <input
                        value={r.name}
                        onChange={(e) => updateField(i, "name", e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded px-1.5 py-1 text-xs"
                      />
                      <TitlesEditor
                        titles={r.titles || []}
                        bracketOptions={tournament.bracketNames}
                        onChange={(newTitles) => updateField(i, "titles", newTitles)}
                      />
                    </td>
                    <td className="py-1 px-1">{numField(i, "played")}</td>
                    <td className="py-1 px-1">{numField(i, "wins")}</td>
                    <td className="py-1 px-1">{numField(i, "draws")}</td>
                    <td className="py-1 px-1">{numField(i, "losses")}</td>
                    <td className="py-1 px-1">{numField(i, "goalsFor")}</td>
                    <td className="py-1 px-1">{numField(i, "goalsAgainst")}</td>
                    <td className="py-1 px-2">
                      <input
                        type="number"
                        step="0.1"
                        value={r.pct}
                        onChange={(e) => updateField(i, "pct", e.target.value === "" ? "" : Number(e.target.value))}
                        className="w-16 bg-white border border-slate-300 rounded px-1 py-1 text-xs text-right font-semibold"
                        style={{ color: BRAND_RED }}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="flex flex-wrap items-center gap-3 mt-3">
            <button onClick={addRow} className="bg-slate-300 hover:bg-slate-400 text-slate-800 text-xs px-3 py-1.5 rounded-md">
              + Добавить игрока
            </button>
            <button
              onClick={save}
              disabled={busy}
              style={{ backgroundColor: BRAND_RED }}
              className="hover:brightness-110 disabled:opacity-50 text-white text-sm font-medium px-4 py-2 rounded-md"
            >
              Сохранить изменения
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function TournamentRow({ index, tournament, onRename, onRemove, onEditContents }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState(tournament.name);

  const commit = () => {
    const trimmed = value.trim();
    if (trimmed && trimmed !== tournament.name) onRename(trimmed);
    else setValue(tournament.name);
    setEditing(false);
  };

  return (
    <li className="flex items-center justify-between gap-2 bg-slate-100/50 rounded-md px-3 py-1.5 text-sm">
      {editing ? (
        <div className="flex items-center gap-2 flex-1">
          <span className="text-slate-400 tabular-nums">{index + 1}.</span>
          <input
            autoFocus
            value={value}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") commit();
              if (e.key === "Escape") { setValue(tournament.name); setEditing(false); }
            }}
            className="flex-1 bg-slate-100 border border-slate-300 rounded px-2 py-1 text-sm text-slate-900"
          />
          <button onClick={commit} className="text-green-600 hover:text-green-700 shrink-0">
            <Check size={15} />
          </button>
        </div>
      ) : (
        <>
          <button onClick={onEditContents} className="text-slate-700 truncate text-left hover:underline" title="Открыть и отредактировать содержимое турнира">
            <span className="text-slate-400 tabular-nums mr-2">{index + 1}.</span>{tournament.name}
          </button>
          <div className="flex items-center gap-2 shrink-0">
            <button onClick={() => setEditing(true)} className="text-slate-400 hover:text-slate-700" title="Переименовать">
              <Pencil size={13} />
            </button>
            <button onClick={onRemove} className="text-slate-400 hover:text-red-500" title="Удалить">
              <Trash2 size={14} />
            </button>
          </div>
        </>
      )}
    </li>
  );
}

function AdminImport({ tournaments, saveTournaments, saveCutoff, cutoffs, roster, addToRoster, removeFromRoster, restoreAll, titlesEnabled, setTitlesEnabled }) {
  const [format, setFormat] = useState("solo");
  const [file, setFile] = useState(null);
  const [name, setName] = useState("");
  const [preview, setPreview] = useState(null);
  const [rows, setRows] = useState([]); // редактируемая копия preview.rows
  const [pairMap, setPairMap] = useState({}); // "Фамилия1/Фамилия2" -> { values:[v1,v2], candidates:[[...],[...]] }
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const [success, setSuccess] = useState(null);
  const [formatMismatch, setFormatMismatch] = useState(null); // "pair" | "solo" | null — что похоже на самом деле

  const surnameIndex = useMemo(() => buildSurnameIndex(roster), [roster]);

  const handleFile = async (e) => {
    const f = e.target.files?.[0];
    setError(null);
    setSuccess(null);
    setPreview(null);
    setRows([]);
    setPairMap({});
    setFormatMismatch(null);
    if (!f) return;
    setFile(f);
    if (!name) setName(f.name.replace(/\.xlsx?$/i, ""));
    setBusy(true);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const parsed = parseTournamentFile(wb);
      setPreview(parsed);
      setRows(parsed.rows.map((r) => ({ ...r })));

      // Проверка "это точно тот формат?": в файле пар имена содержат "/", в соло — нет.
      const pairLike = parsed.rows.filter((r) => r.name.includes("/")).length;
      const detected = pairLike > parsed.rows.length / 2 ? "pair" : "solo";
      if (format !== "retro" && detected !== format) setFormatMismatch(detected);

      if (format === "pair") {
        const uniqueNames = Array.from(new Set(parsed.rows.map((r) => r.name)));
        const map = {};
        uniqueNames.forEach((n) => {
          const { parts } = resolvePairName(n, surnameIndex);
          map[n] = {
            values: parts.map((p) => p.resolved || ""),
            candidates: parts.map((p) => p.matches),
            raw: parts.map((p) => p.raw),
          };
        });
        setPairMap(map);
      }
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusy(false);
    }
  };

  // Если пользователь меняет формат уже после того, как файл распознан — пересчитываем
  // и предупреждение, и сопоставление пар заново.
  useEffect(() => {
    if (!preview) return;
    const pairLike = preview.rows.filter((r) => r.name.includes("/")).length;
    const detected = pairLike > preview.rows.length / 2 ? "pair" : "solo";
    setFormatMismatch(format !== "retro" && detected !== format ? detected : null);

    if (format === "pair") {
      const uniqueNames = Array.from(new Set(preview.rows.map((r) => r.name)));
      const map = {};
      uniqueNames.forEach((n) => {
        const { parts } = resolvePairName(n, surnameIndex);
        map[n] = {
          values: parts.map((p) => p.resolved || ""),
          candidates: parts.map((p) => p.matches),
          raw: parts.map((p) => p.raw),
        };
      });
      setPairMap(map);
    } else {
      setPairMap({});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [format]);

  const updatePairValue = (pairName, partIdx, value) => {
    setPairMap((prev) => ({
      ...prev,
      [pairName]: {
        ...prev[pairName],
        values: prev[pairName].values.map((v, i) => (i === partIdx ? value : v)),
      },
    }));
  };

  const allPairsResolved = format !== "pair" || Object.values(pairMap).every((p) => p.values.every((v) => v && v.trim()));

  const confirmSave = async () => {
    if (!preview) return;
    if (format === "pair" && !allPairsResolved) {
      setError("Сначала укажите полные имена для всех пар — есть неразрешённые фамилии.");
      return;
    }
    setBusy(true);
    try {
      let finalRows = rows;
      if (format === "pair") {
        finalRows = rows.map((r) => {
          const resolved = pairMap[r.name];
          const newName = resolved ? canonicalizePairName(resolved.values.join("/")) : r.name;
          return { ...r, name: newName };
        });
        const allNewNames = Object.values(pairMap).flatMap((p) => p.values);
        await addToRoster(allNewNames);
      } else {
        await addToRoster(rows.map((r) => r.name));
      }

      const record = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        name: name || "Турнир",
        importedAt: new Date().toISOString(),
        rows: finalRows,
        meta: preview.meta,
        bracketNames: preview.bracketNames || [],
      };
      const updated = [...(tournaments[format] || []), record];
      await saveTournaments(format, updated);
      setSuccess(`Турнир «${record.name}» добавлен в формат «${FORMATS.find(f=>f.id===format)?.label}».`);
      setPreview(null);
      setRows([]);
      setPairMap({});
      setFile(null);
      setName("");
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusy(false);
    }
  };

  const removeTournament = async (fmt, id) => {
    const updated = (tournaments[fmt] || []).filter((t) => t.id !== id);
    await saveTournaments(fmt, updated);
    if (cutoffs[fmt] === id) await saveCutoff(fmt, null);
  };

  const renameTournament = async (fmt, id, newName) => {
    const updated = (tournaments[fmt] || []).map((t) => (t.id === id ? { ...t, name: newName } : t));
    await saveTournaments(fmt, updated);
  };

  const [editingTournament, setEditingTournament] = useState(null); // { format, id } | null
  const editingTournamentObj = editingTournament
    ? (tournaments[editingTournament.format] || []).find((t) => t.id === editingTournament.id)
    : null;

  const saveTournamentContents = async (newRows) => {
    const { format, id } = editingTournament;
    const updated = (tournaments[format] || []).map((t) => (t.id === id ? { ...t, rows: newRows } : t));
    await saveTournaments(format, updated);
  };

  const [fixBusy, setFixBusy] = useState(false);
  const [fixResult, setFixResult] = useState(null);

  const fixExistingRatings = async () => {
    setFixBusy(true);
    setFixResult(null);
    let totalFixed = 0;
    try {
      for (const f of FORMATS) {
        const list = tournaments[f.id] || [];
        let changed = false;
        const updated = list.map((t) => {
          const newRows = t.rows.map((r) => {
            if (r.pct < MIN_RATING) {
              changed = true;
              totalFixed++;
              return { ...r, pct: MIN_RATING };
            }
            return r;
          });
          return { ...t, rows: newRows };
        });
        if (changed) await saveTournaments(f.id, updated);
      }
      setFixResult(`Готово: поднято до ${MIN_RATING} — ${totalFixed} записей.`);
    } catch (err) {
      setFixResult(`Ошибка: ${err.message || err}`);
    } finally {
      setFixBusy(false);
    }
  };

  const [fixPairsBusy, setFixPairsBusy] = useState(false);
  const [fixPairsResult, setFixPairsResult] = useState(null);

  const fixPairOrdering = async () => {
    setFixPairsBusy(true);
    setFixPairsResult(null);
    let totalFixed = 0;
    try {
      const list = tournaments.pair || [];
      let changed = false;
      const updated = list.map((t) => {
        const newRows = t.rows.map((r) => {
          const canon = canonicalizePairName(r.name);
          if (canon !== r.name) { changed = true; totalFixed++; return { ...r, name: canon }; }
          return r;
        });
        return { ...t, rows: newRows };
      });
      if (changed) await saveTournaments("pair", updated);
      setFixPairsResult(`Готово: приведено к единому порядку — ${totalFixed} записей.`);
    } catch (err) {
      setFixPairsResult(`Ошибка: ${err.message || err}`);
    } finally {
      setFixPairsBusy(false);
    }
  };

  return (
    <div className="space-y-8">
      <div className="bg-slate-200/60 border border-slate-300 rounded-xl p-5 flex items-center justify-between gap-3">
        <div>
          <h3 className="text-slate-800 font-semibold">Показ титулов на сайте</h3>
          <p className="text-xs text-slate-500 mt-0.5">Если выключено — титулы нигде не отображаются (карточка игрока, карточка турнира), даже если они уже посчитаны и сохранены.</p>
        </div>
        <button
          onClick={() => setTitlesEnabled(!titlesEnabled)}
          className={`shrink-0 relative w-12 h-7 rounded-full transition-colors ${titlesEnabled ? "" : "bg-slate-300"}`}
          style={titlesEnabled ? { backgroundColor: BRAND_BLUE } : undefined}
        >
          <span className={`absolute top-1 w-5 h-5 bg-white rounded-full transition-transform ${titlesEnabled ? "translate-x-6" : "translate-x-1"}`} />
        </button>
      </div>

      <BackupPanel tournaments={tournaments} cutoffs={cutoffs} roster={roster} restoreAll={restoreAll} />

      <HistoryImport tournaments={tournaments} saveTournaments={saveTournaments} addToRoster={addToRoster} />

      <LegacyStatsImport tournaments={tournaments} saveTournaments={saveTournaments} addToRoster={addToRoster} roster={roster} />

      <div className="bg-slate-200/60 border border-slate-300 rounded-xl p-5">
        <h3 className="text-slate-800 font-semibold mb-4 flex items-center gap-2">
          <Upload size={18} /> Импорт турнира
        </h3>
        <div className="flex flex-wrap gap-2 mb-4">
          {FORMATS.map((f) => (
            <button
              key={f.id}
              onClick={() => setFormat(f.id)}
              style={format === f.id ? { backgroundColor: BRAND_BLUE } : undefined}
              className={`px-3 py-1.5 rounded-md text-sm ${
                format === f.id ? "text-white font-medium" : "bg-slate-300 text-slate-700 hover:bg-slate-400"
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>

        <label className="block text-xs text-slate-600 mb-1">Файл турнира (.xlsx)</label>
        <input
          type="file"
          accept=".xlsx"
          onChange={handleFile}
          className="block w-full text-sm text-slate-700 mb-3 file:mr-3 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-slate-300 file:text-slate-800 file:text-sm hover:file:bg-slate-400"
        />

        <label className="block text-xs text-slate-600 mb-1">Название турнира (для вкладки/списка)</label>
        <input
          type="text"
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Водокачка №38"
          className="w-full bg-slate-100 border border-slate-300 rounded-md px-3 py-2 text-sm text-slate-900 mb-4"
        />

        {error && (
          <div className="flex items-start gap-2 bg-red-50 border border-red-300 text-red-700 text-sm rounded-md p-3 mb-3">
            <AlertCircle size={16} className="mt-0.5 shrink-0" />
            <span>{error}</span>
          </div>
        )}
        {success && (
          <div className="flex items-start gap-2 bg-emerald-50 border border-emerald-300 text-emerald-700 text-sm rounded-md p-3 mb-3">
            <CheckCircle2 size={16} className="mt-0.5 shrink-0" />
            <span>{success}</span>
          </div>
        )}

        {preview && (
          <div className="mb-4">
            {formatMismatch && (
              <div className="flex items-start gap-2 bg-amber-50 border border-amber-300 text-amber-800 text-sm rounded-md p-3 mb-3">
                <AlertCircle size={16} className="mt-0.5 shrink-0" />
                <span>
                  Похоже, это на самом деле {formatMismatch === "pair" ? "парный" : "сольный"} турнир
                  (имена {formatMismatch === "pair" ? "в основном через «/»" : "в основном без «/»"}),
                  а выбран формат «{FORMATS.find((f) => f.id === format)?.label}». Проверьте формат выше, прежде чем сохранять.
                </span>
              </div>
            )}
            <p className="text-xs text-slate-600 mb-2">
              Проверка: G={preview.meta.G}, верхняя сетка={preview.meta.vsSize}, максимум %={100}, идеальные очки/матч={preview.meta.idealNorm}
            </p>

            {format === "pair" && Object.keys(pairMap).length > 0 && (
              <div className="mb-4 bg-slate-100/60 border border-slate-300 rounded-lg p-3">
                <p className="text-xs text-slate-600 mb-2">
                  Сопоставление пар с каталогом игроков — проверьте и заполните, где не нашлось однозначного совпадения:
                </p>
                <div className="space-y-2">
                  {Object.entries(pairMap).map(([pairName, info]) => {
                    const resolved = info.values.every((v) => v && v.trim());
                    return (
                      <div key={pairName} className={`flex flex-wrap items-center gap-2 text-xs rounded-md px-2 py-1.5 ${resolved ? "bg-slate-200/50" : "bg-amber-50 border border-amber-300"}`}>
                        <span className="text-slate-500 w-40 shrink-0">{pairName}</span>
                        {info.raw.map((rawSurname, idx) => (
                          <div key={idx} className="flex items-center gap-1">
                            <input
                              list={`roster-${pairName}-${idx}`}
                              value={info.values[idx]}
                              onChange={(e) => updatePairValue(pairName, idx, e.target.value)}
                              placeholder={`Полное имя (${rawSurname})`}
                              className={`bg-slate-100 border rounded px-2 py-1 text-xs w-44 ${info.values[idx] ? "border-slate-300 text-slate-800" : "border-amber-400 text-amber-800"}`}
                            />
                            <datalist id={`roster-${pairName}-${idx}`}>
                              {info.candidates[idx].map((c) => <option key={c} value={c} />)}
                            </datalist>
                          </div>
                        ))}
                        {!resolved && <AlertCircle size={14} className="text-amber-500" />}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            <div className="max-h-96 overflow-x-auto overflow-y-auto rounded-lg border border-slate-300">
              <table className="w-full text-xs">
                <thead className="bg-slate-200 text-slate-600 sticky top-0">
                  <tr>
                    <th className="text-left py-2 px-3">Игрок</th>
                    <th className="text-right py-2 px-3">Матчей</th>
                    <th className="text-right py-2 px-3">В</th>
                    <th className="text-right py-2 px-3">Н</th>
                    <th className="text-right py-2 px-3">П</th>
                    <th className="text-right py-2 px-3">Голы</th>
                    <th className="text-right py-2 px-3">O</th>
                    <th className="text-right py-2 px-3">% (рейтинг за турнир)</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={r.name} className="border-t border-slate-200 text-slate-700 align-top">
                      <td className="py-1.5 px-3">
                        <div>{r.name}</div>
                        <TitlesEditor
                          titles={r.titles || []}
                          bracketOptions={preview.bracketNames}
                          onChange={(newTitles) => setRows((prev) => prev.map((row, idx) => (idx === i ? { ...row, titles: newTitles } : row)))}
                        />
                      </td>
                      <td className="py-1.5 px-3 text-right tabular-nums">{r.played}</td>
                      <td className="py-1.5 px-3 text-right tabular-nums text-green-600">{r.wins}</td>
                      <td className="py-1.5 px-3 text-right tabular-nums text-slate-600">{r.draws}</td>
                      <td className="py-1.5 px-3 text-right tabular-nums text-red-600">{r.losses}</td>
                      <td className="py-1.5 px-3 text-right tabular-nums whitespace-nowrap">{r.goalsFor}:{r.goalsAgainst}</td>
                      <td className="py-1.5 px-3 text-right tabular-nums">{r.O}</td>
                      <td className="py-1 px-3 text-right">
                        <input
                          type="number"
                          step="0.1"
                          value={r.pct}
                          onChange={(e) => {
                            const val = e.target.value === "" ? "" : Number(e.target.value);
                            setRows((prev) => prev.map((row, idx) => (idx === i ? { ...row, pct: val } : row)));
                          }}
                          className="w-20 bg-slate-100 border border-slate-300 rounded px-1.5 py-0.5 text-right tabular-nums"
                          style={{ color: BRAND_RED }}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="text-xs text-slate-400 mt-1">
              Значения "%" можно поправить вручную перед сохранением. Титулы под именем сайт
              определил сам — можно убрать лишний или добавить через "+ титул", если ошибся.
            </p>
            {format === "pair" && !allPairsResolved && (
              <p className="text-amber-600 text-xs mt-2">Есть неразрешённые фамилии — заполните полные имена выше, прежде чем сохранять.</p>
            )}
            <button
              onClick={confirmSave}
              disabled={busy || (format === "pair" && !allPairsResolved)}
              style={{ backgroundColor: BRAND_RED }}
              className="mt-3 hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed text-white font-medium text-sm px-4 py-2 rounded-md"
            >
              Сохранить турнир
            </button>
          </div>
        )}
      </div>

      <div className="bg-slate-200/60 border border-slate-300 rounded-xl p-5">
        <h3 className="text-slate-800 font-semibold mb-4">Турниры по форматам</h3>
        <div className="flex items-center gap-3 mb-4 bg-slate-100/50 rounded-md px-3 py-2">
          <button
            onClick={fixExistingRatings}
            disabled={fixBusy}
            className="bg-slate-300 hover:bg-slate-400 disabled:opacity-50 text-slate-800 text-xs px-3 py-1.5 rounded-md shrink-0"
          >
            Поднять старые значения ниже {MIN_RATING} до {MIN_RATING}
          </button>
          {fixResult && <span className="text-xs text-slate-600">{fixResult}</span>}
        </div>
        <div className="flex items-center gap-3 mb-4 bg-slate-100/50 rounded-md px-3 py-2">
          <button
            onClick={fixPairOrdering}
            disabled={fixPairsBusy}
            className="bg-slate-300 hover:bg-slate-400 disabled:opacity-50 text-slate-800 text-xs px-3 py-1.5 rounded-md shrink-0"
          >
            Привести порядок имён в парах к единому виду
          </button>
          {fixPairsResult && <span className="text-xs text-slate-600">{fixPairsResult}</span>}
        </div>
        {FORMATS.map((f) => (
          <div key={f.id} className="mb-5 last:mb-0">
            <div className="flex items-center justify-between mb-2">
              <h4 className="text-sm text-slate-700 font-medium">{f.label}</h4>
              <div className="flex items-center gap-2">
                <label className="text-xs text-slate-500">Публичный рейтинг с:</label>
                <select
                  value={cutoffs[f.id] || ""}
                  onChange={(e) => saveCutoff(f.id, e.target.value || null)}
                  className="bg-slate-100 border border-slate-300 rounded-md text-xs text-slate-800 px-2 py-1"
                >
                  <option value="">— не задано —</option>
                  {(tournaments[f.id] || []).map((t) => (
                    <option key={t.id} value={t.id}>{t.name}</option>
                  ))}
                </select>
              </div>
            </div>
            {(tournaments[f.id] || []).length === 0 ? (
              <p className="text-xs text-slate-400">Ещё нет турниров.</p>
            ) : (
              <ol className="space-y-1">
                {tournaments[f.id].map((t, idx) => (
                  <TournamentRow
                    key={t.id}
                    index={idx}
                    tournament={t}
                    onRename={(newName) => renameTournament(f.id, t.id, newName)}
                    onRemove={() => removeTournament(f.id, t.id)}
                    onEditContents={() => setEditingTournament({ format: f.id, id: t.id })}
                  />
                ))}
              </ol>
            )}
          </div>
        ))}
      </div>

      <RosterPanel roster={roster} addToRoster={addToRoster} removeFromRoster={removeFromRoster} />

      {editingTournamentObj && (
        <TournamentEditModal
          tournament={editingTournamentObj}
          onClose={() => setEditingTournament(null)}
          onSave={saveTournamentContents}
        />
      )}
    </div>
  );
}

// Пытается сама подставить полное имя по каталогу игроков — используется при довнесении
// статистики старых турниров, где из файла достаются то голые фамилии, то сокращения.
// Для пары — расшифровывает обе половины и сразу приводит порядок к единому виду
// (canonicalizePairName), чтобы "Иванов/Петров" и "Петров/Иванов" не считались разными.
function resolveLegacyName(rawName, surnameIndex) {
  if (!rawName) return { resolved: rawName, autoResolved: false };
  if (String(rawName).includes("/")) {
    const parts = String(rawName).split("/").map((p) => p.trim());
    const resolvedParts = parts.map((p) => {
      const words = p.split(/\s+/);
      if (words.length >= 2) return p; // уже похоже на полное имя — оставляем как есть
      const matches = surnameIndex[p.toLowerCase()] || [];
      return matches.length === 1 ? matches[0] : null;
    });
    if (resolvedParts.every(Boolean)) {
      return { resolved: canonicalizePairName(resolvedParts.join("/")), autoResolved: true };
    }
    return { resolved: rawName, autoResolved: false };
  }
  const trimmed = String(rawName).trim();
  const tryKeys = [trimmed];
  const lastWord = trimmed.split(/\s+/).pop();
  if (lastWord && lastWord !== trimmed) tryKeys.push(lastWord);
  for (const key of tryKeys) {
    const matches = surnameIndex[key.toLowerCase()] || [];
    if (matches.length === 1) return { resolved: matches[0], autoResolved: true };
  }
  return { resolved: rawName, autoResolved: false };
}

function LegacyStatsImport({ tournaments, saveTournaments, addToRoster, roster }) {
  const [format, setFormat] = useState("solo");
  const [tournamentId, setTournamentId] = useState("");
  const [legacyFormatId, setLegacyFormatId] = useState(LEGACY_FORMATS[0].id);
  const [rows, setRows] = useState(null); // редактируемые строки превью
  const [skippedTeams, setSkippedTeams] = useState([]);
  const [bracketNames, setBracketNames] = useState(null); // список сеток, если парсер их знает (напр. "Современный формат")
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [busy, setBusy] = useState(false);

  const list = tournaments[format] || [];
  const targetTournament = list.find((t) => t.id === tournamentId);
  const existingNames = useMemo(() => new Set((targetTournament?.rows || []).map((r) => r.name)), [targetTournament]);

  const handleFile = async (e) => {
    const f = e.target.files?.[0];
    setError(null);
    setSuccess(null);
    setRows(null);
    setBracketNames(null);
    if (!f) return;
    if (!tournamentId) {
      setError("Сначала выберите, к какому уже сохранённому турниру относится этот файл.");
      e.target.value = "";
      return;
    }
    setBusy(true);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const legacyFormat = LEGACY_FORMATS.find((l) => l.id === legacyFormatId);
      const parsed = legacyFormat.parser(wb);

      // Пробуем сразу подставить полное имя по каталогу игроков — админ дальше просто
      // проверяет и подтверждает, а не вбивает всё заново.
      const surnameIndex = buildSurnameIndex(roster);
      const resolvedRows = parsed.rows.map((r) => {
        if (r.unresolved) return r; // "⚠️ Команда" без имени в файле — подставлять нечего
        const { resolved, autoResolved } = resolveLegacyName(r.name, surnameIndex);
        return autoResolved ? { ...r, name: resolved, autoResolved: true } : r;
      });

      setRows(resolvedRows);
      setSkippedTeams(parsed.skippedTeams);
      setBracketNames(parsed.bracketNames || null);
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusy(false);
    }
  };

  const updateRow = (idx, field, value) => {
    setRows((prev) => prev.map((r, i) => (i === idx ? { ...r, [field]: value } : r)));
  };

  const confirmSave = async () => {
    if (!rows || !targetTournament) return;
    setBusy(true);
    setError(null);
    try {
      const statsByName = {};
      rows.forEach((r) => {
        statsByName[r.name] = {
          played: Number(r.played) || 0,
          wins: Number(r.wins) || 0,
          draws: Number(r.draws) || 0,
          losses: Number(r.losses) || 0,
          goalsFor: Number(r.goalsFor) || 0,
          goalsAgainst: Number(r.goalsAgainst) || 0,
          titles: r.titles || [],
        };
      });

      const existingRows = targetTournament.rows.map((r) => {
        const s = statsByName[r.name];
        if (!s) return r;
        delete statsByName[r.name];
        return { ...r, ...s };
      });
      // Игроки из файла, которых не было среди уже сохранённых (новые записи) — добавляем как есть,
      // без %, так как рейтинга за этот турнир у них не было в исходном импорте истории.
      const newRows = Object.keys(statsByName).map((name) => ({
        name, pct: 0, O: null, norm: null, ...statsByName[name],
      }));
      const finalRows = [...existingRows, ...newRows];

      const updatedTournaments = list.map((t) => (t.id === tournamentId ? { ...t, rows: finalRows } : t));
      await saveTournaments(format, updatedTournaments);
      await addToRoster(rows.filter((r) => !r.unresolved).map((r) => r.name));

      setSuccess(`Статистика добавлена в «${targetTournament.name}»: обновлено ${finalRows.length - newRows.length}, добавлено новых ${newRows.length}.`);
      setRows(null);
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="bg-slate-100 border border-slate-300 rounded-xl p-5">
      <h3 className="text-slate-900 font-semibold mb-1">Довнесение статистики старых турниров</h3>
      <p className="text-xs text-slate-500 mb-4">
        Для турниров, которые уже есть на сайте только с итоговым %, — загрузите исходный файл турнира,
        чтобы добавить В/Н/П и голы. Сам % за турнир при этом не меняется.
      </p>

      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <div>
          <label className="block text-xs text-slate-500 mb-1">Формат</label>
          <div className="flex gap-2">
            {FORMATS.map((f) => (
              <button
                key={f.id}
                onClick={() => { setFormat(f.id); setTournamentId(""); setRows(null); }}
                style={format === f.id ? { backgroundColor: BRAND_BLUE } : undefined}
                className={`px-3 py-1.5 rounded-md text-sm ${format === f.id ? "text-white font-medium" : "bg-slate-300 text-slate-700 hover:bg-slate-400"}`}
              >
                {f.label}
              </button>
            ))}
          </div>
        </div>
        <div>
          <label className="block text-xs text-slate-500 mb-1">Какой турнир пополняем</label>
          <select
            value={tournamentId}
            onChange={(e) => { setTournamentId(e.target.value); setRows(null); }}
            className="w-full bg-white border border-slate-300 rounded-md text-sm text-slate-800 px-2 py-1.5"
          >
            <option value="">— выберите турнир —</option>
            {list.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
          </select>
        </div>
      </div>

      <div className="mb-3">
        <label className="block text-xs text-slate-500 mb-1">Тип файла (структура старой таблицы)</label>
        <select
          value={legacyFormatId}
          onChange={(e) => setLegacyFormatId(e.target.value)}
          className="w-full bg-white border border-slate-300 rounded-md text-sm text-slate-800 px-2 py-1.5"
        >
          {LEGACY_FORMATS.map((l) => <option key={l.id} value={l.id}>{l.label}</option>)}
        </select>
      </div>

      <input
        type="file"
        accept=".xlsx"
        onChange={handleFile}
        disabled={!tournamentId || busy}
        className="block w-full text-sm text-slate-700 mb-4 file:mr-3 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-slate-300 file:text-slate-800 file:text-sm hover:file:bg-slate-400 disabled:opacity-50"
      />

      {error && (
        <div className="flex items-start gap-2 bg-red-50 border border-red-300 text-red-700 text-sm rounded-md p-3 mb-3">
          <AlertCircle size={16} className="mt-0.5 shrink-0" />
          <span>{error}</span>
        </div>
      )}
      {success && (
        <div className="flex items-start gap-2 bg-emerald-50 border border-emerald-300 text-emerald-700 text-sm rounded-md p-3 mb-3">
          <CheckCircle2 size={16} className="mt-0.5 shrink-0" />
          <span>{success}</span>
        </div>
      )}

      {rows && (
        <div>
          {skippedTeams.length > 0 && (
            <p className="text-xs text-amber-700 bg-amber-50 border border-amber-300 rounded-md p-2 mb-3">
              Без указанного игрока в файле (пропущены): {skippedTeams.join(", ")}
            </p>
          )}
          <p className="text-xs text-slate-500 mb-2">
            Проверьте и поправьте имена (особенно там, где ⚠️ — игрок не был указан в файле явно).
            Зелёный кружок — имя совпало с уже сохранённым в этом турнире игроком; жёлтый — не совпало,
            при сохранении попадёт как новая запись. Значок <Sparkles size={11} className="inline text-blue-500" /> — имя подставлено
            автоматически по каталогу игроков (по фамилии), проверьте, что оно верное.
          </p>
          <div className="max-h-96 overflow-y-auto rounded-lg border border-slate-300">
            <table className="w-full text-xs">
              <thead className="bg-slate-200 text-slate-600 sticky top-0">
                <tr>
                  <th className="w-4"></th>
                  <th className="text-left py-2 px-2">Имя</th>
                  <th className="text-right py-2 px-1">Игр</th>
                  <th className="text-right py-2 px-1">В</th>
                  <th className="text-right py-2 px-1">Н</th>
                  <th className="text-right py-2 px-1">П</th>
                  <th className="text-right py-2 px-1">ГЗ</th>
                  <th className="text-right py-2 px-2">ГП</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const matched = existingNames.has(r.name);
                  return (
                    <tr key={i} className="border-t border-slate-200">
                      <td className="text-center">
                        <span className={`inline-block w-2 h-2 rounded-full ${matched ? "bg-green-500" : "bg-amber-400"}`} title={matched ? "совпало" : "новая запись"} />
                      </td>
                      <td className="py-1 px-2">
                        <div className="flex items-center gap-1">
                          <input
                            value={r.name}
                            onChange={(e) => updateRow(i, "name", e.target.value)}
                            className={`w-full bg-white border rounded px-1.5 py-1 text-xs ${matched ? "border-slate-300" : "border-amber-400"}`}
                          />
                          {r.autoResolved && (
                            <Sparkles size={12} className="text-blue-500 shrink-0" title="Имя подставлено автоматически по каталогу — проверьте" />
                          )}
                        </div>
                        <TitlesEditor
                          titles={r.titles || []}
                          bracketOptions={bracketNames}
                          onChange={(newTitles) => updateRow(i, "titles", newTitles)}
                        />
                      </td>
                      {["played", "wins", "draws", "losses", "goalsFor", "goalsAgainst"].map((field) => (
                        <td key={field} className="py-1 px-1">
                          <input
                            type="number"
                            value={r[field]}
                            onChange={(e) => updateRow(i, field, e.target.value === "" ? "" : Number(e.target.value))}
                            className="w-12 bg-white border border-slate-300 rounded px-1 py-1 text-xs text-right"
                          />
                        </td>
                      ))}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button
            onClick={confirmSave}
            disabled={busy}
            style={{ backgroundColor: BRAND_RED }}
            className="mt-3 hover:brightness-110 disabled:opacity-50 text-white font-medium text-sm px-4 py-2 rounded-md"
          >
            Сохранить в «{targetTournament?.name}»
          </button>
        </div>
      )}
    </div>
  );
}

function BackupPanel({ tournaments, cutoffs, roster, restoreAll }) {
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState(null);

  const doExport = () => {
    const backup = {
      exportedAt: new Date().toISOString(),
      tournaments,
      cutoffs,
      roster,
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `vodokachka-backup-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const doImport = async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setBusy(true);
    setMessage(null);
    try {
      const text = await f.text();
      const backup = JSON.parse(text);
      if (!backup.tournaments || !backup.roster) throw new Error("Файл не похож на бэкап этого сайта.");
      await restoreAll(backup);
      const counts = FORMATS.map((fmt) => `${fmt.label}: ${(backup.tournaments[fmt.id] || []).length}`).join(", ");
      setMessage(`Восстановлено. Турниры — ${counts}. Каталог игроков — ${backup.roster.length} имён.`);
    } catch (err) {
      setMessage(`Ошибка: ${err.message || err}`);
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };

  return (
    <div className="bg-slate-200/60 border border-slate-300 rounded-xl p-5">
      <h3 className="text-slate-800 font-semibold mb-1">Бэкап всех данных</h3>
      <p className="text-xs text-slate-500 mb-4">
        Скачайте JSON-файл со всеми турнирами, каталогом игроков и точками отсчёта — на случай сбоя
        или переноса на другую версию сайта. Импорт полностью заменяет текущие данные содержимым файла.
      </p>
      <div className="flex flex-wrap items-center gap-3">
        <button
          onClick={doExport}
          className="bg-slate-300 hover:bg-slate-400 text-slate-800 text-sm px-4 py-2 rounded-md"
        >
          Экспорт бэкапа
        </button>
        <label className="bg-slate-300 hover:bg-slate-400 text-slate-800 text-sm px-4 py-2 rounded-md cursor-pointer">
          Импортировать бэкап
          <input type="file" accept=".json" onChange={doImport} disabled={busy} className="hidden" />
        </label>
        {message && <span className="text-xs text-slate-600">{message}</span>}
      </div>
    </div>
  );
}

function RosterPanel({ roster, addToRoster, removeFromRoster }) {
  const [query, setQuery] = useState("");
  const [newName, setNewName] = useState("");
  const [cleanupBusy, setCleanupBusy] = useState(false);
  const [cleanupMsg, setCleanupMsg] = useState(null);
  const filtered = query ? roster.filter((n) => normalizeForSearch(n).includes(normalizeForSearch(query))) : roster;
  const garbageCount = roster.filter((n) => n.includes("/")).length;

  const cleanupGarbage = async () => {
    setCleanupBusy(true);
    const garbage = roster.filter((n) => n.includes("/"));
    await removeFromRoster(garbage);
    setCleanupMsg(`Удалено записей с "/": ${garbage.length}.`);
    setCleanupBusy(false);
  };

  return (
    <div className="bg-slate-200/60 border border-slate-300 rounded-xl p-5">
      <h3 className="text-slate-800 font-semibold mb-1">Каталог игроков</h3>
      <p className="text-xs text-slate-500 mb-4">
        Полные имена, по которым сайт расшифровывает фамилии в парных турнирах. Пополняется автоматически
        при импорте, можно добавить вручную.
      </p>

      {garbageCount > 0 && (
        <div className="flex items-center gap-3 mb-3 bg-amber-50 border border-amber-300 rounded-md px-3 py-2">
          <span className="text-xs text-amber-200 flex-1">
            В каталоге {garbageCount} "мусорных" записей с "/" (вероятно, парный турнир попал в соло) — их не должно быть.
          </span>
          <button
            onClick={cleanupGarbage}
            disabled={cleanupBusy}
            className="bg-amber-700 hover:bg-amber-600 disabled:opacity-50 text-amber-50 text-xs px-3 py-1.5 rounded-md shrink-0"
          >
            Удалить все
          </button>
        </div>
      )}
      {cleanupMsg && <p className="text-xs text-slate-500 mb-3">{cleanupMsg}</p>}

      <div className="flex gap-2 mb-3">
        <input
          type="text"
          value={newName}
          onChange={(e) => setNewName(e.target.value)}
          placeholder="Добавить имя вручную, например: Иван Петров"
          className="flex-1 bg-slate-100 border border-slate-300 rounded-md px-3 py-1.5 text-sm text-slate-900"
        />
        <button
          onClick={async () => { if (newName.trim()) { await addToRoster([newName.trim()]); setNewName(""); } }}
          className="bg-slate-300 hover:bg-slate-400 text-slate-800 text-sm px-3 py-1.5 rounded-md"
        >
          Добавить
        </button>
      </div>
      <input
        type="text"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        placeholder={`Поиск в каталоге (${roster.length})...`}
        className="w-full bg-slate-100 border border-slate-300 rounded-md px-3 py-1.5 text-sm text-slate-900 mb-3"
      />
      <div className="max-h-40 overflow-y-auto text-sm text-slate-600 space-y-0.5">
        {filtered.map((n) => (
          <div key={n} className="flex items-center justify-between group hover:bg-slate-100/40 rounded px-1">
            <span className={n.includes("/") ? "text-amber-600" : ""}>{n}</span>
            <button
              onClick={() => removeFromRoster([n])}
              className="text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity"
              title="Удалить из каталога"
            >
              <X size={13} />
            </button>
          </div>
        ))}
        {filtered.length === 0 && <p className="text-slate-400">Ничего не найдено.</p>}
      </div>
    </div>
  );
}

function AdminGate({ onUnlock }) {
  const [code, setCode] = useState("");
  const [wrong, setWrong] = useState(false);
  return (
    <div className="max-w-sm mx-auto text-center py-16">
      <Lock className="mx-auto mb-4 text-slate-500" size={28} />
      <p className="text-slate-600 text-sm mb-4">Введите код доступа к админке</p>
      <input
        type="password"
        value={code}
        onChange={(e) => { setCode(e.target.value); setWrong(false); }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            if (code === ADMIN_CODE) onUnlock();
            else setWrong(true);
          }
        }}
        className="w-full bg-slate-100 border border-slate-300 rounded-md px-3 py-2 text-sm text-slate-900 text-center mb-3"
        placeholder="код"
      />
      <button
        onClick={() => { if (code === ADMIN_CODE) onUnlock(); else setWrong(true); }}
        style={{ backgroundColor: BRAND_RED }}
        className="hover:brightness-110 text-white font-medium text-sm px-4 py-2 rounded-md"
      >
        Войти
      </button>
      {wrong && <p className="text-red-600 text-xs mt-3">Неверный код.</p>}
    </div>
  );
}

export default function RatingSite() {
  const { tournaments, cutoffs, roster, loaded, saveTournaments, saveCutoff, addToRoster, removeFromRoster, restoreAll, titlesEnabled, setTitlesEnabled } = useStorage();
  const [tab, setTab] = useState("rating");
  const [adminAuthed, setAdminAuthed] = useState(false);

  return (
    <div className="min-h-screen bg-white text-slate-900">
      <div className="h-1" style={{ background: `linear-gradient(90deg, ${BRAND_BLUE}, ${BRAND_RED})` }} />
      <div className="max-w-3xl mx-auto px-4 py-8">
        <header className="flex flex-wrap items-center justify-between gap-3 mb-8">
          <div className="flex items-center gap-2.5 min-w-0">
            <img src={LOGO_DATA_URI} alt="Водокачка" className="w-9 h-9 rounded-full shrink-0" />
            <h1 className="text-base sm:text-lg font-bold tracking-tight leading-tight">Чемпионат Водокачки <span className="text-slate-500">·</span> Рейтинг участников</h1>
          </div>
          <div className="flex gap-1 bg-slate-100 rounded-lg p-1">
            <button
              onClick={() => setTab("rating")}
              style={tab === "rating" ? { backgroundColor: BRAND_BLUE } : undefined}
              className={`px-3 py-1.5 rounded-md text-sm flex items-center gap-1.5 ${tab === "rating" ? "text-white" : "text-slate-600 hover:text-slate-800"}`}
            >
              <Trophy size={14} /> Рейтинг
            </button>
            <button
              onClick={() => setTab("matrix")}
              style={tab === "matrix" ? { backgroundColor: BRAND_BLUE } : undefined}
              className={`px-3 py-1.5 rounded-md text-sm flex items-center gap-1.5 ${tab === "matrix" ? "text-white" : "text-slate-600 hover:text-slate-800"}`}
            >
              <Table size={14} /> Шахматка
            </button>
          </div>
        </header>

        {!loaded ? (
          <p className="text-slate-500 text-sm">Загрузка...</p>
        ) : tab === "rating" ? (
          <PublicView tournaments={tournaments} cutoffs={cutoffs} isAdmin={adminAuthed} titlesEnabled={titlesEnabled} />
        ) : tab === "matrix" ? (
          <MatrixView tournaments={tournaments} cutoffs={cutoffs} />
        ) : adminAuthed ? (
          <div>
            <button
              onClick={() => setTab("rating")}
              className="text-slate-500 hover:text-slate-700 text-sm mb-5 flex items-center gap-1"
            >
              ← Назад к рейтингу
            </button>
            <AdminImport tournaments={tournaments} saveTournaments={saveTournaments} saveCutoff={saveCutoff} cutoffs={cutoffs} roster={roster} addToRoster={addToRoster} removeFromRoster={removeFromRoster} restoreAll={restoreAll} titlesEnabled={titlesEnabled} setTitlesEnabled={setTitlesEnabled} />
          </div>
        ) : (
          <div>
            <button
              onClick={() => setTab("rating")}
              className="text-slate-500 hover:text-slate-700 text-sm mb-5 flex items-center gap-1"
            >
              ← Назад к рейтингу
            </button>
            <AdminGate onUnlock={() => setAdminAuthed(true)} />
          </div>
        )}

        {tab !== "admin" && (
          <div className="mt-16 pt-6 border-t border-slate-100 flex justify-center">
            <button
              onClick={() => setTab("admin")}
              className="text-slate-300 hover:text-slate-600 text-xs flex items-center gap-1.5"
            >
              <Shield size={12} /> Админка
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
