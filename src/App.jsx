import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Upload, Lock, Trophy, Users, ChevronRight, Trash2, AlertCircle, CheckCircle2, Shield, History, Search, X, Table, Pencil, Check } from "lucide-react";
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

function computeBracket(wb, players) {
  const ws = wb.Sheets["5_Сетка плей-офф"];
  if (!ws) throw new Error('Не найден лист "5_Сетка плей-офф"');
  const data = sheetToRows(ws);
  let i = 0;
  while (i < data.length) {
    const title = String((data[i] || [])[1] || "").trim();
    const category = BLOCK_TITLE_TO_CATEGORY[title];
    if (category) {
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
          }
        }
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
      });
      finalNames.forEach((name) => {
        const p = players[name] || (players[name] = newRecord(name));
        if (category !== "intertoto") p.final[category] = true;
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

function parseTournamentFile(wb) {
  const { G, vsSize } = readSettings(wb);
  if (!G) throw new Error('Не удалось найти "кол-во матчей" на листе "0_Настройки"');
  if (!vsSize) throw new Error('Не удалось найти размер верхней сетки на листе "0_Настройки"');

  const players = {};
  computeGroupStats(wb, players);
  computePlayoffStats(wb, players);
  computeBracket(wb, players);

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
    return {
      name: p.name,
      played,
      O: round3(O),
      norm: round3(norm),
      pct: round3(Math.max(pct, MIN_RATING)),
      wins, draws, losses, goalsFor, goalsAgainst,
    };
  });
  rows.sort((a, b) => b.pct - a.pct);
  return { rows, meta: { G, vsSize, R, idealO: round3(idealO), idealNorm: round3(idealNorm) } };
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
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    let cancelled = false;
    async function load() {
      const nextT = { solo: [], pair: [], retro: [] };
      const nextC = { solo: null, pair: null, retro: null };
      let nextRoster = [];
      for (const f of FORMATS) {
        const tRes = await storageGet(`tournaments:${f.id}`);
        if (tRes && tRes.value) nextT[f.id] = tRes.value;
        const cRes = await storageGet(`cutoff:${f.id}`);
        if (cRes && cRes.value) nextC[f.id] = cRes.value.id;
      }
      const rRes = await storageGet("roster");
      if (rRes && rRes.value) nextRoster = rRes.value;
      if (!cancelled) {
        setTournaments(nextT);
        setCutoffs(nextC);
        setRoster(nextRoster);
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

  return { tournaments, cutoffs, roster, loaded, saveTournaments, saveCutoff, addToRoster, removeFromRoster, restoreAll };
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
// сырые данные есть) + список турниров с их индивидуальными показателями.
function computePlayerDetail(tournaments, playerName) {
  const totals = { wins: 0, draws: 0, losses: 0, goalsFor: 0, goalsAgainst: 0, matches: 0, hasStats: false };
  const history = [];
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
    history.push({ tournamentName: t.name, ...row, hasStats });
  });
  history.reverse(); // сначала последние турниры
  return { totals, history };
}

function PlayerDetailModal({ playerName, tournaments, onClose }) {
  const detail = useMemo(() => computePlayerDetail(tournaments, playerName), [tournaments, playerName]);
  const { totals, history } = detail;

  return (
    <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div
        className="bg-slate-100 border border-slate-300 rounded-2xl w-full sm:max-w-lg max-h-[85dvh] overflow-y-auto"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="sticky top-0 bg-slate-100 border-b border-slate-200 px-5 py-4 flex items-center justify-between">
          <h2 className="text-slate-900 font-semibold text-base">{playerName}</h2>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-700">
            <X size={18} />
          </button>
        </div>

        <div className="px-5 py-4">
          <h3 className="text-xs uppercase tracking-wide text-slate-500 mb-2">Итого за все турниры</h3>
          {totals.hasStats ? (
            <div className="grid grid-cols-3 gap-2 mb-2">
              <StatBox label="Матчей" value={totals.matches} />
              <StatBox label="Победы" value={totals.wins} accent="text-green-600" />
              <StatBox label="Ничьи" value={totals.draws} />
              <StatBox label="Поражения" value={totals.losses} accent="text-red-600" />
              <StatBox label="Голы забито" value={totals.goalsFor} />
              <StatBox label="Голы пропущено" value={totals.goalsAgainst} />
            </div>
          ) : (
            <p className="text-sm text-slate-500 mb-2">
              Подробной статистики (В/Н/П, голы) пока нет — эти турниры были перенесены только по итоговому рейтингу.
            </p>
          )}

          <h3 className="text-xs uppercase tracking-wide text-slate-500 mt-5 mb-2">По турнирам</h3>
          <div className="space-y-2">
            {history.map((h, i) => (
              <div key={i} className="bg-slate-200/50 rounded-lg px-3 py-2.5">
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
              </div>
            ))}
          </div>
        </div>
      </div>
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

function MatrixView({ tournaments, cutoffs }) {
  const [format, setFormat] = useState("solo");
  const list = tournaments[format] || [];
  const cutoffId = cutoffs[format];

  // Строки/сортировка — те же игроки и тот же порядок, что и в "Публичном" рейтинге.
  const standings = useMemo(() => computeStandings(list, "public", cutoffId), [list, cutoffId]);

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
                  <th className="sticky left-0 bg-slate-200 z-10 px-3 py-2 text-left font-medium whitespace-nowrap">Игрок</th>
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
                    <td className="sticky left-0 bg-white px-3 py-1.5 text-slate-900 font-medium whitespace-nowrap">{s.name}</td>
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

function PublicView({ tournaments, cutoffs, isAdmin }) {
  const [format, setFormat] = useState("solo");
  const [mode, setMode] = useState("public");
  const [query, setQuery] = useState("");
  const [selectedPlayer, setSelectedPlayer] = useState(null);

  const effectiveMode = isAdmin ? mode : "public";
  const list = tournaments[format] || [];
  const standings = useMemo(() => {
    if (effectiveMode === "last3") return computeLast3EditionsStandings(list, 3);
    return computeStandings(list, effectiveMode, cutoffs[format]);
  }, [list, effectiveMode, cutoffs, format]);
  const cutoffTournament = list.find((t) => t.id === cutoffs[format]);

  return (
    <div>
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

      {isAdmin && (
        <div className="flex flex-wrap items-center gap-2 mb-5 text-sm">
          <button
            onClick={() => setMode("all")}
            className={`px-3 py-1.5 rounded-md ${mode === "all" ? "bg-slate-300 text-white" : "text-slate-600 hover:text-slate-800"}`}
          >
            За всё время (только вы)
          </button>
          <button
            onClick={() => setMode("public")}
            disabled={!cutoffs[format]}
            className={`px-3 py-1.5 rounded-md disabled:opacity-30 disabled:cursor-not-allowed ${mode === "public" ? "bg-slate-300 text-white" : "text-slate-600 hover:text-slate-800"}`}
          >
            Публичный (актуальные)
          </button>
          <button
            onClick={() => setMode("last3")}
            className={`px-3 py-1.5 rounded-md ${mode === "last3" ? "bg-slate-300 text-white" : "text-slate-600 hover:text-slate-800"}`}
          >
            Топ-30 за последние 3 турнира
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

      <p className="text-xs text-slate-500 mb-1">
        {effectiveMode === "last3"
          ? `Топ-30 по среднему % за личные последние 3 турнира каждого — но только те, кто играл хотя бы в одном из последних 3 турниров формата «${FORMATS.find(f=>f.id===format)?.label}».`
          : "Очки рейтинга — это процент от максимально возможного количества набранных очков за один турнир. Рейтинг считается за 3 последних турнирах для каждого участника, чтобы отражать его актуальную форму."}
      </p>
      {effectiveMode === "public" && cutoffTournament && (
        <p className="text-xs text-slate-400 mb-4">
          В рейтинге учитываются только участники начиная с турнира «{cutoffTournament.name}».
        </p>
      )}
      {effectiveMode !== "public" && <div className="mb-4" />}

      <RatingTable standings={standings} query={query} onSelectPlayer={setSelectedPlayer} />

      {selectedPlayer && (
        <PlayerDetailModal
          playerName={selectedPlayer}
          tournaments={list}
          onClose={() => setSelectedPlayer(null)}
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

function TournamentRow({ index, tournament, onRename, onRemove }) {
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
          <span className="text-slate-700 truncate">
            <span className="text-slate-400 tabular-nums mr-2">{index + 1}.</span>{tournament.name}
          </span>
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

function AdminImport({ tournaments, saveTournaments, saveCutoff, cutoffs, roster, addToRoster, removeFromRoster, restoreAll }) {
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
      <BackupPanel tournaments={tournaments} cutoffs={cutoffs} roster={roster} restoreAll={restoreAll} />

      <HistoryImport tournaments={tournaments} saveTournaments={saveTournaments} addToRoster={addToRoster} />

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

            <div className="max-h-96 overflow-y-auto rounded-lg border border-slate-300">
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
                    <tr key={r.name} className="border-t border-slate-200 text-slate-700">
                      <td className="py-1.5 px-3">{r.name}</td>
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
            <p className="text-xs text-slate-400 mt-1">Значения "%" можно поправить вручную перед сохранением — например, если нужно скорректировать спорный случай.</p>
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
                  <TournamentRow key={t.id} index={idx} tournament={t} onRename={(newName) => renameTournament(f.id, t.id, newName)} onRemove={() => removeTournament(f.id, t.id)} />
                ))}
              </ol>
            )}
          </div>
        ))}
      </div>

      <RosterPanel roster={roster} addToRoster={addToRoster} removeFromRoster={removeFromRoster} />
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
  const { tournaments, cutoffs, roster, loaded, saveTournaments, saveCutoff, addToRoster, removeFromRoster, restoreAll } = useStorage();
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
          <PublicView tournaments={tournaments} cutoffs={cutoffs} isAdmin={adminAuthed} />
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
            <AdminImport tournaments={tournaments} saveTournaments={saveTournaments} saveCutoff={saveCutoff} cutoffs={cutoffs} roster={roster} addToRoster={addToRoster} removeFromRoster={removeFromRoster} restoreAll={restoreAll} />
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
