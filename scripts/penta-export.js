#!/usr/bin/env node
"use strict";

const fs = require("node:fs/promises");
const fssync = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const dotenv = require("dotenv");
const ExcelJS = require("exceljs");
const { chromium } = require("playwright");

dotenv.config();

const REQUIRED_ENV = ["PENTA_USER", "PENTA_PASS", "PENTA_BASE_URL"];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) {
    console.error(`[config] Falta variable de entorno requerida: ${key}`);
    process.exit(1);
  }
}

const CONFIG = {
  user: process.env.PENTA_USER,
  pass: process.env.PENTA_PASS,
  baseUrl: process.env.PENTA_BASE_URL.replace(/\/+$/, ""),
  headless: process.env.HEADLESS !== "false",
  timeoutMs: Number(process.env.PLAYWRIGHT_TIMEOUT_MS || 45000),
  maxFilterCombos: Number(process.env.MAX_FILTER_COMBOS || 1000),
  maxPagesPerFilterSet: Number(process.env.MAX_PAGES_PER_FILTER_SET || 500),
};

const ROOT_DIR = path.resolve(__dirname, "..");
const DATA_DIR = path.join(ROOT_DIR, "data");
const MODULES = [
  {
    key: "importadores_argentina",
    prettyName: "Importadores en Argentina",
    candidates: [
      "/home/formulario/AR/importDetalladas",
      "/home/formulario/AR/importacionesDetalladas",
      "/home/formulario/AR/importadores",
    ],
    menuRegex: /(import|importaciones|importador)/i,
  },
  {
    key: "exportadores_argentina",
    prettyName: "Exportadores de Argentina",
    candidates: [
      "/home/formulario/AR/exportDetalladas",
      "/home/formulario/AR/exportacionesDetalladas",
      "/home/formulario/AR/exportadores",
    ],
    menuRegex: /(export|exportaciones|exportador)/i,
  },
];

const globalEvents = [];

function nowIso() {
  return new Date().toISOString();
}

function safeSlug(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80) || "na";
}

function isTruthyString(v) {
  return typeof v === "string" && v.trim().length > 0;
}

function summarizeError(error) {
  if (!error) return "unknown_error";
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`;
  }
  return String(error);
}

function hashRow(row) {
  return crypto.createHash("sha1").update(JSON.stringify(row)).digest("hex");
}

function createLogger(moduleKey, localLogs) {
  return (message, level = "info", extra = undefined) => {
    const event = {
      ts: nowIso(),
      level,
      module: moduleKey,
      message,
      extra: extra ? JSON.stringify(extra) : "",
    };
    localLogs.push(event);
    globalEvents.push(event);
    const printable =
      extra !== undefined ? `${message} | ${JSON.stringify(extra)}` : message;
    console.log(`[${event.ts}] [${level.toUpperCase()}] [${moduleKey}] ${printable}`);
  };
}

async function ensureDirectories() {
  await fs.mkdir(DATA_DIR, { recursive: true });
}

async function saveDebugArtifacts(page, label, log) {
  const ts = nowIso().replace(/[:.]/g, "-");
  const base = path.join(DATA_DIR, `${safeSlug(label)}_${ts}`);
  const screenshotPath = `${base}.png`;
  const htmlPath = `${base}.html`;

  try {
    await page.screenshot({ path: screenshotPath, fullPage: true });
    log("Screenshot de error guardado", "warn", { screenshotPath });
  } catch (error) {
    log("No se pudo guardar screenshot", "warn", { error: summarizeError(error) });
  }

  try {
    const html = await page.content();
    await fs.writeFile(htmlPath, html, "utf8");
    log("HTML de error guardado", "warn", { htmlPath });
  } catch (error) {
    log("No se pudo guardar HTML", "warn", { error: summarizeError(error) });
  }
}

async function retry(action, options = {}) {
  const {
    retries = 3,
    delayMs = 1000,
    onRetry = () => {},
    actionName = "action",
  } = options;

  let lastError;
  for (let attempt = 1; attempt <= retries; attempt += 1) {
    try {
      return await action(attempt);
    } catch (error) {
      lastError = error;
      if (attempt < retries) {
        await onRetry(attempt, error);
        await new Promise((resolve) => setTimeout(resolve, delayMs * attempt));
      }
    }
  }

  throw new Error(`${actionName} failed: ${summarizeError(lastError)}`);
}

async function waitForPageSettled(page) {
  await page.waitForLoadState("domcontentloaded", { timeout: 20000 }).catch(() => {});
  await page.waitForLoadState("networkidle", { timeout: 12000 }).catch(() => {});
}

async function readElementDescriptor(locator) {
  return locator.evaluate((node) => {
    const text = (node.innerText || node.textContent || "").trim();
    return {
      text,
      value: node.value || "",
      title: node.getAttribute("title") || "",
      ariaLabel: node.getAttribute("aria-label") || "",
      role: node.getAttribute("role") || "",
    };
  });
}

async function findFirstClickableByRegex(page, regex) {
  const candidates = page.locator(
    'button, [role="button"], a, input[type="button"], input[type="submit"]'
  );
  const count = await candidates.count();
  for (let i = 0; i < count; i += 1) {
    const item = candidates.nth(i);
    if (!(await item.isVisible().catch(() => false))) continue;
    if (!(await item.isEnabled().catch(() => true))) continue;
    const desc = await readElementDescriptor(item).catch(() => null);
    if (!desc) continue;
    const haystack = `${desc.text} ${desc.value} ${desc.title} ${desc.ariaLabel}`.trim();
    if (regex.test(haystack)) {
      return item;
    }
  }
  return null;
}

async function maybeLogin(page, log) {
  const loginNeeded = await page
    .evaluate(() => {
      const url = window.location.href.toLowerCase();
      const hasPassword = Boolean(document.querySelector('input[type="password"]'));
      const hasLoginWord = /login|signin|ingresar|sesion|session/.test(url);
      return hasPassword || hasLoginWord;
    })
    .catch(() => false);

  if (!loginNeeded) {
    log("Sesion ya autenticada o login no requerido");
    return;
  }

  log("Login detectado, completando credenciales");

  const userInputSelectors = [
    'input[name*="user" i]',
    'input[id*="user" i]',
    'input[name*="mail" i]',
    'input[type="email"]',
    'input[placeholder*="usuario" i]',
    'input[placeholder*="correo" i]',
    "input[type='text']",
  ];
  const passInputSelectors = [
    'input[name*="pass" i]',
    'input[id*="pass" i]',
    'input[placeholder*="clave" i]',
    'input[placeholder*="contras" i]',
    'input[type="password"]',
  ];

  const userLocator = page.locator(userInputSelectors.join(","));
  const passLocator = page.locator(passInputSelectors.join(","));

  if (!(await userLocator.first().isVisible().catch(() => false))) {
    throw new Error("No se encontro input de usuario en login");
  }
  if (!(await passLocator.first().isVisible().catch(() => false))) {
    throw new Error("No se encontro input de password en login");
  }

  await userLocator.first().fill(CONFIG.user, { timeout: 12000 });
  await passLocator.first().fill(CONFIG.pass, { timeout: 12000 });

  const submit = await findFirstClickableByRegex(
    page,
    /(ingresar|iniciar|acceder|login|entrar|submit)/i
  );

  if (submit) {
    await Promise.all([
      page.waitForLoadState("networkidle", { timeout: 30000 }).catch(() => {}),
      submit.click({ timeout: 12000 }),
    ]);
  } else {
    await passLocator.first().press("Enter");
    await page.waitForLoadState("networkidle", { timeout: 30000 }).catch(() => {});
  }

  const stillLogin = await page
    .evaluate(() => {
      const hasPassword = Boolean(document.querySelector('input[type="password"]'));
      const url = window.location.href.toLowerCase();
      return hasPassword && /login|signin|ingresar|sesion|session/.test(url);
    })
    .catch(() => true);

  if (stillLogin) {
    throw new Error("No se pudo autenticar: pantalla de login continua visible");
  }

  log("login ok");
}

async function maybeOpenAdvancedSearch(page, log) {
  const trigger = await findFirstClickableByRegex(
    page,
    /(busqueda avanzada|búsqueda avanzada|advanced search|filtros|mostrar filtros)/i
  );
  if (!trigger) return false;

  try {
    await trigger.click({ timeout: 10000 });
    await waitForPageSettled(page);
    log("Búsqueda avanzada/filtros expandida");
    return true;
  } catch {
    return false;
  }
}

async function detectFilters(page, log) {
  const filterInfo = await page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const getLabel = (el) => {
      const id = el.id;
      if (id) {
        const byFor = document.querySelector(`label[for="${id}"]`);
        if (byFor) return normalize(byFor.textContent);
      }
      const withinLabel = el.closest("label");
      if (withinLabel) return normalize(withinLabel.textContent);
      const wrapper = el.closest("div, section, form, td, th");
      if (wrapper) {
        const label = wrapper.querySelector("label, strong, .label, .form-label");
        if (label) return normalize(label.textContent);
      }
      return normalize(
        el.getAttribute("aria-label") ||
          el.getAttribute("placeholder") ||
          el.name ||
          el.id ||
          ""
      );
    };
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return (
        style &&
        style.visibility !== "hidden" &&
        style.display !== "none" &&
        rect.width > 0 &&
        rect.height > 0
      );
    };

    const selectNodes = Array.from(document.querySelectorAll("select")).filter(visible);
    const inputNodes = Array.from(
      document.querySelectorAll('input[type="date"], input[placeholder*="fecha" i]')
    ).filter(visible);

    const selects = selectNodes.map((el, index) => {
      const options = Array.from(el.options || []).map((opt) => ({
        value: opt.value,
        text: normalize(opt.textContent),
      }));
      return {
        index,
        label: getLabel(el),
        name: el.name || "",
        id: el.id || "",
        options,
      };
    });

    const dateInputs = inputNodes.map((el, index) => ({
      index,
      label: getLabel(el),
      name: el.name || "",
      id: el.id || "",
      min: el.min || "",
      max: el.max || "",
    }));

    return { selects, dateInputs };
  });

  const pageSizeRegex = /(registros|filas|mostrar|cantidad|items|resultados|por p[aá]gina|page size)/i;
  const countryRegex = /(pa[ií]s|country|origen|destino)/i;
  const periodRegex = /(periodo|per[ií]odo|fecha|a[nñ]o|anio|mes|desde|hasta)/i;

  const plans = {
    pageSize: null,
    country: [],
    period: [],
    generic: [],
    dateInputs: filterInfo.dateInputs || [],
    allSelectsCount: (filterInfo.selects || []).length,
  };

  for (const selectMeta of filterInfo.selects || []) {
    const enrichedLabel = `${selectMeta.label} ${selectMeta.name} ${selectMeta.id}`.trim();
    const validOptions = (selectMeta.options || []).filter((opt) =>
      isTruthyString(opt.value) || isTruthyString(opt.text)
    );
    if (!validOptions.length) continue;

    if (pageSizeRegex.test(enrichedLabel)) {
      const withNumeric = validOptions
        .map((opt) => ({
          ...opt,
          number: Number(String(opt.value || opt.text).replace(/[^\d]/g, "")),
        }))
        .filter((x) => Number.isFinite(x.number) && x.number > 0)
        .sort((a, b) => a.number - b.number);
      plans.pageSize = {
        ...selectMeta,
        selectedOption:
          withNumeric.length > 0 ? withNumeric[withNumeric.length - 1] : validOptions[validOptions.length - 1],
      };
      continue;
    }

    if (countryRegex.test(enrichedLabel)) {
      plans.country.push({
        ...selectMeta,
        values: validOptions,
      });
      continue;
    }

    if (periodRegex.test(enrichedLabel)) {
      plans.period.push({
        ...selectMeta,
        values: validOptions,
      });
      continue;
    }

    plans.generic.push(selectMeta);
  }

  log("filtros encontrados", "info", {
    totalSelects: (filterInfo.selects || []).length,
    totalDateInputs: (filterInfo.dateInputs || []).length,
    countryFilters: plans.country.length,
    periodFilters: plans.period.length,
    hasPageSize: Boolean(plans.pageSize),
  });

  return plans;
}

function uniqueOptions(options) {
  const dedup = new Map();
  for (const opt of options || []) {
    const key = `${String(opt.value || "").trim()}|${String(opt.text || "").trim()}`;
    if (!dedup.has(key)) dedup.set(key, opt);
  }
  return [...dedup.values()];
}

function isPlaceholderOption(opt) {
  const txt = `${opt.text || ""} ${opt.value || ""}`.toLowerCase();
  return /(seleccione|seleccionar|elija|choose|select|--|todos\.\.\.|all\.\.\.)/.test(txt);
}

function findBroadOption(options) {
  return (options || []).find((opt) =>
    /(todos|todas|all|any|general|global|completo|total)/i.test(
      `${opt.text || ""} ${opt.value || ""}`
    )
  );
}

function chooseEarliestOption(options) {
  const numeric = options
    .map((opt) => ({
      opt,
      n: Number(String(opt.value || opt.text).replace(/[^\d]/g, "")),
    }))
    .filter((x) => Number.isFinite(x.n) && x.n > 0)
    .sort((a, b) => a.n - b.n);
  if (numeric.length > 0) return numeric[0].opt;
  return options[0];
}

function chooseLatestOption(options) {
  const numeric = options
    .map((opt) => ({
      opt,
      n: Number(String(opt.value || opt.text).replace(/[^\d]/g, "")),
    }))
    .filter((x) => Number.isFinite(x.n) && x.n > 0)
    .sort((a, b) => a.n - b.n);
  if (numeric.length > 0) return numeric[numeric.length - 1].opt;
  return options[options.length - 1];
}

function buildFilterSets(plans, maxCombos, log) {
  let sets = [];
  const base = { selects: [], dates: [] };

  if (plans.pageSize) {
    base.selects.push({
      index: plans.pageSize.index,
      label: plans.pageSize.label,
      value: plans.pageSize.selectedOption.value || plans.pageSize.selectedOption.text,
      text: plans.pageSize.selectedOption.text || plans.pageSize.selectedOption.value,
    });
  }

  if (plans.dateInputs.length) {
    for (const dateInput of plans.dateInputs) {
      const valueFrom = dateInput.min || "1990-01-01";
      const valueTo = dateInput.max || new Date().toISOString().slice(0, 10);
      if (/desde|from|inicio/i.test(`${dateInput.label} ${dateInput.name} ${dateInput.id}`)) {
        base.dates.push({ index: dateInput.index, label: dateInput.label, value: valueFrom });
      } else if (/hasta|to|fin/i.test(`${dateInput.label} ${dateInput.name} ${dateInput.id}`)) {
        base.dates.push({ index: dateInput.index, label: dateInput.label, value: valueTo });
      } else {
        base.dates.push({ index: dateInput.index, label: dateInput.label, value: valueFrom });
      }
    }
  }

  for (const genericFilter of plans.generic || []) {
    const options = uniqueOptions((genericFilter.values || []).filter((x) => !isPlaceholderOption(x)));
    const broad = findBroadOption(options);
    if (broad) {
      base.selects.push({
        index: genericFilter.index,
        label: genericFilter.label,
        value: broad.value || broad.text,
        text: broad.text || broad.value,
      });
    }
  }

  for (const periodFilter of plans.period || []) {
    const options = uniqueOptions((periodFilter.values || []).filter((x) => !isPlaceholderOption(x)));
    if (!options.length) continue;

    const broad = findBroadOption(options);
    const label = `${periodFilter.label || ""} ${periodFilter.name || ""} ${periodFilter.id || ""}`;
    const isFrom = /(desde|from|inicio|inicial|min)/i.test(label);
    const isTo = /(hasta|to|fin|final|max)/i.test(label);

    let selected = broad;
    if (!selected && isFrom) selected = chooseEarliestOption(options);
    if (!selected && isTo) selected = chooseLatestOption(options);
    if (!selected) selected = chooseLatestOption(options);

    base.selects.push({
      index: periodFilter.index,
      label: periodFilter.label,
      value: selected.value || selected.text,
      text: selected.text || selected.value,
    });
  }

  sets = [{ selects: [...base.selects], dates: [...base.dates] }];
  for (const countryFilter of plans.country || []) {
    const allCountryOptions = uniqueOptions(
      (countryFilter.values || []).filter((x) => !isPlaceholderOption(x))
    );
    if (!allCountryOptions.length) continue;

    const broad = findBroadOption(allCountryOptions);
    const countryValues = allCountryOptions.filter((x) => x !== broad);
    const valuesToIterate = countryValues.length > 0 ? countryValues : allCountryOptions;

    const nextSets = [];
    for (const existing of sets) {
      for (const country of valuesToIterate) {
        nextSets.push({
          selects: [
            ...existing.selects,
            {
              index: countryFilter.index,
              label: countryFilter.label,
              value: country.value || country.text,
              text: country.text || country.value,
            },
          ],
          dates: [...existing.dates],
        });
        if (nextSets.length >= maxCombos) break;
      }
      if (nextSets.length >= maxCombos) break;
    }
    sets = nextSets.length > 0 ? nextSets : sets;
    if (sets.length >= maxCombos) break;
  }

  if (!sets.length) sets = [base];
  if (sets.length > maxCombos) {
    log("Combinaciones de filtros truncadas por seguridad", "warn", {
      requested: sets.length,
      kept: maxCombos,
    });
    return sets.slice(0, maxCombos);
  }
  return sets;
}

async function applyFilterSet(page, filterSet, log) {
  for (const sel of filterSet.selects || []) {
    const locator = page.locator("select").nth(sel.index);
    if (!(await locator.isVisible().catch(() => false))) continue;
    try {
      await locator.selectOption({ value: String(sel.value) }, { timeout: 10000 });
    } catch {
      try {
        await locator.selectOption({ label: String(sel.text) }, { timeout: 10000 });
      } catch (error) {
        log("No se pudo aplicar select", "warn", {
          label: sel.label,
          value: sel.value,
          error: summarizeError(error),
        });
      }
    }
  }

  for (const date of filterSet.dates || []) {
    const locator = page
      .locator('input[type="date"], input[placeholder*="fecha" i]')
      .nth(date.index);
    if (!(await locator.isVisible().catch(() => false))) continue;
    await locator.fill(date.value).catch((error) => {
      log("No se pudo completar fecha", "warn", {
        label: date.label,
        value: date.value,
        error: summarizeError(error),
      });
    });
  }

  const applyBtn = await findFirstClickableByRegex(
    page,
    /(buscar|aplicar|filtrar|consultar|actualizar|mostrar|search)/i
  );
  if (applyBtn) {
    await applyBtn.click({ timeout: 10000 }).catch(() => {});
  }

  await waitForPageSettled(page);
}

async function findNativeExportButton(page) {
  return findFirstClickableByRegex(page, /(export|excel|csv|descargar|download|reporte)/i);
}

async function downloadNativeExport(page, moduleKey, filterSet, fileIndex, log) {
  const button = await findNativeExportButton(page);
  if (!button) return null;

  const tag = (filterSet.selects || [])
    .map((s) => safeSlug(`${s.label}_${s.text || s.value}`))
    .join("__");

  const prefix = tag ? `${moduleKey}__${tag}` : moduleKey;
  let download = await Promise.all([
    page.waitForEvent("download", { timeout: 20000 }),
    button.click({ timeout: 8000 }),
  ])
    .then(([dl]) => dl)
    .catch(() => null);

  if (!download) {
    const secondChance = await findFirstClickableByRegex(
      page,
      /(excel|csv|descargar|download|reporte|export)/i
    );
    if (secondChance) {
      download = await Promise.all([
        page.waitForEvent("download", { timeout: 20000 }),
        secondChance.click({ timeout: 8000 }),
      ])
        .then(([dl]) => dl)
        .catch(() => null);
    }
  }
  if (!download) return null;

  const ext = path.extname(download.suggestedFilename() || "") || ".dat";
  const dest = path.join(DATA_DIR, `${prefix}__${fileIndex}${ext}`);
  await download.saveAs(dest);
  log("Exportación nativa descargada", "info", { file: dest });
  return dest;
}

async function parseNativeFile(filePath, log) {
  const ext = path.extname(filePath).toLowerCase();
  if (![".xlsx", ".csv"].includes(ext)) {
    log("Formato nativo no parseable por script (se conserva archivo)", "warn", { filePath });
    return [];
  }

  const workbook = new ExcelJS.Workbook();
  if (ext === ".xlsx") {
    await workbook.xlsx.readFile(filePath);
  } else {
    await workbook.csv.readFile(filePath);
  }
  const worksheet = workbook.worksheets[0];
  if (!worksheet) return [];

  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values
    .slice(1)
    .map((v, i) => (isTruthyString(v) ? String(v).trim() : `col_${i + 1}`));
  const rows = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const values = row.values.slice(1);
    if (!values.some((v) => isTruthyString(v) || typeof v === "number")) return;
    const out = {};
    for (let i = 0; i < headers.length; i += 1) {
      out[headers[i]] = values[i] ?? "";
    }
    rows.push(out);
  });

  return rows;
}

async function extractRowsFromPage(page, moduleKey, filterLabel, pageNumber) {
  return page.evaluate(
    ({ moduleKeyArg, filterLabelArg, pageNumberArg }) => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const visible = (el) => {
        const style = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        return (
          style &&
          style.visibility !== "hidden" &&
          style.display !== "none" &&
          rect.width > 0 &&
          rect.height > 0
        );
      };

      const collected = [];
      const tables = Array.from(document.querySelectorAll("table")).filter(visible);

      for (let tableIndex = 0; tableIndex < tables.length; tableIndex += 1) {
        const table = tables[tableIndex];
        const headers = Array.from(table.querySelectorAll("thead th")).map((th, i) => {
          const txt = normalize(th.textContent);
          return txt || `col_${i + 1}`;
        });

        const bodyRows = Array.from(table.querySelectorAll("tbody tr")).filter(visible);
        for (const tr of bodyRows) {
          const cells = Array.from(tr.querySelectorAll("td"));
          if (!cells.length) continue;
          const row = {
            _module: moduleKeyArg,
            _filter: filterLabelArg,
            _page: pageNumberArg,
            _table: tableIndex + 1,
          };
          for (let i = 0; i < cells.length; i += 1) {
            const key = headers[i] || `col_${i + 1}`;
            row[key] = normalize(cells[i].innerText || cells[i].textContent);
          }
          const anchor = tr.querySelector("a[href]");
          if (anchor) row._detail_href = anchor.getAttribute("href") || "";
          collected.push(row);
        }
      }

      if (collected.length > 0) return collected;

      const roleRows = Array.from(document.querySelectorAll('[role="row"]')).filter(visible);
      if (roleRows.length > 1) {
        const headerCandidates = Array.from(roleRows[0].querySelectorAll('[role="columnheader"]'));
        const headers = headerCandidates.map((h, i) => normalize(h.textContent) || `col_${i + 1}`);
        for (let i = 1; i < roleRows.length; i += 1) {
          const rr = roleRows[i];
          const cells = Array.from(rr.querySelectorAll('[role="gridcell"], [role="cell"]'));
          if (!cells.length) continue;
          const row = {
            _module: moduleKeyArg,
            _filter: filterLabelArg,
            _page: pageNumberArg,
            _table: 1,
          };
          for (let c = 0; c < cells.length; c += 1) {
            const key = headers[c] || `col_${c + 1}`;
            row[key] = normalize(cells[c].textContent);
          }
          const anchor = rr.querySelector("a[href]");
          if (anchor) row._detail_href = anchor.getAttribute("href") || "";
          collected.push(row);
        }
      }

      return collected;
    },
    { moduleKeyArg: moduleKey, filterLabelArg: filterLabel, pageNumberArg: pageNumber }
  );
}

async function getRowsCount(page) {
  const tableRows = await page.locator("table tbody tr:visible").count().catch(() => 0);
  if (tableRows > 0) return tableRows;
  const roleRows = await page.locator('[role="row"]:visible').count().catch(() => 0);
  return roleRows;
}

async function getDataSignature(page) {
  return page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll("table tbody tr")).slice(0, 3);
    const txt = rows
      .map((r) => (r.innerText || "").replace(/\s+/g, " ").trim())
      .join("||");
    return `${window.location.href}::${txt}`;
  });
}

async function tryAdvancePagination(page, log) {
  const signatureBefore = await getDataSignature(page).catch(() => "");
  const rowsBefore = await getRowsCount(page).catch(() => 0);

  const loadMore = await findFirstClickableByRegex(
    page,
    /(cargar m[aá]s|mostrar m[aá]s|load more|ver m[aá]s)/i
  );
  if (loadMore) {
    await loadMore.click({ timeout: 10000 }).catch(() => {});
    await waitForPageSettled(page);
    const rowsAfter = await getRowsCount(page).catch(() => 0);
    if (rowsAfter > rowsBefore) {
      log("Paginación por load more detectada", "info", { rowsBefore, rowsAfter });
      return true;
    }
  }

  const nextButton = await findFirstClickableByRegex(
    page,
    /(siguiente|next|proxima|pr[oó]xima|›|»)/i
  );
  if (nextButton) {
    await nextButton.click({ timeout: 10000 }).catch(() => {});
    await waitForPageSettled(page);
    const signatureAfter = await getDataSignature(page).catch(() => "");
    if (signatureAfter && signatureAfter !== signatureBefore) {
      return true;
    }
  }

  const rowsAfterScrollAttempt = await getRowsCount(page).catch(() => rowsBefore);
  await page.mouse.wheel(0, 2500).catch(() => {});
  await waitForPageSettled(page);
  const rowsAfterScroll = await getRowsCount(page).catch(() => rowsAfterScrollAttempt);
  if (rowsAfterScroll > rowsAfterScrollAttempt) {
    log("Paginación/infinite scroll detectada", "info", {
      rowsAfterScrollAttempt,
      rowsAfterScroll,
    });
    return true;
  }

  return false;
}

async function extractDetailsFromHref(context, href, log) {
  if (!isTruthyString(href)) return {};
  const absolute = href.startsWith("http") ? href : `${CONFIG.baseUrl}${href.startsWith("/") ? "" : "/"}${href}`;
  if (!absolute.startsWith(CONFIG.baseUrl)) return {};

  const detailPage = await context.newPage();
  detailPage.setDefaultTimeout(CONFIG.timeoutMs);

  try {
    await detailPage.goto(absolute, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await waitForPageSettled(detailPage);
    const detail = await detailPage.evaluate(() => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const out = {};

      const dtNodes = Array.from(document.querySelectorAll("dt"));
      for (const dt of dtNodes) {
        const dd = dt.nextElementSibling;
        if (dd && dd.tagName.toLowerCase() === "dd") {
          const key = normalize(dt.textContent);
          const value = normalize(dd.textContent);
          if (key) out[`detail_${key}`] = value;
        }
      }

      const labels = Array.from(document.querySelectorAll("label"));
      for (const lbl of labels) {
        const key = normalize(lbl.textContent);
        if (!key || out[`detail_${key}`]) continue;
        const sibling = lbl.nextElementSibling;
        if (sibling) {
          out[`detail_${key}`] = normalize(sibling.textContent);
        }
      }

      return out;
    });
    await detailPage.close();
    return detail;
  } catch (error) {
    await detailPage.close();
    log("No se pudo extraer detalle individual", "warn", {
      href: absolute,
      error: summarizeError(error),
    });
    return {};
  }
}

async function manualScrapeAll(page, context, moduleKey, filterSet, log) {
  const seen = new Set();
  const allRows = [];
  let pagesTraversed = 0;

  const filterLabel = (filterSet.selects || [])
    .map((s) => `${s.label || "filtro"}=${s.text || s.value}`)
    .join(" | ") || "sin_filtro";

  for (;;) {
    pagesTraversed += 1;
    if (pagesTraversed > CONFIG.maxPagesPerFilterSet) {
      log("Corte de seguridad por exceso de páginas", "warn", {
        limit: CONFIG.maxPagesPerFilterSet,
      });
      break;
    }
    log("Página actual", "info", { page: pagesTraversed, filterLabel });

    const rows = await extractRowsFromPage(page, moduleKey, filterLabel, pagesTraversed).catch(() => []);
    for (const row of rows) {
      const hash = row.id || row.ID || row._id || hashRow(row);
      if (seen.has(hash)) continue;

      if (row._detail_href) {
        const detail = await extractDetailsFromHref(context, row._detail_href, log);
        Object.assign(row, detail);
      }

      seen.add(hash);
      allRows.push(row);
    }

    log("registros acumulados", "info", { total: allRows.length });
    const moved = await tryAdvancePagination(page, log);
    if (!moved) break;
  }

  return { rows: allRows, pagesTraversed };
}

async function writeExcelAndJson(moduleKey, rows, metadata, logs) {
  const jsonPath = path.join(DATA_DIR, `${moduleKey}.json`);
  const xlsxPath = path.join(DATA_DIR, `${moduleKey}.xlsx`);

  await fs.writeFile(jsonPath, JSON.stringify(rows, null, 2), "utf8");

  const workbook = new ExcelJS.Workbook();
  const dataSheet = workbook.addWorksheet("data");
  const metaSheet = workbook.addWorksheet("metadata");
  const logsSheet = workbook.addWorksheet("logs");

  const columns = [...new Set(rows.flatMap((r) => Object.keys(r)))];
  dataSheet.columns = columns.map((col) => ({ header: col, key: col, width: 18 }));
  for (const row of rows) dataSheet.addRow(row);

  dataSheet.views = [{ state: "frozen", ySplit: 1 }];
  dataSheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: Math.max(columns.length, 1) },
  };
  dataSheet.columns.forEach((column) => {
    let maxLen = String(column.header || "").length;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const val = cell.value == null ? "" : String(cell.value);
      if (val.length > maxLen) maxLen = val.length;
    });
    column.width = Math.min(Math.max(maxLen + 2, 12), 60);
  });

  metaSheet.columns = [
    { header: "campo", key: "campo", width: 38 },
    { header: "valor", key: "valor", width: 120 },
  ];
  Object.entries(metadata).forEach(([campo, valor]) => {
    metaSheet.addRow({
      campo,
      valor: typeof valor === "string" ? valor : JSON.stringify(valor),
    });
  });

  logsSheet.columns = [
    { header: "timestamp", key: "ts", width: 28 },
    { header: "level", key: "level", width: 10 },
    { header: "module", key: "module", width: 30 },
    { header: "message", key: "message", width: 70 },
    { header: "extra", key: "extra", width: 120 },
  ];
  logs.forEach((entry) => logsSheet.addRow(entry));

  await workbook.xlsx.writeFile(xlsxPath);
  return { jsonPath, xlsxPath };
}

async function navigateToModule(page, moduleDef, log) {
  for (const candidate of moduleDef.candidates) {
    const url = `${CONFIG.baseUrl}${candidate.startsWith("/") ? "" : "/"}${candidate}`;
    try {
      await page.goto(url, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
      await maybeLogin(page, log);
      await waitForPageSettled(page);
      log("módulo detectado", "info", { url });
      return url;
    } catch (error) {
      log("Ruta candidata no disponible", "warn", {
        url,
        error: summarizeError(error),
      });
    }
  }

  const menuItem = await findFirstClickableByRegex(page, moduleDef.menuRegex);
  if (menuItem) {
    await menuItem.click({ timeout: 10000 });
    await waitForPageSettled(page);
    const current = page.url();
    log("módulo detectado por menú/tab", "info", { url: current });
    return current;
  }

  throw new Error(`No se pudo navegar al módulo ${moduleDef.key}`);
}

async function processModule(page, context, moduleDef) {
  const moduleLogs = [];
  const errors = [];
  const nativeDownloads = [];
  const log = createLogger(moduleDef.key, moduleLogs);

  const result = {
    moduleKey: moduleDef.key,
    rows: [],
    pagesTraversed: 0,
    extractionMode: "manual",
    filtersFound: {},
    filtersUsed: [],
    nativeDownloads,
    errors,
    logs: moduleLogs,
  };

  try {
    await navigateToModule(page, moduleDef, log);
    await maybeOpenAdvancedSearch(page, log);

    const filters = await detectFilters(page, log);
    result.filtersFound = filters;

    const filterSets = buildFilterSets(filters, CONFIG.maxFilterCombos, log);
    if (filterSets.length === 0) filterSets.push({ selects: [], dates: [] });
    result.filtersUsed = filterSets.map((set) => ({
      selects: set.selects?.map((s) => ({ label: s.label, value: s.value, text: s.text })) || [],
      dates: set.dates || [],
    }));

    const hasNativeExport = Boolean(await findNativeExportButton(page));
    log(
      hasNativeExport
        ? "exportación nativa encontrada"
        : "exportación nativa no encontrada"
    );

    if (hasNativeExport) {
      result.extractionMode = "native_export";
      let index = 1;
      for (const filterSet of filterSets) {
        await applyFilterSet(page, filterSet, log);
        const downloaded = await retry(
          () => downloadNativeExport(page, moduleDef.key, filterSet, index, log),
          {
            retries: 2,
            delayMs: 1500,
            actionName: "downloadNativeExport",
            onRetry: (attempt, error) =>
              log("Reintento de exportación nativa", "warn", {
                attempt,
                error: summarizeError(error),
              }),
          }
        ).catch((error) => {
          errors.push(`native_export_filter_${index}: ${summarizeError(error)}`);
          return null;
        });
        if (downloaded) nativeDownloads.push(downloaded);
        index += 1;
      }

      for (const file of nativeDownloads) {
        const parsed = await parseNativeFile(file, log).catch((error) => {
          errors.push(`parse_native_${path.basename(file)}: ${summarizeError(error)}`);
          return [];
        });
        result.rows.push(...parsed);
      }
    }

    if (result.rows.length === 0) {
      result.extractionMode = "manual_scraping";
      for (const filterSet of filterSets) {
        await applyFilterSet(page, filterSet, log);
        const partial = await manualScrapeAll(
          page,
          context,
          moduleDef.key,
          filterSet,
          log
        ).catch((error) => {
          errors.push(`manual_scrape: ${summarizeError(error)}`);
          return { rows: [], pagesTraversed: 0 };
        });
        result.rows.push(...partial.rows);
        result.pagesTraversed += partial.pagesTraversed;
      }
    }
  } catch (error) {
    errors.push(summarizeError(error));
    log("Error procesando módulo", "error", { error: summarizeError(error) });
    await saveDebugArtifacts(page, `${moduleDef.key}_failure`, log);
  }

  const dedup = new Map();
  for (const row of result.rows) {
    const id = row.id || row.ID || row._id || hashRow(row);
    if (!dedup.has(id)) dedup.set(id, row);
  }
  result.rows = [...dedup.values()];
  return result;
}

async function main() {
  await ensureDirectories();

  const browser = await chromium.launch({
    headless: CONFIG.headless,
  });
  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1600, height: 1000 },
  });
  const page = await context.newPage();
  page.setDefaultTimeout(CONFIG.timeoutMs);

  console.log(`[${nowIso()}] [INFO] Iniciando extracción contra ${CONFIG.baseUrl}`);

  try {
    const landing = `${CONFIG.baseUrl}/home/formulario/AR/importDetalladas`;
    await page.goto(landing, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await maybeLogin(page, createLogger("auth", globalEvents));
    await waitForPageSettled(page);
  } catch (error) {
    const authLogger = createLogger("auth", globalEvents);
    authLogger("Fallo durante autenticación inicial", "error", {
      error: summarizeError(error),
    });
    await saveDebugArtifacts(page, "login_failure", authLogger);
    await browser.close();
    process.exit(1);
  }

  for (const moduleDef of MODULES) {
    const moduleResult = await processModule(page, context, moduleDef);
    const fields = [...new Set(moduleResult.rows.flatMap((r) => Object.keys(r)))];
    const metadata = {
      extraction_datetime: nowIso(),
      base_url: CONFIG.baseUrl,
      module: moduleDef.key,
      module_name: moduleDef.prettyName,
      total_records: moduleResult.rows.length,
      pages_traversed: moduleResult.pagesTraversed,
      filters_used: moduleResult.filtersUsed,
      filters_found: {
        totalSelects: moduleResult.filtersFound.allSelectsCount || 0,
        totalDateInputs: moduleResult.filtersFound.dateInputs?.length || 0,
      },
      extraction_mode: moduleResult.extractionMode,
      native_downloads: moduleResult.nativeDownloads,
      errors_detected: moduleResult.errors,
      available_fields: fields,
    };

    const { jsonPath, xlsxPath } = await writeExcelAndJson(
      moduleDef.key,
      moduleResult.rows,
      metadata,
      moduleResult.logs
    );

    console.log(
      `[${nowIso()}] [INFO] excel generado correctamente | ${moduleDef.key} | ${xlsxPath}`
    );
    console.log(
      `[${nowIso()}] [INFO] json generado correctamente  | ${moduleDef.key} | ${jsonPath}`
    );
  }

  await browser.close();
  console.log(`[${nowIso()}] [INFO] Proceso finalizado`);
}

main().catch((error) => {
  console.error(`[fatal] ${summarizeError(error)}`);
  process.exit(1);
});

