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
  timeoutMs: Number(process.env.PLAYWRIGHT_TIMEOUT_MS || 50000),
  maxPagesPerModule: Number(process.env.MAX_PAGES_PER_MODULE || 500),
  maxFilterCombos: Number(process.env.MAX_FILTER_COMBOS || 500),
};

const ROOT_DIR = path.resolve(__dirname, "..");
const DATA_DIR = path.join(ROOT_DIR, "data");
const MODULES = [
  {
    key: "importadores",
    urlPath: "/home/formulario/AR/importDetalladas",
    companyType: "consignee",
    debugBaseName: "debug-importadores",
  },
  {
    key: "exportadores",
    urlPath: "/home/formulario/AR/exportDetalladas",
    companyType: "shipper",
    debugBaseName: "debug-exportadores",
  },
];

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
    .slice(0, 120) || "na";
}

function isTruthy(v) {
  return v !== null && v !== undefined && String(v).trim().length > 0;
}

function summarizeError(error) {
  if (!error) return "unknown_error";
  if (error instanceof Error) return `${error.name}: ${error.message}`;
  return String(error);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function makeLogger(scope, sink) {
  return (message, level = "info", extra = undefined) => {
    const entry = {
      ts: nowIso(),
      level,
      scope,
      message,
      extra: extra ? JSON.stringify(extra) : "",
    };
    sink.push(entry);
    const extraTxt = extra ? ` | ${JSON.stringify(extra)}` : "";
    console.log(`[${entry.ts}] [${level.toUpperCase()}] [${scope}] ${message}${extraTxt}`);
  };
}

async function ensureDirectories() {
  if (!fssync.existsSync(DATA_DIR)) {
    await fs.mkdir(DATA_DIR, { recursive: true });
  }
}

async function waitForSettled(page) {
  await page.waitForLoadState("domcontentloaded", { timeout: 30000 }).catch(() => {});
  await page.waitForLoadState("networkidle", { timeout: 15000 }).catch(() => {});
}

async function retry(action, options = {}) {
  const { retries = 3, delayMs = 1200, actionName = "action", onRetry = () => {} } = options;
  let lastErr;
  for (let attempt = 1; attempt <= retries; attempt += 1) {
    try {
      return await action(attempt);
    } catch (error) {
      lastErr = error;
      if (attempt < retries) {
        await onRetry(attempt, error);
        await sleep(delayMs * attempt);
      }
    }
  }
  throw new Error(`${actionName} failed: ${summarizeError(lastErr)}`);
}

async function findVisibleClickable(page, regex) {
  const loc = page.locator(
    'button, [role="button"], a, input[type="button"], input[type="submit"], .btn'
  );
  const count = await loc.count().catch(() => 0);
  for (let i = 0; i < count; i += 1) {
    const item = loc.nth(i);
    const visible = await item.isVisible().catch(() => false);
    if (!visible) continue;
    const enabled = await item.isEnabled().catch(() => true);
    if (!enabled) continue;
    const text = await item
      .evaluate((el) => {
        const v = el.value || "";
        const t = el.innerText || el.textContent || "";
        const a = el.getAttribute("aria-label") || "";
        const ttl = el.getAttribute("title") || "";
        return `${v} ${t} ${a} ${ttl}`.replace(/\s+/g, " ").trim();
      })
      .catch(() => "");
    if (regex.test(text)) return item;
  }
  return null;
}

async function maybeLogin(page, log) {
  const loginNeeded = await page
    .evaluate(() => {
      const url = window.location.href.toLowerCase();
      const passInput = Boolean(document.querySelector('input[type="password"]'));
      const loginWord = /login|signin|ingresar|sesion|session/.test(url);
      return passInput || loginWord;
    })
    .catch(() => false);

  if (!loginNeeded) {
    log("Sesión reutilizada / login no requerido");
    return;
  }

  log("Login detectado, completando credenciales");

  const userLocator = page.locator(
    [
      'input[name*="user" i]',
      'input[id*="user" i]',
      'input[name*="mail" i]',
      'input[type="email"]',
      'input[placeholder*="usuario" i]',
      "input[type='text']",
    ].join(",")
  );

  const passLocator = page.locator(
    [
      'input[name*="pass" i]',
      'input[id*="pass" i]',
      'input[placeholder*="clave" i]',
      'input[placeholder*="contras" i]',
      'input[type="password"]',
    ].join(",")
  );

  if (!(await userLocator.first().isVisible().catch(() => false))) {
    throw new Error("Input de usuario no encontrado");
  }
  if (!(await passLocator.first().isVisible().catch(() => false))) {
    throw new Error("Input de contraseña no encontrado");
  }

  await userLocator.first().fill(CONFIG.user, { timeout: 15000 });
  await passLocator.first().fill(CONFIG.pass, { timeout: 15000 });

  const submit = await findVisibleClickable(
    page,
    /(ingresar|iniciar|acceder|login|entrar|submit)/i
  );
  if (submit) {
    await Promise.all([
      page.waitForLoadState("networkidle", { timeout: 35000 }).catch(() => {}),
      submit.click({ timeout: 12000 }),
    ]);
  } else {
    await passLocator.first().press("Enter");
    await page.waitForLoadState("networkidle", { timeout: 35000 }).catch(() => {});
  }

  const stillLogin = await page
    .evaluate(() => {
      const url = window.location.href.toLowerCase();
      const passInput = Boolean(document.querySelector('input[type="password"]'));
      return passInput && /login|signin|ingresar|sesion|session/.test(url);
    })
    .catch(() => true);

  if (stillLogin) {
    throw new Error("Login no completado: la vista de autenticación sigue visible");
  }
  log("login ok");
}

function createNetworkCollector(page, log) {
  const collected = [];
  const keywordRegex = /(import|export|detalle|search|data|grid|formulario)/i;

  const onResponse = async (response) => {
    try {
      const req = response.request();
      const url = response.url();
      const method = req.method();
      const resourceType = req.resourceType();
      const contentType = response.headers()["content-type"] || "";

      if (!keywordRegex.test(url) && !/(xhr|fetch)/i.test(resourceType)) return;

      const item = {
        ts: nowIso(),
        url,
        method,
        status: response.status(),
        resourceType,
        contentType,
        payloadType: "none",
        sampleSize: 0,
        parsedRecords: [],
      };

      const isJson = /application\/json|text\/json|javascript|problem\+json/i.test(contentType);
      if (isJson) {
        const body = await response.text().catch(() => "");
        if (body && body.length > 0) {
          item.sampleSize = body.length;
          try {
            const parsed = JSON.parse(body);
            item.payloadType = "json";
            item.parsedRecords = extractRecordsFromJsonPayload(parsed, url);
          } catch {
            item.payloadType = "text";
          }
        }
      }

      collected.push(item);
    } catch (error) {
      log("Error capturando response de red", "warn", { error: summarizeError(error) });
    }
  };

  page.on("response", onResponse);
  return {
    getAll: () => collected,
    stop: () => page.off("response", onResponse),
  };
}

function flattenObject(input, prefix = "", out = {}) {
  if (input === null || input === undefined) return out;
  if (typeof input !== "object") {
    if (prefix) out[prefix] = input;
    return out;
  }

  if (Array.isArray(input)) {
    if (input.length === 0 && prefix) out[prefix] = "";
    input.forEach((val, idx) => {
      const key = prefix ? `${prefix}[${idx}]` : `[${idx}]`;
      flattenObject(val, key, out);
    });
    return out;
  }

  for (const [k, v] of Object.entries(input)) {
    const nextKey = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === "object") {
      flattenObject(v, nextKey, out);
    } else {
      out[nextKey] = v;
    }
  }
  return out;
}

function extractRecordsFromJsonPayload(payload, sourceUrl) {
  const records = [];
  const seen = new Set();

  const scoreObject = (obj) => {
    if (!obj || typeof obj !== "object" || Array.isArray(obj)) return 0;
    const keys = Object.keys(obj).map((k) => k.toLowerCase());
    const hasName = keys.some((k) => /(name|empresa|company|razon|consignee|shipper|importador|exportador)/.test(k));
    const hasCountry = keys.some((k) => /(country|pais|origen|destino)/.test(k));
    const hasAddress = keys.some((k) => /(address|direccion|ciudad|provincia|city|state)/.test(k));
    const hasTax = keys.some((k) => /(cuit|tax|vat|id)/.test(k));
    let score = 0;
    if (hasName) score += 3;
    if (hasCountry) score += 2;
    if (hasAddress) score += 1;
    if (hasTax) score += 1;
    return score;
  };

  const visit = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      if (node.length > 0 && typeof node[0] === "object") {
        for (const item of node) visit(item);
      }
      return;
    }
    if (typeof node !== "object") return;

    const score = scoreObject(node);
    if (score >= 2) {
      const flat = flattenObject(node);
      const hash = crypto
        .createHash("sha1")
        .update(JSON.stringify(flat).slice(0, 5000))
        .digest("hex");
      if (!seen.has(hash)) {
        seen.add(hash);
        records.push({
          _source: "api_json",
          _sourceUrl: sourceUrl,
          ...flat,
        });
      }
    }

    for (const value of Object.values(node)) {
      if (typeof value === "object") visit(value);
    }
  };

  visit(payload);
  return records;
}

async function captureModuleDiagnostics(page, moduleDef, log) {
  const screenshotPath = path.join(DATA_DIR, `${moduleDef.debugBaseName}.png`);
  const htmlPath = path.join(DATA_DIR, `${moduleDef.debugBaseName}.html`);

  await page.screenshot({ path: screenshotPath, fullPage: true }).catch((error) => {
    log("No se pudo guardar screenshot de diagnóstico", "warn", {
      error: summarizeError(error),
    });
  });
  const html = await page.content().catch(() => "");
  if (html) {
    await fs.writeFile(htmlPath, html, "utf8");
  }

  const domStats = await page
    .evaluate(() => {
      const q = (sel) => document.querySelectorAll(sel).length;
      const visibleButtons = Array.from(
        document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], .btn')
      )
        .filter((el) => {
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return (
            style.visibility !== "hidden" &&
            style.display !== "none" &&
            rect.width > 0 &&
            rect.height > 0
          );
        })
        .map((el) => ((el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim()))
        .filter((t) => t.length > 0)
        .slice(0, 200);

      return {
        table: q("table"),
        tr: q("tr"),
        input: q("input"),
        select: q("select"),
        button: q("button"),
        roleRow: q('[role="row"]'),
        agRow: q(".ag-row"),
        pDatatableRows: q(".p-datatable-tbody tr"),
        visibleButtons,
      };
    })
    .catch(() => ({
      table: 0,
      tr: 0,
      input: 0,
      select: 0,
      button: 0,
      roleRow: 0,
      agRow: 0,
      pDatatableRows: 0,
      visibleButtons: [],
    }));

  log("Diagnóstico DOM", "info", domStats);
  return { screenshotPath, htmlPath, domStats };
}

async function clickSearchActions(page, log) {
  const searchRegex = /(buscar|consultar|filtrar|aplicar|search|actualizar|mostrar)/i;
  const clicked = [];

  for (let i = 0; i < 4; i += 1) {
    const btn = await findVisibleClickable(page, searchRegex);
    if (!btn) break;
    const label = await btn
      .evaluate((el) => (el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim())
      .catch(() => "boton");
    try {
      await Promise.all([
        page.waitForLoadState("networkidle", { timeout: 15000 }).catch(() => {}),
        btn.click({ timeout: 8000 }),
      ]);
      clicked.push(label);
      await waitForSettled(page);
      await sleep(600);
    } catch {
      break;
    }
  }

  if (clicked.length > 0) {
    log("Se ejecutaron acciones de búsqueda", "info", { clicked });
  } else {
    log("No se detectaron botones de búsqueda/aplicación");
  }
}

async function maximizeRowsPerPage(page, log) {
  const selects = page.locator("select");
  const count = await selects.count().catch(() => 0);
  let changed = 0;

  for (let i = 0; i < count; i += 1) {
    const s = selects.nth(i);
    const visible = await s.isVisible().catch(() => false);
    if (!visible) continue;

    const meta = await s
      .evaluate((el) => {
        const label =
          (el.getAttribute("aria-label") ||
            el.getAttribute("name") ||
            el.getAttribute("id") ||
            "") +
          " " +
          (el.closest("label")?.textContent || "");
        const options = Array.from(el.options || []).map((o) => ({
          value: o.value,
          text: (o.textContent || "").trim(),
        }));
        return { label: label.trim(), options };
      })
      .catch(() => ({ label: "", options: [] }));

    const looksLikePageSize = /(fila|rows|page size|por pagina|por p[aá]gina|cantidad|mostrar|registros|items)/i.test(
      meta.label
    );
    if (!looksLikePageSize) continue;

    const numericOptions = meta.options
      .map((opt) => ({
        ...opt,
        n: Number(String(opt.value || opt.text).replace(/[^\d]/g, "")),
      }))
      .filter((x) => Number.isFinite(x.n) && x.n > 0)
      .sort((a, b) => a.n - b.n);

    if (numericOptions.length === 0) continue;
    const maxOption = numericOptions[numericOptions.length - 1];

    try {
      await s.selectOption({ value: String(maxOption.value) }, { timeout: 8000 });
      changed += 1;
    } catch {
      await s.selectOption({ label: String(maxOption.text) }, { timeout: 8000 }).catch(() => {});
      changed += 1;
    }
  }

  if (changed > 0) {
    log("Selector de filas por página maximizado", "info", { changed });
    await waitForSettled(page);
  }
}

async function detectFilterCombos(page, log) {
  const filters = await page
    .evaluate(() => {
      const normalize = (v) => (v || "").replace(/\s+/g, " ").trim();
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const arr = [];
      const selects = Array.from(document.querySelectorAll("select")).filter(visible);
      selects.forEach((el, idx) => {
        const label = normalize(
          el.getAttribute("aria-label") ||
            el.getAttribute("name") ||
            el.getAttribute("id") ||
            el.closest("label")?.textContent ||
            ""
        );
        const options = Array.from(el.options || []).map((o) => ({
          value: o.value,
          text: normalize(o.textContent),
        }));
        arr.push({ index: idx, label, options });
      });
      return arr;
    })
    .catch(() => []);

  const eligible = filters
    .map((f) => {
      const options = (f.options || []).filter((o) => {
        const txt = (o.text || "").toLowerCase();
        return (
          (isTruthy(o.value) || isTruthy(o.text)) &&
          !/todos|all|seleccione|select|--|sin filtro/.test(txt)
        );
      });
      return { ...f, options };
    })
    .filter((f) => {
      if ((f.options || []).length <= 1) return false;
      return /(pais|country|period|fecha|anio|año|mes|origen|destino)/i.test(f.label);
    })
    .slice(0, 3);

  if (eligible.length === 0) {
    log("No se detectaron filtros de país/período útiles");
    return [{ selects: [] }];
  }

  let combos = [{ selects: [] }];
  for (const f of eligible) {
    const next = [];
    const limitedOptions = f.options.slice(0, 100);
    for (const combo of combos) {
      for (const opt of limitedOptions) {
        next.push({
          selects: [
            ...combo.selects,
            {
              index: f.index,
              label: f.label,
              value: opt.value || opt.text,
              text: opt.text || opt.value,
            },
          ],
        });
        if (next.length >= CONFIG.maxFilterCombos) break;
      }
      if (next.length >= CONFIG.maxFilterCombos) break;
    }
    combos = next.length ? next : combos;
    if (combos.length >= CONFIG.maxFilterCombos) break;
  }

  log("Filtros detectados", "info", {
    filters: eligible.map((x) => ({ label: x.label, options: x.options.length })),
    totalCombos: combos.length,
  });
  return combos.length ? combos : [{ selects: [] }];
}

async function applyFilterCombo(page, combo, log) {
  for (const sel of combo.selects || []) {
    const locator = page.locator("select").nth(sel.index);
    const visible = await locator.isVisible().catch(() => false);
    if (!visible) continue;
    try {
      await locator.selectOption({ value: String(sel.value) }, { timeout: 8000 });
    } catch {
      await locator.selectOption({ label: String(sel.text) }, { timeout: 8000 }).catch(() => {});
    }
  }
  await clickSearchActions(page, log);
  await waitForSettled(page);
}

function extractRelevantRequests(networkEvents) {
  const keywordRegex = /(import|export|detalle|search|data|grid|formulario)/i;
  return networkEvents
    .filter((e) => keywordRegex.test(e.url))
    .map((e) => ({
      ts: e.ts,
      method: e.method,
      status: e.status,
      url: e.url,
      contentType: e.contentType,
      parsedRecords: (e.parsedRecords || []).length,
    }));
}

async function detectNativeExport(page) {
  return findVisibleClickable(page, /(export|excel|csv|descargar|download|reporte)/i);
}

async function runNativeExport(page, moduleDef, comboTag, log) {
  const button = await detectNativeExport(page);
  if (!button) return null;
  const filePrefix = `${moduleDef.key}_native_${safeSlug(comboTag || "base")}`;

  const download = await Promise.all([
    page.waitForEvent("download", { timeout: 25000 }),
    button.click({ timeout: 8000 }),
  ])
    .then(([dl]) => dl)
    .catch(() => null);

  if (!download) return null;
  const ext = path.extname(download.suggestedFilename() || "") || ".dat";
  const outPath = path.join(DATA_DIR, `${filePrefix}${ext}`);
  await download.saveAs(outPath);
  log("Exportación nativa descargada", "info", { outPath });
  return outPath;
}

async function parseNativeFileToRows(filePath, moduleDef, log) {
  const ext = path.extname(filePath).toLowerCase();
  if (![".xlsx", ".csv"].includes(ext)) {
    log("Archivo nativo no parseable (se conserva)", "warn", { filePath });
    return [];
  }

  const workbook = new ExcelJS.Workbook();
  if (ext === ".xlsx") {
    await workbook.xlsx.readFile(filePath);
  } else {
    await workbook.csv.readFile(filePath);
  }
  const ws = workbook.worksheets[0];
  if (!ws) return [];

  const headers = ws
    .getRow(1)
    .values.slice(1)
    .map((h, i) => (isTruthy(h) ? String(h).trim() : `col_${i + 1}`));

  const rows = [];
  ws.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const vals = row.values.slice(1);
    if (!vals.some((v) => isTruthy(v) || typeof v === "number")) return;
    const obj = {};
    for (let i = 0; i < headers.length; i += 1) {
      obj[headers[i]] = vals[i] ?? "";
    }
    obj._source = "native_export";
    obj._module = moduleDef.key;
    rows.push(obj);
  });
  return rows;
}

async function extractDomRows(page, moduleDef, comboLabel, pageNo) {
  return page.evaluate(
    ({ moduleKey, companyType, comboLabelArg, pageNoArg }) => {
      const norm = (t) => (t || "").replace(/\s+/g, " ").trim();
      const visible = (el) => {
        const s = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return s.visibility !== "hidden" && s.display !== "none" && r.width > 0 && r.height > 0;
      };

      const rows = [];

      const pushRow = (raw, sourceKind) => {
        const output = {
          _source: sourceKind,
          _module: moduleKey,
          _companyType: companyType,
          _filter: comboLabelArg,
          _page: pageNoArg,
        };
        for (const [k, v] of Object.entries(raw)) {
          const key = norm(k) || "col";
          output[key] = norm(v);
        }
        rows.push(output);
      };

      // HTML tables
      const tables = Array.from(document.querySelectorAll("table")).filter(visible);
      tables.forEach((table, tIdx) => {
        const headers = Array.from(table.querySelectorAll("thead th")).map(
          (th, i) => norm(th.textContent) || `col_${i + 1}`
        );
        const trs = Array.from(table.querySelectorAll("tbody tr")).filter(visible);
        trs.forEach((tr) => {
          const tds = Array.from(tr.querySelectorAll("td"));
          if (!tds.length) return;
          const raw = { _table: tIdx + 1 };
          tds.forEach((td, idx) => {
            raw[headers[idx] || `col_${idx + 1}`] = norm(td.innerText || td.textContent);
          });
          const detailAnchor = tr.querySelector("a[href]");
          if (detailAnchor) raw._detailHref = detailAnchor.getAttribute("href") || "";
          pushRow(raw, "dom_table");
        });
      });

      // role=row grids
      const roleRows = Array.from(document.querySelectorAll('[role="row"]')).filter(visible);
      if (roleRows.length > 1) {
        const headerCells = Array.from(
          roleRows[0].querySelectorAll('[role="columnheader"], .ag-header-cell-text')
        );
        const headers = headerCells.map((h, i) => norm(h.textContent) || `col_${i + 1}`);
        for (let i = 1; i < roleRows.length; i += 1) {
          const rr = roleRows[i];
          const cells = Array.from(
            rr.querySelectorAll('[role="gridcell"], [role="cell"], .ag-cell, .p-datatable-tbody td')
          );
          if (!cells.length) continue;
          const raw = {};
          cells.forEach((cell, idx) => {
            raw[headers[idx] || `col_${idx + 1}`] = norm(cell.innerText || cell.textContent);
          });
          const detailAnchor = rr.querySelector("a[href]");
          if (detailAnchor) raw._detailHref = detailAnchor.getAttribute("href") || "";
          pushRow(raw, "dom_role_grid");
        }
      }

      // ag-grid fallback by row/cell classes
      const agRows = Array.from(document.querySelectorAll(".ag-row")).filter(visible);
      if (agRows.length) {
        agRows.forEach((r) => {
          const cells = Array.from(r.querySelectorAll(".ag-cell"));
          if (!cells.length) return;
          const raw = {};
          cells.forEach((c, idx) => {
            raw[`ag_col_${idx + 1}`] = norm(c.innerText || c.textContent);
          });
          pushRow(raw, "dom_ag_grid");
        });
      }

      // PrimeNG fallback
      const pRows = Array.from(document.querySelectorAll(".p-datatable-tbody tr")).filter(visible);
      if (pRows.length) {
        pRows.forEach((tr) => {
          const tds = Array.from(tr.querySelectorAll("td"));
          if (!tds.length) return;
          const raw = {};
          tds.forEach((td, idx) => {
            raw[`p_col_${idx + 1}`] = norm(td.innerText || td.textContent);
          });
          pushRow(raw, "dom_primeng");
        });
      }

      // Generic div row-like fallback
      const divRows = Array.from(
        document.querySelectorAll(
          '.row, .grid-row, .table-row, [class*="row"], [data-row-index], [data-testid*="row"]'
        )
      )
        .filter(visible)
        .slice(0, 2000);
      if (divRows.length && rows.length === 0) {
        divRows.forEach((r) => {
          const txt = norm(r.innerText || r.textContent);
          if (!txt || txt.length < 3) return;
          pushRow({ rawText: txt }, "dom_div_grid");
        });
      }

      return rows;
    },
    {
      moduleKey: moduleDef.key,
      companyType: moduleDef.companyType,
      comboLabelArg: comboLabel,
      pageNoArg: pageNo,
    }
  );
}

async function openDetailIfAnyAndExtract(page, context, row, log) {
  const href = row._detailHref || row._detail_href || row.detailHref;
  if (!isTruthy(href)) return row;

  let absolute = String(href);
  if (!absolute.startsWith("http")) {
    absolute = `${CONFIG.baseUrl}${absolute.startsWith("/") ? "" : "/"}${absolute}`;
  }
  if (!absolute.startsWith(CONFIG.baseUrl)) return row;

  const p = await context.newPage();
  p.setDefaultTimeout(CONFIG.timeoutMs);
  try {
    await p.goto(absolute, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await waitForSettled(p);
    const detail = await p.evaluate(() => {
      const norm = (t) => (t || "").replace(/\s+/g, " ").trim();
      const out = {};
      document.querySelectorAll("dt").forEach((dt) => {
        const dd = dt.nextElementSibling;
        if (dd && dd.tagName.toLowerCase() === "dd") {
          const k = norm(dt.textContent);
          if (k) out[`detail_${k}`] = norm(dd.textContent);
        }
      });
      document.querySelectorAll("label").forEach((lb) => {
        const k = norm(lb.textContent);
        if (!k || out[`detail_${k}`]) return;
        const sib = lb.nextElementSibling;
        if (sib) out[`detail_${k}`] = norm(sib.textContent);
      });
      return out;
    });
    await p.close();
    return { ...row, ...detail, _detailUrl: absolute };
  } catch (error) {
    await p.close();
    log("No se pudo extraer detalle de registro", "warn", {
      href: absolute,
      error: summarizeError(error),
    });
    return row;
  }
}

async function advancePagination(page, log) {
  const signatureBefore = await page
    .evaluate(() => {
      const snap = Array.from(document.querySelectorAll("table tbody tr, .ag-row, .p-datatable-tbody tr"))
        .slice(0, 5)
        .map((n) => (n.innerText || "").replace(/\s+/g, " ").trim())
        .join("||");
      return `${window.location.href}::${snap}`;
    })
    .catch(() => "");

  const loadMore = await findVisibleClickable(page, /(cargar m[aá]s|mostrar m[aá]s|load more|ver m[aá]s)/i);
  if (loadMore) {
    await loadMore.click({ timeout: 8000 }).catch(() => {});
    await waitForSettled(page);
    return true;
  }

  const next = await findVisibleClickable(page, /(siguiente|next|pr[oó]xima|›|»)/i);
  if (next) {
    await next.click({ timeout: 8000 }).catch(() => {});
    await waitForSettled(page);
    const signatureAfter = await page
      .evaluate(() => {
        const snap = Array.from(document.querySelectorAll("table tbody tr, .ag-row, .p-datatable-tbody tr"))
          .slice(0, 5)
          .map((n) => (n.innerText || "").replace(/\s+/g, " ").trim())
          .join("||");
        return `${window.location.href}::${snap}`;
      })
      .catch(() => "");
    if (signatureAfter !== signatureBefore) return true;
  }

  const beforeCount = await page
    .locator("table tbody tr, .ag-row, .p-datatable-tbody tr, [role='row']")
    .count()
    .catch(() => 0);
  await page.mouse.wheel(0, 2200).catch(() => {});
  await sleep(600);
  await waitForSettled(page);
  const afterCount = await page
    .locator("table tbody tr, .ag-row, .p-datatable-tbody tr, [role='row']")
    .count()
    .catch(() => beforeCount);
  if (afterCount > beforeCount) {
    log("Infinite scroll detectado", "info", { beforeCount, afterCount });
    return true;
  }

  return false;
}

function getValueByKeyPatterns(raw, patterns) {
  const entries = Object.entries(raw || {});
  for (const [key, value] of entries) {
    const lk = key.toLowerCase();
    if (patterns.some((p) => p.test(lk)) && isTruthy(value)) {
      return String(value).trim();
    }
  }
  return "";
}

function normalizeRecord(raw, moduleDef, sourceUrl) {
  const name = getValueByKeyPatterns(raw, [
    /^(name|nombre|empresa|razon|raz[oó]n|consignee|shipper|importador|exportador)$/,
    /(name|nombre|empresa|razon|consignee|shipper|importador|exportador)/,
  ]);
  const country = getValueByKeyPatterns(raw, [/(country|pais|pa[ií]s|origen|destino)/]);
  const city = getValueByKeyPatterns(raw, [/(city|ciudad|localidad)/]);
  const province = getValueByKeyPatterns(raw, [/(province|provincia|state|estado)/]);
  const address = getValueByKeyPatterns(raw, [/(address|direccion|domicilio|street|calle)/]);
  const taxId = getValueByKeyPatterns(raw, [/(cuit|tax|vat|fiscal|taxid|tax_id|id tributaria)/]);
  const contact = getValueByKeyPatterns(raw, [/(contact|contacto|responsable|attn|persona)/]);
  const phone = getValueByKeyPatterns(raw, [/(phone|telefono|tel|celular|mobile)/]);
  const email = getValueByKeyPatterns(raw, [/(email|correo|mail|e-mail)/]);

  if (!name && !taxId) return null;

  return {
    name: name || "",
    companyType: moduleDef.companyType,
    country: country || "",
    city: city || "",
    province: province || "",
    address: address || "",
    taxId: taxId || "",
    contact: contact || "",
    phone: phone || "",
    email: email || "",
    sourceModule: moduleDef.key,
    sourceUrl,
    extractedAt: nowIso(),
    _raw: raw,
  };
}

function dedupeNormalized(records) {
  const out = new Map();
  for (const r of records) {
    if (!r) continue;
    const base = `${(r.name || "").toLowerCase().trim()}|${(r.country || "").toLowerCase().trim()}`;
    const key = r.taxId ? `${base}|tax:${String(r.taxId).toLowerCase().trim()}` : base;
    if (!out.has(key)) {
      out.set(key, r);
    } else {
      const prev = out.get(key);
      const merged = { ...prev };
      for (const [k, v] of Object.entries(r)) {
        if (!isTruthy(merged[k]) && isTruthy(v)) merged[k] = v;
      }
      out.set(key, merged);
    }
  }
  return [...out.values()];
}

async function writeMasterOutputs(records, metadata, logs) {
  const jsonPath = path.join(DATA_DIR, "consignees_shippers_master.json");
  const xlsxPath = path.join(DATA_DIR, "consignees_shippers_master.xlsx");
  const csvPath = path.join(DATA_DIR, "consignees_shippers_master.csv");

  await fs.writeFile(jsonPath, JSON.stringify({ metadata, records }, null, 2), "utf8");

  const workbook = new ExcelJS.Workbook();
  const dataSheet = workbook.addWorksheet("data");
  const metadataSheet = workbook.addWorksheet("metadata");
  const logsSheet = workbook.addWorksheet("logs");

  const columns = [
    "name",
    "companyType",
    "country",
    "city",
    "province",
    "address",
    "taxId",
    "contact",
    "phone",
    "email",
    "sourceModule",
    "sourceUrl",
    "extractedAt",
  ];

  dataSheet.columns = columns.map((c) => ({ header: c, key: c, width: 24 }));
  records.forEach((r) => dataSheet.addRow(r));
  dataSheet.views = [{ state: "frozen", ySplit: 1 }];
  dataSheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };
  dataSheet.columns.forEach((col) => {
    let max = String(col.header).length;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const len = String(cell.value ?? "").length;
      if (len > max) max = len;
    });
    col.width = Math.min(Math.max(max + 2, 14), 60);
  });

  metadataSheet.columns = [
    { header: "campo", key: "campo", width: 40 },
    { header: "valor", key: "valor", width: 120 },
  ];
  Object.entries(metadata).forEach(([campo, valor]) => {
    metadataSheet.addRow({
      campo,
      valor: typeof valor === "string" ? valor : JSON.stringify(valor),
    });
  });

  logsSheet.columns = [
    { header: "timestamp", key: "ts", width: 28 },
    { header: "level", key: "level", width: 10 },
    { header: "scope", key: "scope", width: 26 },
    { header: "message", key: "message", width: 80 },
    { header: "extra", key: "extra", width: 120 },
  ];
  logs.forEach((l) => logsSheet.addRow(l));

  await workbook.xlsx.writeFile(xlsxPath);

  const csvWorkbook = new ExcelJS.Workbook();
  const csvSheet = csvWorkbook.addWorksheet("master");
  csvSheet.columns = columns.map((c) => ({ header: c, key: c }));
  records.forEach((r) => csvSheet.addRow(r));
  await csvWorkbook.csv.writeFile(csvPath);

  return { jsonPath, xlsxPath, csvPath };
}

async function processModule(page, context, moduleDef, globalLogs) {
  const log = makeLogger(moduleDef.key, globalLogs);
  const moduleUrl = `${CONFIG.baseUrl}${moduleDef.urlPath}`;
  const moduleRawRows = [];
  const moduleNormalized = [];
  const usedMechanisms = new Set();
  const moduleErrors = [];

  log("Navegando módulo", "info", { moduleUrl });
  await page.goto(moduleUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
  await maybeLogin(page, log);
  await waitForSettled(page);

  const networkCollector = createNetworkCollector(page, log);
  await clickSearchActions(page, log);
  await maximizeRowsPerPage(page, log);
  await clickSearchActions(page, log);
  await waitForSettled(page);

  const diagnostics = await captureModuleDiagnostics(page, moduleDef, log);

  const filterCombos = await detectFilterCombos(page, log);
  let nativeDownloadCount = 0;

  for (let comboIdx = 0; comboIdx < filterCombos.length; comboIdx += 1) {
    const combo = filterCombos[comboIdx];
    const comboLabel =
      (combo.selects || [])
        .map((s) => `${s.label || "filter"}=${s.text || s.value}`)
        .join(" | ") || "sin_filtro";

    log("Aplicando filtro", "info", { comboIdx: comboIdx + 1, comboLabel });
    await applyFilterCombo(page, combo, log);

    // Intento exportación nativa por combinación.
    const nativePath = await retry(
      async () => runNativeExport(page, moduleDef, `${comboIdx + 1}_${comboLabel}`, log),
      {
        retries: 2,
        delayMs: 1200,
        actionName: "runNativeExport",
        onRetry: (attempt, err) =>
          log("Reintento exportación nativa", "warn", {
            attempt,
            error: summarizeError(err),
          }),
      }
    ).catch(() => null);

    if (nativePath) {
      usedMechanisms.add("native_export");
      nativeDownloadCount += 1;
      const parsedRows = await parseNativeFileToRows(nativePath, moduleDef, log).catch((error) => {
        moduleErrors.push(`parse_native: ${summarizeError(error)}`);
        return [];
      });
      moduleRawRows.push(...parsedRows);
    }

    // Si no hubo datos de export nativa, o vino vacía, recorrer DOM + paginación.
    if (!nativePath || moduleRawRows.length === 0) {
      let pageNo = 1;
      let pagesVisited = 0;
      while (pagesVisited < CONFIG.maxPagesPerModule) {
        pagesVisited += 1;
        log("Extrayendo grilla DOM", "info", { comboLabel, pageNo });
        const rows = await extractDomRows(page, moduleDef, comboLabel, pageNo).catch((error) => {
          moduleErrors.push(`extract_dom: ${summarizeError(error)}`);
          return [];
        });

        for (const row of rows) {
          const withDetail = await openDetailIfAnyAndExtract(page, context, row, log);
          moduleRawRows.push(withDetail);
        }

        const moved = await advancePagination(page, log);
        if (!moved) break;
        pageNo += 1;
      }
      if (moduleRawRows.length > 0) usedMechanisms.add("dom_grid");
    }
  }

  // Si sigue vacío, última chance: usar eventos de red parseados.
  const networkEvents = networkCollector.getAll();
  networkCollector.stop();
  const apiRows = networkEvents.flatMap((e) => e.parsedRecords || []);
  if (apiRows.length > 0) {
    usedMechanisms.add("api_json");
    moduleRawRows.push(...apiRows.map((r) => ({ ...r, _module: moduleDef.key })));
  }

  // Normalización
  for (const raw of moduleRawRows) {
    const normalized = normalizeRecord(raw, moduleDef, moduleUrl);
    if (normalized) moduleNormalized.push(normalized);
  }

  log("Resumen módulo", "info", {
    rawRows: moduleRawRows.length,
    normalizedRows: moduleNormalized.length,
    nativeDownloadCount,
    mechanisms: [...usedMechanisms],
  });

  if (moduleNormalized.length === 0) {
    moduleErrors.push("Sin registros normalizados tras agotar export nativa + API + DOM");
    log("Módulo sin datos tras agotar estrategias", "warn");
  }

  return {
    module: moduleDef.key,
    moduleUrl,
    rawRows: moduleRawRows,
    normalizedRows: moduleNormalized,
    diagnostics,
    usedMechanisms: [...usedMechanisms],
    relevantRequests: extractRelevantRequests(networkEvents),
    errors: moduleErrors,
  };
}

async function main() {
  await ensureDirectories();
  const allLogs = [];
  const log = makeLogger("main", allLogs);

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1600, height: 1000 },
  });
  const page = await context.newPage();
  page.setDefaultTimeout(CONFIG.timeoutMs);

  const moduleResults = [];
  try {
    const initialUrl = `${CONFIG.baseUrl}${MODULES[0].urlPath}`;
    log("Inicio de ejecución", "info", { baseUrl: CONFIG.baseUrl, initialUrl });
    await page.goto(initialUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await maybeLogin(page, makeLogger("auth", allLogs));
    await waitForSettled(page);

    for (const moduleDef of MODULES) {
      const res = await processModule(page, context, moduleDef, allLogs);
      moduleResults.push(res);
    }
  } catch (error) {
    log("Fallo fatal durante ejecución", "error", { error: summarizeError(error) });
    await page.screenshot({
      path: path.join(DATA_DIR, `fatal_${safeSlug(nowIso())}.png`),
      fullPage: true,
    }).catch(() => {});
  } finally {
    await browser.close();
  }

  const allNormalized = moduleResults.flatMap((m) => m.normalizedRows || []);
  const deduped = dedupeNormalized(allNormalized);

  const metadata = {
    extractedAt: nowIso(),
    baseUrl: CONFIG.baseUrl,
    totalModules: MODULES.length,
    totalRawRows: moduleResults.reduce((acc, m) => acc + (m.rawRows?.length || 0), 0),
    totalNormalizedRows: allNormalized.length,
    totalDedupedRows: deduped.length,
    mechanismsByModule: moduleResults.map((m) => ({
      module: m.module,
      mechanisms: m.usedMechanisms,
    })),
    moduleErrors: moduleResults.map((m) => ({ module: m.module, errors: m.errors })),
    diagnostics: moduleResults.map((m) => ({
      module: m.module,
      screenshot: m.diagnostics?.screenshotPath || "",
      html: m.diagnostics?.htmlPath || "",
      domStats: m.diagnostics?.domStats || {},
      relevantRequests: m.relevantRequests || [],
    })),
  };

  const outputPaths = await writeMasterOutputs(deduped, metadata, allLogs);

  log("Proceso finalizado", "info", {
    totalDedupedRows: deduped.length,
    outputs: outputPaths,
  });

  if (deduped.length === 0) {
    console.log(
      `[${nowIso()}] [WARN] El resultado final quedó en 0 filas. Revisar archivos debug y logs en /data`
    );
  }
}

main().catch((error) => {
  console.error(`[fatal] ${summarizeError(error)}`);
  process.exit(1);
});

