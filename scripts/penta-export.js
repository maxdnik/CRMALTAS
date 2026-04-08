#!/usr/bin/env node
"use strict";

const fs = require("node:fs/promises");
const fssync = require("node:fs");
const path = require("node:path");
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
};

const ROOT_DIR = path.resolve(__dirname, "..");
const DATA_DIR = path.join(ROOT_DIR, "data");
const AUTH_STATE_PATH = path.join(DATA_DIR, "auth-state.json");
const LOGIN_FAILED_SCREENSHOT = path.join(DATA_DIR, "login-failed.png");
const LOGIN_FAILED_HTML = path.join(DATA_DIR, "login-failed.html");
const LOGIN_SUCCESS_SCREENSHOT = path.join(DATA_DIR, "login-success.png");

const TARGET_MODULE_URL = `${CONFIG.baseUrl}/home/formulario/AR/importDetalladas`;

const OUTPUTS = {
  paisesOrigen: path.join(DATA_DIR, "paises-origen.json"),
  paisesSinResultados: path.join(DATA_DIR, "paises-sin-resultados.json"),
  paisesConError: path.join(DATA_DIR, "paises-con-error.json"),
  porPaisJson: path.join(DATA_DIR, "importadores_por_pais.json"),
  porPaisCsv: path.join(DATA_DIR, "importadores_por_pais.csv"),
  porPaisXlsx: path.join(DATA_DIR, "importadores_por_pais.xlsx"),
  unicosJson: path.join(DATA_DIR, "importadores_unicos.json"),
  unicosCsv: path.join(DATA_DIR, "importadores_unicos.csv"),
  unicosXlsx: path.join(DATA_DIR, "importadores_unicos.xlsx"),
};

const LOGIN_TEXT_REGEX =
  /(user login|user password|forgot my password|enter|iniciar sesi[oó]n|ingresar|password)/i;

const NO_RESULTS_REGEX = /no se encontraron resultados|no results|sin datos|no data|no records/i;

function nowIso() {
  return new Date().toISOString();
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function summarizeError(error) {
  if (!error) return "unknown_error";
  if (error instanceof Error) return `${error.name}: ${error.message}`;
  return String(error);
}

function makeLogger(scope) {
  return (message, level = "info", extra = undefined) => {
    const extraTxt = extra ? ` | ${JSON.stringify(extra)}` : "";
    console.log(`[${nowIso()}] [${level.toUpperCase()}] [${scope}] ${message}${extraTxt}`);
  };
}

async function ensureDirectories() {
  if (!fssync.existsSync(DATA_DIR)) {
    await fs.mkdir(DATA_DIR, { recursive: true });
  }
}

async function waitForSettled(page) {
  await page.waitForLoadState("domcontentloaded", { timeout: 25000 }).catch(() => {});
  await page.waitForLoadState("networkidle", { timeout: 12000 }).catch(() => {});
}

async function saveJson(filePath, data) {
  await fs.writeFile(filePath, JSON.stringify(data, null, 2), "utf8");
}

async function writeCsv(filePath, headers, rows) {
  const escape = (v) => {
    const txt = String(v ?? "");
    if (txt.includes('"') || txt.includes(",") || txt.includes("\n")) {
      return `"${txt.replace(/"/g, '""')}"`;
    }
    return txt;
  };
  const lines = [headers.join(",")];
  for (const row of rows) {
    lines.push(headers.map((h) => escape(row[h])).join(","));
  }
  await fs.writeFile(filePath, `${lines.join("\n")}\n`, "utf8");
}

async function writeXlsx(filePath, sheetName, headers, rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  ws.columns = headers.map((h) => ({ header: h, key: h, width: Math.max(16, h.length + 4) }));
  rows.forEach((r) => ws.addRow(r));
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: headers.length || 1 },
  };
  await wb.xlsx.writeFile(filePath);
}

async function createContextWithAuthState(browser, log) {
  const options = {
    acceptDownloads: true,
    viewport: { width: 1600, height: 1000 },
  };
  if (fssync.existsSync(AUTH_STATE_PATH)) {
    try {
      log("Intentando reutilizar auth-state", "info", { path: AUTH_STATE_PATH });
      return await browser.newContext({ ...options, storageState: AUTH_STATE_PATH });
    } catch (error) {
      log("No se pudo cargar auth-state; se crea contexto limpio", "warn", {
        error: summarizeError(error),
      });
    }
  }
  return browser.newContext(options);
}

async function isLoginLikeState(page) {
  return page
    .evaluate((regexSource) => {
      const loginRegex = new RegExp(regexSource, "i");
      const url = window.location.href;
      const txt = normalizeText(document.body?.innerText || "").slice(0, 3000);
      const isLoginUrl = /\/login(?:\/|$|\?)/i.test(url.toLowerCase());
      const hasLoginText = loginRegex.test(txt);
      return {
        url,
        isLoginUrl,
        hasLoginText,
        visibleTextPreview: txt.slice(0, 500),
      };
    }, LOGIN_TEXT_REGEX.source)
    .catch(() => ({
      url: page.url(),
      isLoginUrl: /\/login(?:\/|$|\?)/i.test(page.url().toLowerCase()),
      hasLoginText: false,
      visibleTextPreview: "",
    }));
}

async function saveLoginFailureArtifacts(page, log) {
  await page
    .screenshot({ path: LOGIN_FAILED_SCREENSHOT, fullPage: true })
    .then(() => log("Screenshot de login fallido guardado", "warn", { path: LOGIN_FAILED_SCREENSHOT }))
    .catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) {
    await fs.writeFile(LOGIN_FAILED_HTML, html, "utf8");
    log("HTML de login fallido guardado", "warn", { path: LOGIN_FAILED_HTML });
  }
}

async function saveLoginSuccessScreenshot(page, log) {
  await page
    .screenshot({ path: LOGIN_SUCCESS_SCREENSHOT, fullPage: true })
    .then(() => log("Screenshot de login exitoso guardado", "info", { path: LOGIN_SUCCESS_SCREENSHOT }))
    .catch(() => {});
}

async function getFirstVisibleLocator(page, selectors) {
  for (const selector of selectors) {
    const loc = page.locator(selector).first();
    if (await loc.isVisible().catch(() => false)) {
      return { locator: loc, selector };
    }
  }
  return null;
}

async function ensureLoggedIn(page, context, log) {
  log("login iniciado");
  await page.goto(`${CONFIG.baseUrl}/login`, {
    waitUntil: "domcontentloaded",
    timeout: CONFIG.timeoutMs,
  });
  await waitForSettled(page);

  const preState = await isLoginLikeState(page);
  if (!preState.isLoginUrl && !preState.hasLoginText) {
    await context.storageState({ path: AUTH_STATE_PATH });
    log("Sesión ya activa, auth-state actualizado", "info", { url: preState.url });
    return;
  }

  const userCandidate = await getFirstVisibleLocator(page, [
    'input[name*="user" i]',
    'input[id*="user" i]',
    'input[name*="mail" i]',
    'input[type="email"]',
    'input[placeholder*="usuario" i]',
    'input[placeholder*="user" i]',
    'input[autocomplete="username"]',
    'input[type="text"]',
  ]);
  const passCandidate = await getFirstVisibleLocator(page, [
    'input[name*="pass" i]',
    'input[id*="pass" i]',
    'input[placeholder*="clave" i]',
    'input[placeholder*="password" i]',
    'input[autocomplete="current-password"]',
    'input[type="password"]',
  ]);

  if (!userCandidate || !passCandidate) {
    await saveLoginFailureArtifacts(page, log);
    throw new Error("No se detectaron campos de login");
  }

  log("campos detectados", "info", {
    userSelector: userCandidate.selector,
    passSelector: passCandidate.selector,
  });

  await userCandidate.locator.fill(CONFIG.user, { timeout: 12000 });
  await passCandidate.locator.fill(CONFIG.pass, { timeout: 12000 });

  const submitCandidates = await page
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const norm = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const form =
        document.querySelector('input[type="password"]')?.closest("form") ||
        document.querySelector("form");
      if (!form) return [];
      const nodes = Array.from(
        form.querySelectorAll(
          "button, input, ion-button, a, [role='button'], span, div, .button, .botonLogin"
        )
      ).filter(visible);

      return nodes.map((node, index) => {
        const parent = node.parentElement;
        return {
          index,
          tagName: (node.tagName || "").toLowerCase(),
          text: norm(node.innerText || node.textContent || node.value || ""),
          type: norm(node.getAttribute("type") || ""),
          href: norm(node.getAttribute("href") || ""),
          parentTag: (parent?.tagName || "").toLowerCase(),
          parentType: norm(parent?.getAttribute?.("type") || ""),
          selectorHint: `${(node.tagName || "").toLowerCase()}${node.getAttribute("type") ? `[type="${node.getAttribute("type")}"]` : ""}`,
        };
      });
    })
    .catch(() => []);

  log("candidatos de submit detectados", "info", {
    total: submitCandidates.length,
    candidates: submitCandidates.map((c) => ({
      tagName: c.tagName,
      text: c.text,
      type: c.type,
      href: c.href,
    })),
  });

  const isDisallowed = (c) => {
    const txt = String(c.text || "").toLowerCase();
    const href = String(c.href || "").toLowerCase();
    return c.tagName === "a" || txt.includes("forgot") || txt.includes("password") || href.includes("olvide");
  };
  const exact = (v) => (txt) => String(txt || "").trim().toLowerCase() === v;
  const allowed = submitCandidates.filter((c) => !isDisallowed(c));

  const selected =
    allowed.find((c) => c.tagName === "button" && c.type === "submit") ||
    allowed.find((c) => c.tagName === "input" && c.type === "submit") ||
    allowed.find((c) => c.tagName === "button" && exact("enter")(c.text)) ||
    allowed.find((c) => c.tagName === "button" && exact("ingresar")(c.text)) ||
    allowed.find((c) => c.tagName === "ion-button" && c.type === "submit") ||
    allowed.find((c) => c.tagName === "ion-button" && exact("enter")(c.text)) ||
    allowed.find((c) => c.tagName === "ion-button" && exact("ingresar")(c.text)) ||
    allowed.find(
      (c) =>
        c.tagName === "span" &&
        exact("enter")(c.text) &&
        ((c.parentTag === "ion-button" && c.parentType === "submit") ||
          (c.parentTag === "button" && c.parentType === "submit"))
    ) ||
    allowed.find(
      (c) =>
        c.tagName === "span" &&
        exact("ingresar")(c.text) &&
        ((c.parentTag === "ion-button" && c.parentType === "submit") ||
          (c.parentTag === "button" && c.parentType === "submit"))
    ) ||
    null;

  log("submit elegido", "info", {
    selector: selected?.selectorHint || "form.requestSubmit()/form.submit()",
    tagName: selected?.tagName || "form",
    text: selected?.text || "",
    type: selected?.type || "",
    href: selected?.href || "",
  });

  if (selected) {
    await page
      .evaluate((candidate) => {
        const visible = (el) => {
          const st = window.getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
        };
        const form =
          document.querySelector('input[type="password"]')?.closest("form") ||
          document.querySelector("form");
        if (!form) return false;
        const nodes = Array.from(
          form.querySelectorAll(
            "button, input, ion-button, a, [role='button'], span, div, .button, .botonLogin"
          )
        ).filter(visible);
        const node = nodes[candidate.index];
        if (!node) return false;
        node.click();
        return true;
      }, selected)
      .catch(async () => {
        await page
          .locator("form")
          .first()
          .evaluate((form) => {
            if (form.requestSubmit) form.requestSubmit();
            else form.submit();
          })
          .catch(() => {});
      });
  } else {
    await page
      .locator("form")
      .first()
      .evaluate((form) => {
        if (form.requestSubmit) form.requestSubmit();
        else form.submit();
      })
      .catch(() => {});
  }

  await page
    .waitForURL((url) => url.toString().toLowerCase().includes("/home/"), { timeout: 10000 })
    .catch(() => null);
  await Promise.race([
    page.waitForURL((url) => !url.toString().toLowerCase().includes("/login"), { timeout: 10000 }).catch(() => null),
    page.locator('input[type="password"]').first().waitFor({ state: "hidden", timeout: 10000 }).catch(() => null),
    sleep(3500),
  ]);
  await waitForSettled(page);

  const postState = await isLoginLikeState(page);
  const urlLower = String(postState.url || "").toLowerCase();
  log("url post-login", "info", {
    currentUrl: postState.url,
    isLoginUrl: postState.isLoginUrl,
    hasLoginText: postState.hasLoginText,
  });

  if (urlLower.includes("/olvide-mi-password") || urlLower.includes("olvide")) {
    log("se hizo click en recuperación de contraseña por error", "error", { currentUrl: postState.url });
    await saveLoginFailureArtifacts(page, log);
    throw new Error("Login fallido: se navegó a recuperación de contraseña");
  }

  const loginSuccess = !urlLower.includes("/login") && urlLower.includes("/home");
  if (!loginSuccess) {
    await saveLoginFailureArtifacts(page, log);
    throw new Error("Login fallido: sigue en /login");
  }

  log("login exitoso por navegación a /home/dashboard", "info", { url: postState.url });
  await saveLoginSuccessScreenshot(page, log);
  await context.storageState({ path: AUTH_STATE_PATH });
}

async function navigateAndEnsureSession(page, context, log, url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
  await waitForSettled(page);
  const state = await isLoginLikeState(page);
  if (state.isLoginUrl) {
    log("Redirección a /login detectada. Reautenticando", "warn", { currentUrl: state.url });
    await ensureLoggedIn(page, context, log);
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await waitForSettled(page);
  }
}

async function getVisibleTextSnapshot(page) {
  return page
    .evaluate(() => normalizeText(document.body?.innerText || "").slice(0, 4000))
    .catch(() => "");
}

async function openOriginCountryControl(page, log) {
  const exactLabel = page.getByText("País de Origen", { exact: true });
  if (await exactLabel.first().isVisible().catch(() => false)) {
    await exactLabel.first().click({ timeout: 7000 }).catch(() => {});
    await sleep(300);
  }

  const selectLike = page.locator(
    [
      '[aria-label*="País de Origen" i]',
      '[placeholder*="País de Origen" i]',
      'label:has-text("País de Origen") + *',
      '[id*="pais" i]',
      '[name*="pais" i]',
      "ion-select",
      "p-dropdown",
      "ng-select",
      '[role="combobox"]',
      "select",
    ].join(",")
  );

  const count = await selectLike.count().catch(() => 0);
  for (let i = 0; i < count; i += 1) {
    const node = selectLike.nth(i);
    if (!(await node.isVisible().catch(() => false))) continue;
    await node.click({ timeout: 7000 }).catch(() => {});
    await sleep(350);
    log("Combo País de Origen abierto", "info");
    return true;
  }

  return false;
}

async function readCountryOptions(page) {
  return page
    .evaluate(() => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };

      const candidates = Array.from(
        document.querySelectorAll(
          [
            '[role="option"]',
            "option",
            ".p-dropdown-item",
            ".ng-option",
            "ion-select-option",
            "li",
            ".mat-option",
            ".cdk-option",
          ].join(",")
        )
      ).filter(visible);

      const out = [];
      for (const el of candidates) {
        const txt = normalize(el.innerText || el.textContent || el.getAttribute("label") || "");
        if (!txt) continue;
        const lower = txt.toLowerCase();
        if (/(seleccione|select|todos|all|--|elija)/.test(lower)) continue;
        out.push(txt);
      }
      return out;
    })
    .catch(() => []);
}

async function closeCountryDropdown(page) {
  await page.keyboard.press("Escape").catch(() => {});
  await sleep(200);
}

async function getCountriesFromOriginField(page, log) {
  const opened = await openOriginCountryControl(page, log);
  if (!opened) {
    throw new Error('No se pudo abrir el combo "País de Origen"');
  }
  const options = await readCountryOptions(page);
  await closeCountryDropdown(page);
  const dedup = [...new Set(options.map((x) => normalizeText(x)).filter(Boolean))];
  return dedup;
}

async function selectCountryExact(page, country, log) {
  await closeCountryDropdown(page);
  const opened = await openOriginCountryControl(page, log);
  if (!opened) throw new Error("No se pudo abrir combo para seleccionar país");

  // Intentar limpiar filtros previos de texto.
  const filterInputs = page.locator(
    'input[type="search"], input[placeholder*="buscar" i], input[placeholder*="search" i], input[type="text"]'
  );
  const fiCount = await filterInputs.count().catch(() => 0);
  for (let i = 0; i < Math.min(fiCount, 3); i += 1) {
    const input = filterInputs.nth(i);
    if (!(await input.isVisible().catch(() => false))) continue;
    await input.fill("").catch(() => {});
    await input.fill(country).catch(() => {});
    await sleep(250);
    break;
  }

  const optionClicked = await page
    .evaluate((countryArg) => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const candidates = Array.from(
        document.querySelectorAll(
          [
            '[role="option"]',
            "option",
            ".p-dropdown-item",
            ".ng-option",
            "ion-select-option",
            "li",
            ".mat-option",
            ".cdk-option",
          ].join(",")
        )
      ).filter(visible);

      const wanted = normalize(countryArg).toLowerCase();
      const match = candidates.find((el) => normalize(el.innerText || el.textContent || "").toLowerCase() === wanted);
      if (!match) return false;
      match.click();
      return true;
    }, country)
    .catch(() => false);

  if (!optionClicked) {
    await closeCountryDropdown(page);
    throw new Error(`No se pudo seleccionar país exacto: ${country}`);
  }

  await sleep(350);
  await closeCountryDropdown(page);
  log("país seleccionado", "info", { pais: country });
}

async function clickBuscar(page, log) {
  const clicked = await page
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const candidates = Array.from(
        document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], ion-button')
      ).filter(visible);
      const target = candidates.find((el) => normalize(el.innerText || el.textContent || el.value || "") === "buscar");
      if (!target) return false;
      target.click();
      return true;
    })
    .catch(() => false);

  if (!clicked) {
    const fallback = page.getByText("Buscar", { exact: true }).first();
    if (await fallback.isVisible().catch(() => false)) {
      await fallback.click({ timeout: 8000 }).catch(() => {});
    } else {
      throw new Error('No se encontró botón "Buscar"');
    }
  }
  log("buscar ejecutado");
  await waitForSettled(page);
}

async function detectNoResultsState(page) {
  const snapshot = await getVisibleTextSnapshot(page);
  const noResultsByText = NO_RESULTS_REGEX.test(snapshot);
  const tableExists = await page.locator("table:visible").count().then((n) => n > 0).catch(() => false);
  const roleRows = await page.locator('[role="row"]:visible').count().catch(() => 0);
  const hasAnyGrid = tableExists || roleRows > 1;
  return {
    noResultsByText,
    hasAnyGrid,
    isNoResults: noResultsByText || !hasAnyGrid,
  };
}

async function getImportadorColumnIndex(page) {
  return page
    .evaluate(() => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const table = document.querySelector("table");
      if (!table) return -1;

      const headers = Array.from(table.querySelectorAll("thead th"));
      for (let i = 0; i < headers.length; i += 1) {
        if (normalize(headers[i].innerText || headers[i].textContent) === "importador") return i;
      }

      const firstRow = table.querySelector("tr");
      if (!firstRow) return -1;
      const anyHeaders = Array.from(firstRow.querySelectorAll("th, td"));
      for (let i = 0; i < anyHeaders.length; i += 1) {
        if (normalize(anyHeaders[i].innerText || anyHeaders[i].textContent) === "importador") return i;
      }
      return -1;
    })
    .catch(() => -1);
}

async function readImportadoresFromCurrentPage(page, importadorColIndex) {
  return page
    .evaluate((colIndex) => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const out = [];
      const table = document.querySelector("table");
      if (!table) return out;

      const rows = Array.from(table.querySelectorAll("tbody tr"));
      for (const row of rows) {
        const cells = Array.from(row.querySelectorAll("td"));
        if (!cells.length) continue;
        const idx = colIndex >= 0 ? colIndex : 0;
        const value = normalize(cells[idx]?.innerText || cells[idx]?.textContent || "");
        if (!value) continue;
        if (/^sin informaci[oó]n$/i.test(value)) continue;
        out.push(value);
      }
      return out;
    }, importadorColIndex)
    .catch(() => []);
}

async function readPaginationInfo(page) {
  return page
    .evaluate(() => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const text = normalize(document.body?.innerText || "");

      // Busca patrones como "254 / 274"
      const m = text.match(/(\d+)\s*\/\s*(\d+)/);
      if (m) {
        return {
          current: Number(m[1]),
          total: Number(m[2]),
          source: "ratio_text",
        };
      }

      const pageButtons = Array.from(document.querySelectorAll("button, a, [role='button']"))
        .map((el) => normalize(el.innerText || el.textContent || ""))
        .filter((v) => /^\d+$/.test(v))
        .map(Number);

      if (pageButtons.length > 0) {
        return {
          current: Math.min(...pageButtons),
          total: Math.max(...pageButtons),
          source: "numeric_buttons",
        };
      }

      return { current: 1, total: 1, source: "fallback" };
    })
    .catch(() => ({ current: 1, total: 1, source: "fallback" }));
}

async function goNextPage(page) {
  const moved = await page
    .evaluate(() => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const nodes = Array.from(
        document.querySelectorAll('button, a, [role="button"], .p-paginator-next, .mat-paginator-navigation-next')
      ).filter(visible);

      const next = nodes.find((el) => {
        const txt = normalize(el.innerText || el.textContent || "");
        const cls = (el.className || "").toString().toLowerCase();
        const aria = (el.getAttribute("aria-label") || "").toLowerCase();
        return (
          /siguiente|next|pr[oó]xima|›|»/.test(txt) ||
          cls.includes("next") ||
          aria.includes("next") ||
          aria.includes("siguiente")
        );
      });
      if (!next) return false;

      if (next.getAttribute("disabled") !== null) return false;
      if ((next.className || "").toString().toLowerCase().includes("disabled")) return false;

      next.click();
      return true;
    })
    .catch(() => false);

  if (!moved) return false;
  await waitForSettled(page);
  await sleep(350);
  return true;
}

async function returnToFiltersView(page, log) {
  const clicked = await page
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const nodes = Array.from(
        document.querySelectorAll('button, a, [role="button"], ion-button, .btn, .button')
      ).filter(visible);
      const back = nodes.find((el) => {
        const txt = normalize(el.innerText || el.textContent || el.value || "");
        return txt === "volver" || txt.includes("volver");
      });
      if (!back) return false;
      back.click();
      return true;
    })
    .catch(() => false);

  if (clicked) {
    log("volver al formulario ejecutado");
    await waitForSettled(page);
    await sleep(300);
    return;
  }

  // fallback: volver a cargar módulo
  log("No se detectó botón Volver; recargando módulo", "warn");
  await page.goto(TARGET_MODULE_URL, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
  await waitForSettled(page);
}

async function extractImportadoresForCountry(page, country, log) {
  const perCountrySet = new Set();
  let importadorColIndex = -1;
  let pagination = { current: 1, total: 1 };
  const visitedPages = new Set();

  const searchResult = await detectNoResultsState(page);
  if (searchResult.isNoResults) {
    return {
      hasResults: false,
      pages: 0,
      importadores: [],
    };
  }

  importadorColIndex = await getImportadorColumnIndex(page);
  if (importadorColIndex < 0) {
    throw new Error('No se detectó columna "Importador"');
  }

  pagination = await readPaginationInfo(page);
  const totalPages = Math.max(1, Number(pagination.total || 1));
  let currentPage = 1;

  for (;;) {
    const pageKey = `${country}::${currentPage}`;
    if (visitedPages.has(pageKey)) break;
    visitedPages.add(pageKey);

    // reintento por página si falla
    let pageRows = [];
    let ok = false;
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      try {
        pageRows = await readImportadoresFromCurrentPage(page, importadorColIndex);
        ok = true;
        break;
      } catch {
        if (attempt === 2) throw new Error(`Fallo leyendo página ${currentPage}`);
        await waitForSettled(page);
        await sleep(400);
      }
    }
    if (ok) {
      pageRows.forEach((name) => perCountrySet.add(normalizeText(name)));
      log("página procesada", "info", {
        pais: country,
        paginaActual: currentPage,
        totalPaginas: totalPages,
        importadoresPagina: pageRows.length,
        importadoresAcumuladosPais: perCountrySet.size,
      });
    }

    if (currentPage >= totalPages) break;
    const moved = await goNextPage(page);
    if (!moved) break;
    currentPage += 1;
  }

  return {
    hasResults: perCountrySet.size > 0,
    pages: visitedPages.size,
    importadores: [...perCountrySet],
  };
}

async function main() {
  const log = makeLogger("main");
  await ensureDirectories();

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await createContextWithAuthState(browser, log);
  const page = await context.newPage();
  page.setDefaultTimeout(CONFIG.timeoutMs);

  const paisesSinResultados = [];
  const paisesConError = [];
  const porPais = [];
  const globalSet = new Set();

  try {
    await ensureLoggedIn(page, context, makeLogger("auth"));
    await navigateAndEnsureSession(page, context, log, TARGET_MODULE_URL);

    // 1) obtener lista completa de países al inicio
    const countries = await getCountriesFromOriginField(page, log);
    if (!countries.length) {
      throw new Error("La lista de países de origen salió vacía");
    }
    await saveJson(OUTPUTS.paisesOrigen, countries);
    log("Lista de países detectada", "info", { total: countries.length });

    // 2) iterar países uno por uno en orden
    for (const country of countries) {
      const countryLog = makeLogger(`pais:${country}`);
      countryLog("inicio país");

      try {
        // asegurar sesión y formulario
        await navigateAndEnsureSession(page, context, countryLog, TARGET_MODULE_URL);
        await selectCountryExact(page, country, countryLog);
        await clickBuscar(page, countryLog);

        const searchState = await detectNoResultsState(page);
        if (searchState.isNoResults) {
          paisesSinResultados.push(country);
          countryLog("sin resultados", "info", { pais: country });
          await returnToFiltersView(page, countryLog);
          continue;
        }

        const result = await extractImportadoresForCountry(page, country, countryLog);
        if (!result.hasResults) {
          paisesSinResultados.push(country);
          countryLog("sin resultados (tabla vacía)", "info", { pais: country });
          await returnToFiltersView(page, countryLog);
          continue;
        }

        for (const importer of result.importadores) {
          const importador = normalizeText(importer);
          if (!importador) continue;
          porPais.push({ importador, paisOrigen: country });
          globalSet.add(importador);
        }

        countryLog("país finalizado", "info", {
          pais: country,
          paginasRecorridas: result.pages,
          importadoresUnicosPais: result.importadores.length,
          importadoresUnicosGlobal: globalSet.size,
        });

        await returnToFiltersView(page, countryLog);
      } catch (error) {
        paisesConError.push({
          pais: country,
          error: summarizeError(error),
        });
        countryLog("error de país (continúa siguiente)", "error", {
          pais: country,
          error: summarizeError(error),
        });
        await navigateAndEnsureSession(page, context, countryLog, TARGET_MODULE_URL).catch(() => {});
      }
    }

    // 3) dedupe final por importador+pais y set global
    const dedupPorPaisMap = new Map();
    for (const row of porPais) {
      const key = `${row.importador.toLowerCase()}|${row.paisOrigen.toLowerCase()}`;
      if (!dedupPorPaisMap.has(key)) dedupPorPaisMap.set(key, row);
    }
    const dedupPorPais = [...dedupPorPaisMap.values()];
    const importadoresUnicos = [...new Set([...globalSet])].map((importador) => ({ importador }));

    // 4) guardar archivos solicitados
    await saveJson(OUTPUTS.paisesSinResultados, paisesSinResultados);
    await saveJson(OUTPUTS.paisesConError, paisesConError);

    await saveJson(OUTPUTS.porPaisJson, dedupPorPais);
    await writeCsv(OUTPUTS.porPaisCsv, ["importador", "paisOrigen"], dedupPorPais);
    await writeXlsx(OUTPUTS.porPaisXlsx, "importadores_por_pais", ["importador", "paisOrigen"], dedupPorPais);

    await saveJson(OUTPUTS.unicosJson, importadoresUnicos);
    await writeCsv(OUTPUTS.unicosCsv, ["importador"], importadoresUnicos);
    await writeXlsx(OUTPUTS.unicosXlsx, "importadores_unicos", ["importador"], importadoresUnicos);

    log("Extracción finalizada", "info", {
      totalPaises: countries.length,
      paisesSinResultados: paisesSinResultados.length,
      paisesConError: paisesConError.length,
      importadoresPorPais: dedupPorPais.length,
      importadoresUnicos: importadoresUnicos.length,
      outputs: OUTPUTS,
    });
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(`[fatal] ${summarizeError(error)}`);
  process.exit(1);
});

