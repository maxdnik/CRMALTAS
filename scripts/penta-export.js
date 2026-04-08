#!/usr/bin/env node
"use strict";

const fs = require("node:fs/promises");
const fssync = require("node:fs");
const path = require("node:path");
const dotenv = require("dotenv");

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
  fullHtml: path.join(DATA_DIR, "importadores-module-full.html"),
  fullPng: path.join(DATA_DIR, "importadores-module-full.png"),
  openedHtml: path.join(DATA_DIR, "pais-origen-opened.html"),
  openedPng: path.join(DATA_DIR, "pais-origen-opened.png"),
  analysisJson: path.join(DATA_DIR, "importadores-ui-analysis.json"),
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
      const txt = (document.body?.innerText || "").replace(/\s+/g, " ").trim().slice(0, 3000);
      return {
        url,
        isLoginUrl: /\/login(?:\/|$|\?)/i.test(url.toLowerCase()),
        hasLoginText: loginRegex.test(txt),
      };
    }, LOGIN_TEXT_REGEX.source)
    .catch(() => ({
      url: page.url(),
      isLoginUrl: /\/login(?:\/|$|\?)/i.test(page.url().toLowerCase()),
      hasLoginText: false,
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
    if (await loc.isVisible().catch(() => false)) return { locator: loc, selector };
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

  await userCandidate.locator.fill(CONFIG.user, { timeout: 12000 });
  await passCandidate.locator.fill(CONFIG.pass, { timeout: 12000 });

  const submit = await getFirstVisibleLocator(page, [
    'button[type="submit"]',
    'input[type="submit"]',
    'ion-button[type="submit"]',
    'button:has-text("Enter")',
    'button:has-text("Ingresar")',
  ]);

  if (submit) {
    await submit.locator.click({ timeout: 10000 }).catch(() => {});
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
  await waitForSettled(page);

  const postState = await isLoginLikeState(page);
  const urlLower = String(postState.url || "").toLowerCase();
  if (urlLower.includes("/olvide-mi-password")) {
    await saveLoginFailureArtifacts(page, log);
    throw new Error("Login fallido: se navegó a recuperación de contraseña");
  }

  const loginSuccess = !urlLower.includes("/login") && urlLower.includes("/home");
  if (!loginSuccess) {
    await saveLoginFailureArtifacts(page, log);
    throw new Error("Login fallido: sigue en /login");
  }

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

async function captureInitialModuleArtifacts(page, log) {
  await page.screenshot({ path: OUTPUTS.fullPng, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) await fs.writeFile(OUTPUTS.fullHtml, html, "utf8");
  log("Archivos iniciales de módulo guardados", "info", {
    html: OUTPUTS.fullHtml,
    png: OUTPUTS.fullPng,
  });
}

async function collectCountryOriginCandidates(page) {
  return page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const visible = (el) => {
      const st = window.getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
    };
    const shortHtml = (el) => (el.outerHTML || "").replace(/\s+/g, " ").trim().slice(0, 300);
    const bbox = (el) => {
      const r = el.getBoundingClientRect();
      return { x: r.x, y: r.y, width: r.width, height: r.height };
    };

    const roots = [];
    const byLabelText = Array.from(document.querySelectorAll("*")).filter((el) =>
      /pa[ií]s de origen/i.test(normalize(el.textContent || ""))
    );
    roots.push(...byLabelText);
    roots.push(...Array.from(document.querySelectorAll('label, [aria-label*="País" i], [role="combobox"], [role="listbox"], [role="option"], ion-select, ion-item, ion-input, p-dropdown, ng-select, mat-select, .p-dropdown, .ng-select, .mat-mdc-select')));
    roots.push(...Array.from(document.querySelectorAll('input[name*="pais" i], input[id*="pais" i], [placeholder*="País de Origen" i], [placeholder*="pais" i]')));

    const uniq = [];
    const seen = new Set();
    for (const el of roots) {
      if (!el || seen.has(el)) continue;
      seen.add(el);

      const near = el.closest("ion-item, .p-field, .form-group, .row, .col, div, section, mat-form-field, p-dropdown, ng-select") || el;
      const keyEl = near || el;
      if (seen.has(keyEl)) continue;
      seen.add(keyEl);

      const id = keyEl.id || "";
      const name = keyEl.getAttribute("name") || "";
      const placeholder = keyEl.getAttribute("placeholder") || "";
      const role = keyEl.getAttribute("role") || "";
      const ariaLabel = keyEl.getAttribute("aria-label") || "";
      const ariaExpanded = keyEl.getAttribute("aria-expanded") || "";
      const textContent = normalize(keyEl.textContent || "");
      const innerText = normalize(keyEl.innerText || "");
      const className = normalize(keyEl.className || "");
      const tagName = (keyEl.tagName || "").toLowerCase();
      const isVisible = visible(keyEl);
      const box = bbox(keyEl);

      uniq.push({
        tagName,
        textContent: textContent.slice(0, 250),
        innerText: innerText.slice(0, 250),
        className: className.slice(0, 250),
        id,
        name,
        placeholder,
        role,
        ariaLabel,
        ariaExpanded,
        outerHTML: shortHtml(keyEl),
        boundingBox: box,
        visible: isVisible,
      });
    }

    return uniq
      .filter((c) => {
        const all = `${c.textContent} ${c.innerText} ${c.className} ${c.id} ${c.name} ${c.placeholder} ${c.role} ${c.ariaLabel}`.toLowerCase();
        return /pa[ií]s|origen|country|combobox|dropdown|select|option|listbox|ion-|p-dropdown|ng-select|mat-select/.test(all);
      })
      .slice(0, 120);
  });
}

async function detectOpenedPanels(page) {
  return page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const visible = (el) => {
      const st = window.getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
    };

    const overlaySelectors = [
      '[role="listbox"]',
      '[role="option"]',
      ".cdk-overlay-pane",
      ".cdk-overlay-container",
      ".p-dropdown-panel",
      ".ng-dropdown-panel",
      ".mat-mdc-select-panel",
      "ion-alert",
      "ion-popover",
      "ion-modal",
      ".dropdown-menu",
      ".menu",
      ".popover",
      ".overlay",
    ];

    const overlays = Array.from(document.querySelectorAll(overlaySelectors.join(","))).filter(visible);
    const options = Array.from(
      document.querySelectorAll('[role="option"], .p-dropdown-item, .ng-option, mat-option, .mat-mdc-option, ion-select-option, li')
    )
      .filter(visible)
      .map((el) => normalize(el.innerText || el.textContent || ""))
      .filter((t) => t.length > 0);

    return {
      overlayCount: overlays.length,
      optionCount: options.length,
      firstOptions: options.slice(0, 30),
      hasListbox: document.querySelectorAll('[role="listbox"]').length > 0,
      hasOptionRole: document.querySelectorAll('[role="option"]').length > 0,
      optionDomSample:
        overlays[0]?.outerHTML?.replace(/\s+/g, " ").trim().slice(0, 400) ||
        (document.querySelector('[role="option"], .p-dropdown-item, .ng-option, mat-option, .mat-mdc-option, ion-select-option, li')?.outerHTML || "")
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, 400),
    };
  });
}

async function tryOpenCountryFilterWithStrategies(page, candidates, log) {
  const attempts = [];

  const runAttempt = async (strategyName, fn) => {
    let success = false;
    let error = "";
    try {
      await fn();
      await sleep(450);
      const panelState = await detectOpenedPanels(page);
      success = panelState.overlayCount > 0 || panelState.optionCount > 0 || panelState.hasListbox;
      attempts.push({
        strategy: strategyName,
        success,
        panelState,
      });
      log(`Intento apertura País de Origen: ${strategyName}`, success ? "info" : "warn", {
        overlayCount: panelState.overlayCount,
        optionCount: panelState.optionCount,
      });
      return success;
    } catch (e) {
      error = summarizeError(e);
      attempts.push({ strategy: strategyName, success: false, error });
      log(`Intento apertura País de Origen falló: ${strategyName}`, "warn", { error });
      return false;
    }
  };

  const labelLike = page.getByText("País de Origen", { exact: true }).first();
  const inputLike = page.locator('[placeholder*="País de Origen" i], [aria-label*="País de Origen" i], input[name*="pais" i], input[id*="pais" i]').first();
  const arrowLike = page.locator('.p-dropdown-trigger, .ng-arrow-wrapper, .mat-mdc-select-arrow, ion-icon, [class*="arrow"], [class*="chevron"]').first();
  const comboRole = page.locator('[role="combobox"]').first();

  const firstVisibleCandidate = candidates.find((c) => c.visible && c.boundingBox?.width > 10 && c.boundingBox?.height > 10);

  const strategies = [
    {
      name: "click-label",
      fn: async () => {
        if (await labelLike.isVisible().catch(() => false)) await labelLike.click({ timeout: 3000 });
      },
    },
    {
      name: "click-input",
      fn: async () => {
        if (await inputLike.isVisible().catch(() => false)) await inputLike.click({ timeout: 3000 });
      },
    },
    {
      name: "click-combobox-role",
      fn: async () => {
        if (await comboRole.isVisible().catch(() => false)) await comboRole.click({ timeout: 3000 });
      },
    },
    {
      name: "click-arrow-icon",
      fn: async () => {
        if (await arrowLike.isVisible().catch(() => false)) await arrowLike.click({ timeout: 3000 });
      },
    },
    {
      name: "focus-and-arrowdown",
      fn: async () => {
        if (await inputLike.isVisible().catch(() => false)) {
          await inputLike.focus();
          await page.keyboard.press("ArrowDown");
        } else if (await comboRole.isVisible().catch(() => false)) {
          await comboRole.focus();
          await page.keyboard.press("ArrowDown");
        }
      },
    },
    {
      name: "bbox-center-click",
      fn: async () => {
        if (!firstVisibleCandidate) return;
        const { x, y, width, height } = firstVisibleCandidate.boundingBox;
        const cx = x + width / 2;
        const cy = y + height / 2;
        await page.mouse.move(cx, cy);
        await page.mouse.click(cx, cy);
      },
    },
  ];

  let winner = null;
  for (const strategy of strategies) {
    const ok = await runAttempt(strategy.name, strategy.fn);
    if (ok) {
      winner = strategy.name;
      break;
    }
  }

  if (winner) {
    await page.screenshot({ path: OUTPUTS.openedPng, fullPage: true }).catch(() => {});
    const openedHtml = await page.content().catch(() => "");
    if (openedHtml) await fs.writeFile(OUTPUTS.openedHtml, openedHtml, "utf8");
  }

  const openedState = await detectOpenedPanels(page);
  return { attempts, winner, openedState };
}

async function selectCountryByText(page, country, log) {
  const result = await page
    .evaluate((countryArg) => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };

      const wanted = normalize(countryArg).toLowerCase();
      const options = Array.from(
        document.querySelectorAll('[role="option"], .p-dropdown-item, .ng-option, mat-option, .mat-mdc-option, ion-select-option, li')
      ).filter(visible);

      const exact = options.find((el) => normalize(el.innerText || el.textContent || "").toLowerCase() === wanted);
      if (exact) {
        exact.click();
        return { selected: true, mode: "click-option", selectedText: normalize(exact.innerText || exact.textContent || "") };
      }

      const searchInput = Array.from(
        document.querySelectorAll('input[type="search"], input[placeholder*="buscar" i], input[placeholder*="search" i], input[type="text"]')
      ).find(visible);

      if (searchInput) {
        searchInput.value = countryArg;
        searchInput.dispatchEvent(new Event("input", { bubbles: true }));
        searchInput.dispatchEvent(new Event("change", { bubbles: true }));
      }

      const optionsAfter = Array.from(
        document.querySelectorAll('[role="option"], .p-dropdown-item, .ng-option, mat-option, .mat-mdc-option, ion-select-option, li')
      ).filter(visible);
      const exactAfter = optionsAfter.find((el) => normalize(el.innerText || el.textContent || "").toLowerCase() === wanted);
      if (exactAfter) {
        exactAfter.click();
        return { selected: true, mode: "search+click-option", selectedText: normalize(exactAfter.innerText || exactAfter.textContent || "") };
      }

      return { selected: false, mode: "not-found", selectedText: "" };
    }, country)
    .catch(() => ({ selected: false, mode: "error", selectedText: "" }));

  log(`Selección de país prueba (${country})`, result.selected ? "info" : "warn", result);
  return result;
}

async function findAndClickBuscar(page, log) {
  const clicked = await page
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const buttons = Array.from(
        document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], ion-button')
      ).filter(visible);
      const target = buttons.find((el) => normalize(el.innerText || el.textContent || el.value || "") === "buscar");
      if (!target) return false;
      target.click();
      return true;
    })
    .catch(() => false);

  if (!clicked) {
    const fallback = page.getByText("Buscar", { exact: true }).first();
    if (await fallback.isVisible().catch(() => false)) {
      await fallback.click({ timeout: 5000 }).catch(() => {});
      return true;
    }
    log('No se encontró botón "Buscar"', "warn");
    return false;
  }
  return true;
}

async function analyzeSearchResultState(page, log) {
  await waitForSettled(page);
  await sleep(600);
  const state = await page.evaluate((noResultsRegexSrc) => {
    const noResultsRegex = new RegExp(noResultsRegexSrc, "i");
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const visibleText = normalize(document.body?.innerText || "");
    const noResultsByText = noResultsRegex.test(visibleText);
    const table = document.querySelector("table");
    const tableVisible = !!table;
    const rows = table ? table.querySelectorAll("tbody tr").length : 0;

    return {
      noResultsByText,
      tableVisible,
      tableRowCount: rows,
      visibleTextPreview: visibleText.slice(0, 500),
    };
  }, NO_RESULTS_REGEX.source);

  log("Resultado tras Buscar", "info", state);
  return state;
}

async function detectImportadorColumnAndPagination(page, log) {
  return page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const result = {
      tableSelector: "",
      importadorColumnIndex: -1,
      importadorHeaderText: "",
      paginationSelector: "",
      paginationInfo: "",
      backSelectorHint: "",
    };

    const table = document.querySelector("table");
    if (!table) return result;
    result.tableSelector = "table";

    const headers = Array.from(table.querySelectorAll("thead th, tr th"));
    for (let i = 0; i < headers.length; i += 1) {
      const h = normalize(headers[i].innerText || headers[i].textContent || "");
      if (h.toLowerCase() === "importador" || h.toLowerCase().includes("importador")) {
        result.importadorColumnIndex = i;
        result.importadorHeaderText = h;
        break;
      }
    }

    const pagCandidates = [
      ".p-paginator",
      ".mat-paginator",
      ".pagination",
      '[aria-label*="pagination" i]',
      '[class*="paginator"]',
    ];
    const pagEl = pagCandidates.map((s) => document.querySelector(s)).find(Boolean);
    if (pagEl) {
      result.paginationSelector = pagCandidates.find((s) => document.querySelector(s)) || "";
      result.paginationInfo = normalize(pagEl.innerText || pagEl.textContent || "").slice(0, 250);
    }

    const backCandidates = Array.from(
      document.querySelectorAll('button, [role="button"], a, ion-button, .btn, .button')
    ).find((el) => /volver/i.test(normalize(el.innerText || el.textContent || el.value || "")));
    if (backCandidates) {
      result.backSelectorHint = `${(backCandidates.tagName || "").toLowerCase()}${backCandidates.className ? "." + String(backCandidates.className).replace(/\s+/g, ".") : ""}`.slice(0, 180);
    }
    return result;
  }).then((result) => {
    log("Detección tabla/paginación/importador", "info", result);
    return result;
  });
}

async function returnToParametersIfPossible(page, log) {
  const clicked = await page
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const nodes = Array.from(
        document.querySelectorAll('button, [role="button"], a, ion-button, .btn, .button')
      ).filter(visible);
      const back = nodes.find((el) => normalize(el.innerText || el.textContent || el.value || "").includes("volver"));
      if (!back) return false;
      back.click();
      return true;
    })
    .catch(() => false);

  log("Volver a parámetros", clicked ? "info" : "warn", { clicked });
  if (clicked) {
    await waitForSettled(page);
    await sleep(400);
  }
}

function buildStrategySummary(analysisState) {
  const winner = analysisState.opening?.winner || "";
  const firstAttemptWithOptions = (analysisState.opening?.attempts || []).find(
    (a) => a.panelState && a.panelState.optionCount > 0
  );
  const sampleOptions = analysisState.opening?.openedState?.firstOptions || [];
  const selectedAfghanistan = analysisState.tests?.afganistan?.selected || false;
  const selectedGermany = analysisState.tests?.alemania?.selected || false;

  return {
    selectorOpenPaisOrigen: winner || "click en contenedor cercano a label País de Origen",
    selectorChooseOption: "role=option / .p-dropdown-item / .ng-option (texto exacto)",
    selectorBuscar: "button/ion-button con texto exacto 'Buscar'",
    selectorNoResultados: "texto visible /No se encontraron resultados/i",
    selectorTabla: analysisState.table?.tableSelector || "table",
    selectorColumnaImportador:
      analysisState.table?.importadorColumnIndex >= 0
        ? `table thead th[index=${analysisState.table.importadorColumnIndex}] => Importador`
        : "detectar header por texto 'Importador'",
    selectorPaginacion: analysisState.table?.paginationSelector || "contenedor paginador (.p-paginator/.mat-paginator/.pagination)",
    selectorVolverParametros: analysisState.table?.backSelectorHint || "botón/acción con texto 'Volver'",
    strategyEndToEnd: [
      "1) Abrir módulo importDetalladas autenticado",
      "2) Abrir País de Origen usando estrategia ganadora",
      "3) Seleccionar país por texto exacto en opción visible",
      "4) Click Buscar",
      "5) Si texto 'No se encontraron resultados' => país sin datos",
      "6) Si hay tabla => detectar columna Importador y leer filas",
      "7) Recorrer paginación con botón siguiente/índices hasta agotar",
      "8) Volver a parámetros y repetir país siguiente",
    ],
    proof: {
      openedWinner: winner,
      optionsDetected: sampleOptions.slice(0, 30),
      afganistanSelectionWorked: selectedAfghanistan,
      alemaniaSelectionWorked: selectedGermany,
    },
  };
}

async function runAnalysis() {
  await ensureDirectories();
  const log = makeLogger("analysis");

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await createContextWithAuthState(browser, log);
  const page = await context.newPage();
  page.setDefaultTimeout(CONFIG.timeoutMs);

  const state = {
    startedAt: nowIso(),
    moduleUrl: TARGET_MODULE_URL,
    candidates: [],
    opening: null,
    tests: {},
    searchState: null,
    table: null,
    summary: null,
    errors: [],
  };

  try {
    await ensureLoggedIn(page, context, makeLogger("auth"));
    await navigateAndEnsureSession(page, context, log, TARGET_MODULE_URL);
    await captureInitialModuleArtifacts(page, log);

    const candidates = await collectCountryOriginCandidates(page);
    state.candidates = candidates;
    log("Candidatos relacionados a País de Origen detectados", "info", { count: candidates.length });

    const opening = await tryOpenCountryFilterWithStrategies(page, candidates, log);
    state.opening = opening;

    if (opening.winner) {
      const af = await selectCountryByText(page, "Afganistán", log);
      state.tests.afganistan = af;
      await sleep(300);
      const de = await selectCountryByText(page, "Alemania", log);
      state.tests.alemania = de;
      await sleep(300);
    } else {
      state.errors.push('No se pudo abrir "País de Origen" con estrategias automáticas');
    }

    const clickedBuscar = await findAndClickBuscar(page, log);
    if (clickedBuscar) {
      state.searchState = await analyzeSearchResultState(page, log);
      if (state.searchState.tableVisible && !state.searchState.noResultsByText) {
        state.table = await detectImportadorColumnAndPagination(page, log);
      }
      await returnToParametersIfPossible(page, log);
    } else {
      state.errors.push('No se pudo clickear botón "Buscar"');
    }
  } catch (error) {
    state.errors.push(summarizeError(error));
    log("Error durante la fase de análisis", "error", { error: summarizeError(error) });
  } finally {
    state.summary = buildStrategySummary(state);
    state.finishedAt = nowIso();
    await saveJson(OUTPUTS.analysisJson, state);
    log("Análisis UI guardado", "info", { file: OUTPUTS.analysisJson });
    await browser.close();
  }
}

runAnalysis().catch((error) => {
  console.error(`[fatal] ${summarizeError(error)}`);
  process.exit(1);
});
