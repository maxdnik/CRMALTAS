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
const NETWORK_DEBUG_PATH = path.join(DATA_DIR, "network-debug.json");
const AUTH_STATE_PATH = path.join(DATA_DIR, "auth-state.json");
const LOGIN_FAILED_SCREENSHOT = path.join(DATA_DIR, "login-failed.png");
const LOGIN_FAILED_HTML = path.join(DATA_DIR, "login-failed.html");
const LOGIN_SUCCESS_SCREENSHOT = path.join(DATA_DIR, "login-success.png");

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

const SEARCH_TEXT_REGEX = /(buscar|consultar|aplicar|filtrar|ver resultados|search|submit)/i;
const EMPTY_TEXT_REGEX = /(sin resultados|no records|no data|sin datos)/i;
const KEYWORD_REGEX = /(import|export|detalle|data|formulario|search|grid)/i;
const LOGIN_TEXT_REGEX =
  /(user login|user password|forgot my password|enter|iniciar sesi[oó]n|ingresar|password)/i;

function nowIso() {
  return new Date().toISOString();
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

function safeSlug(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 120) || "na";
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

async function createContextWithAuthState(browser, log) {
  const baseContextOptions = {
    acceptDownloads: true,
    viewport: { width: 1600, height: 1000 },
  };

  if (fssync.existsSync(AUTH_STATE_PATH)) {
    try {
      log("Intentando reutilizar auth-state", "info", { path: AUTH_STATE_PATH });
      return await browser.newContext({
        ...baseContextOptions,
        storageState: AUTH_STATE_PATH,
      });
    } catch (error) {
      log("No se pudo cargar auth-state, se inicia sesión limpia", "warn", {
        error: summarizeError(error),
      });
    }
  }

  return browser.newContext(baseContextOptions);
}

async function waitForSettled(page) {
  await page.waitForLoadState("domcontentloaded", { timeout: 30000 }).catch(() => {});
  await page.waitForLoadState("networkidle", { timeout: 15000 }).catch(() => {});
}

async function retry(action, options = {}) {
  const { retries = 3, delayMs = 1000, actionName = "action", onRetry = () => {} } = options;
  let lastError;
  for (let attempt = 1; attempt <= retries; attempt += 1) {
    try {
      return await action(attempt);
    } catch (error) {
      lastError = error;
      if (attempt < retries) {
        await onRetry(attempt, error);
        await sleep(delayMs * attempt);
      }
    }
  }
  throw new Error(`${actionName} failed: ${summarizeError(lastError)}`);
}

async function isLoginLikeState(page) {
  return page
    .evaluate((loginRegexSource) => {
      const loginRegex = new RegExp(loginRegexSource, "i");
      const url = window.location.href;
      const urlLower = url.toLowerCase();
      const visibleText = (document.body?.innerText || "").replace(/\s+/g, " ").trim().slice(0, 3000);
      const hasPasswordInput = document.querySelectorAll('input[type="password"]').length > 0;
      const hasLoginText = loginRegex.test(visibleText);
      const isLoginUrl = /\/login(?:\/|$|\?)/i.test(urlLower) || /signin|session|sesion|ingresar/.test(urlLower);
      return {
        url,
        isLoginUrl,
        hasPasswordInput,
        hasLoginText,
        isLoginLike: isLoginUrl || hasPasswordInput || hasLoginText,
        visibleTextPreview: visibleText.slice(0, 600),
      };
    }, LOGIN_TEXT_REGEX.source)
    .catch(() => ({
      url: page.url(),
      isLoginUrl: /\/login(?:\/|$|\?)/i.test(page.url().toLowerCase()),
      hasPasswordInput: false,
      hasLoginText: false,
      isLoginLike: /\/login(?:\/|$|\?)/i.test(page.url().toLowerCase()),
      visibleTextPreview: "",
    }));
}

async function saveLoginFailureArtifacts(page, log) {
  await page
    .screenshot({ path: LOGIN_FAILED_SCREENSHOT, fullPage: true })
    .then(() => log("Screenshot de login fallido guardado", "warn", { path: LOGIN_FAILED_SCREENSHOT }))
    .catch((error) => log("No se pudo guardar screenshot de login fallido", "warn", { error: summarizeError(error) }));
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
    .catch((error) => log("No se pudo guardar screenshot de login exitoso", "warn", { error: summarizeError(error) }));
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
  const loginUrl = `${CONFIG.baseUrl}/login`;

  await page.goto(loginUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
  await waitForSettled(page);

  const preState = await isLoginLikeState(page);
  if (!preState.isLoginLike && !preState.isLoginUrl) {
    log("Sesión ya activa al abrir /login", "info", { url: preState.url });
    await context.storageState({ path: AUTH_STATE_PATH });
    log("Estado autenticado guardado", "info", { path: AUTH_STATE_PATH });
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
    log("No se detectaron campos de login", "error", {
      userFound: Boolean(userCandidate),
      passFound: Boolean(passCandidate),
      url: page.url(),
    });
    await saveLoginFailureArtifacts(page, log);
    throw new Error("No se detectaron campos de usuario/password en /login");
  }

  log("campos detectados", "info", {
    userSelector: userCandidate.selector,
    passSelector: passCandidate.selector,
  });

  await userCandidate.locator.fill(CONFIG.user, { timeout: 15000 });
  await passCandidate.locator.fill(CONFIG.pass, { timeout: 15000 });

  const submitCandidates = await page
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
      const form =
        document.querySelector('input[type="password"]')?.closest("form") ||
        document.querySelector('input[type="text"], input[type="email"]')?.closest("form") ||
        document.querySelector("form");
      if (!form) return [];

      const all = Array.from(form.querySelectorAll("button, input, a, [role='button'], div, span")).filter(visible);
      return all.map((el, idx) => ({
        index: idx,
        tagName: (el.tagName || "").toLowerCase(),
        text: normalize(el.innerText || el.textContent || el.value || ""),
        type: normalize(el.getAttribute("type") || ""),
        href: normalize(el.getAttribute("href") || ""),
        selectorHint: `${(el.tagName || "").toLowerCase()}${el.getAttribute("type") ? `[type="${el.getAttribute("type")}"]` : ""}`,
      }));
    })
    .catch(() => []);

  log("candidatos de submit detectados", "info", {
    total: submitCandidates.length,
    candidates: submitCandidates.slice(0, 40).map((c) => ({
      tagName: c.tagName,
      text: c.text,
      type: c.type,
      href: c.href,
    })),
  });

  const isDisallowedCandidate = (candidate) => {
    const text = String(candidate.text || "").toLowerCase();
    const href = String(candidate.href || "").toLowerCase();
    return (
      candidate.tagName === "a" ||
      text.includes("forgot") ||
      text.includes("password") ||
      href.includes("olvide")
    );
  };

  const findCandidateByPriority = () => {
    const allowed = submitCandidates.filter((c) => !isDisallowedCandidate(c));
    const exactText = (value) => (txt) => String(txt || "").trim().toLowerCase() === value;
    return (
      allowed.find((c) => c.tagName === "button" && c.type === "submit") ||
      allowed.find((c) => c.tagName === "input" && c.type === "submit") ||
      allowed.find((c) => c.tagName === "button" && exactText("enter")(c.text)) ||
      allowed.find((c) => c.tagName === "button" && exactText("ingresar")(c.text)) ||
      null
    );
  };

  const selectedSubmit = findCandidateByPriority();
  log("submit elegido", "info", {
    selector: selectedSubmit?.selectorHint || "form.requestSubmit()/form.submit()",
    tagName: selectedSubmit?.tagName || "form",
    text: selectedSubmit?.text || "",
    type: selectedSubmit?.type || "",
    href: selectedSubmit?.href || "",
  });

  const previousUrl = page.url();
  if (selectedSubmit) {
    const clicked = await page
      .evaluate((target) => {
        const visible = (el) => {
          const st = window.getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
        };
        const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
        const form =
          document.querySelector('input[type="password"]')?.closest("form") ||
          document.querySelector('input[type="text"], input[type="email"]')?.closest("form") ||
          document.querySelector("form");
        if (!form) return false;
        const all = Array.from(form.querySelectorAll("button, input, a, [role='button'], div, span")).filter(visible);
        const node = all[target.index];
        if (!node) return false;
        const tagName = (node.tagName || "").toLowerCase();
        const text = normalize(node.innerText || node.textContent || node.value || "").toLowerCase();
        const href = normalize(node.getAttribute("href") || "").toLowerCase();
        if (tagName === "a" || text.includes("forgot") || text.includes("password") || href.includes("olvide")) {
          return false;
        }
        node.click();
        return true;
      }, selectedSubmit)
      .catch(() => false);
    if (!clicked) {
      log("No se pudo clickear submit seleccionado, se usa submit del formulario", "warn");
      await page
        .locator("form")
        .first()
        .evaluate((form) => {
          if (form.requestSubmit) form.requestSubmit();
          else form.submit();
        })
        .catch(() => {});
    } else {
      log("submit ejecutado", "info", {
        tagName: selectedSubmit.tagName,
        text: selectedSubmit.text,
        type: selectedSubmit.type,
      });
    }
  } else {
    log("No se encontró submit explícito válido, se hace submit del formulario", "warn");
    await page
      .locator("form")
      .first()
      .evaluate((form) => {
        if (form.requestSubmit) form.requestSubmit();
        else form.submit();
      })
      .catch(() => {});
  }

  await Promise.race([
    page.waitForURL((url) => !url.toString().toLowerCase().includes("/login"), { timeout: 20000 }).catch(() => null),
    page.locator('input[type="password"]').first().waitFor({ state: "hidden", timeout: 20000 }).catch(() => null),
    sleep(4500),
  ]);
  await waitForSettled(page);

  const postState = await isLoginLikeState(page);
  const postUrlLower = String(postState.url || "").toLowerCase();
  log("url post-login", "info", {
    previousUrl,
    currentUrl: postState.url,
    isLoginUrl: postState.isLoginUrl,
    hasLoginText: postState.hasLoginText,
    hasPasswordInput: postState.hasPasswordInput,
  });

  if (postUrlLower.includes("/olvide-mi-password") || postUrlLower.includes("olvide")) {
    log("se hizo click en recuperación de contraseña por error", "error", { currentUrl: postState.url });
    await saveLoginFailureArtifacts(page, log);
    throw new Error("Login fallido: se navegó a recuperación de contraseña");
  }

  const loginFailed = postState.isLoginUrl || postState.hasLoginText || postState.hasPasswordInput;
  if (loginFailed) {
    log("login fallido", "error", {
      url: postState.url,
      visibleTextPreview: postState.visibleTextPreview,
    });
    await saveLoginFailureArtifacts(page, log);
    throw new Error("Login fallido: sigue en /login o la UI mantiene señales de formulario de login");
  }

  await context.storageState({ path: AUTH_STATE_PATH });
  log("login exitoso", "info", { url: postState.url });
  await saveLoginSuccessScreenshot(page, log);
  log("Estado autenticado guardado", "info", { path: AUTH_STATE_PATH });
}

function frameScore(frameInfo) {
  return (
    Number(frameInfo?.tr || 0) +
    Number(frameInfo?.roleRow || 0) * 2 +
    Number(frameInfo?.agRow || 0) * 2 +
    Number(frameInfo?.pDatatableRows || 0) * 2 +
    Number(frameInfo?.table || 0)
  );
}

function isLoginLikeFrameInfo(frameInfo) {
  if (!frameInfo) return false;
  const frameUrl = String(frameInfo.url || "").toLowerCase();
  const visibleText = String(frameInfo.visibleText || "");
  const isLoginUrl = /\/login(?:\/|$|\?)/i.test(frameUrl) || /signin|session|sesion|ingresar/.test(frameUrl);
  const hasLoginText = LOGIN_TEXT_REGEX.test(visibleText);
  return isLoginUrl || hasLoginText;
}

function createNetworkCollector(page, log, bucket) {
  const collected = [];

  const onRequest = (request) => {
    try {
      const resourceType = request.resourceType();
      if (!/(xhr|fetch)/i.test(resourceType)) return;
      const url = request.url();
      const body = request.postData() || "";
      const relevant = KEYWORD_REGEX.test(url) || KEYWORD_REGEX.test(body);
      if (!relevant) return;
      collected.push({
        ts: nowIso(),
        type: "request",
        method: request.method(),
        url,
        resourceType,
        requestBodyPreview: body.slice(0, 1500),
      });
    } catch (error) {
      log("Error capturando request", "warn", { error: summarizeError(error) });
    }
  };

  const onResponse = async (response) => {
    try {
      const req = response.request();
      const resourceType = req.resourceType();
      if (!/(xhr|fetch)/i.test(resourceType)) return;
      const url = response.url();
      const contentType = response.headers()["content-type"] || "";
      const relevantUrl = KEYWORD_REGEX.test(url);
      const isJsonLike = /json|javascript|problem\+json/i.test(contentType);
      const item = {
        ts: nowIso(),
        type: "response",
        method: req.method(),
        url,
        status: response.status(),
        resourceType,
        contentType,
        relevantUrl,
        hasJson: false,
        jsonPreview: "",
        parsedRecords: 0,
      };

      if (isJsonLike || relevantUrl) {
        const text = await response.text().catch(() => "");
        if (text) {
          try {
            const parsed = JSON.parse(text);
            const jsonString = JSON.stringify(parsed);
            const relevantBody = KEYWORD_REGEX.test(jsonString);
            if (relevantUrl || relevantBody) {
              item.hasJson = true;
              item.jsonPreview = jsonString.slice(0, 4000);
              item.parsedRecords = extractRecordsFromJsonPayload(parsed, url).length;
            }
          } catch {
            // ignore non-json
          }
        }
      }
      if (item.relevantUrl || item.hasJson) {
        collected.push(item);
      }
    } catch (error) {
      log("Error capturando response", "warn", { error: summarizeError(error) });
    }
  };

  page.on("request", onRequest);
  page.on("response", onResponse);

  return {
    getAll: () => collected,
    stop: () => {
      page.off("request", onRequest);
      page.off("response", onResponse);
      bucket.push(...collected);
    },
  };
}

function flattenObject(input, prefix = "", out = {}) {
  if (input === null || input === undefined) return out;
  if (typeof input !== "object") {
    if (prefix) out[prefix] = input;
    return out;
  }
  if (Array.isArray(input)) {
    input.forEach((val, idx) => {
      const key = prefix ? `${prefix}[${idx}]` : `[${idx}]`;
      flattenObject(val, key, out);
    });
    return out;
  }
  for (const [k, v] of Object.entries(input)) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === "object") flattenObject(v, key, out);
    else out[key] = v;
  }
  return out;
}

function extractRecordsFromJsonPayload(payload, sourceUrl) {
  const out = [];
  const seen = new Set();

  const scoreObject = (obj) => {
    const keys = Object.keys(obj || {}).map((k) => k.toLowerCase());
    let score = 0;
    if (keys.some((k) => /(name|nombre|empresa|consignee|shipper|importador|exportador|razon)/.test(k))) score += 3;
    if (keys.some((k) => /(country|pais|origen|destino)/.test(k))) score += 2;
    if (keys.some((k) => /(address|direccion|city|provincia|state)/.test(k))) score += 1;
    if (keys.some((k) => /(cuit|tax|vat|id)/.test(k))) score += 1;
    return score;
  };

  const visit = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      for (const item of node) visit(item);
      return;
    }
    if (typeof node !== "object") return;
    if (scoreObject(node) >= 2) {
      const flat = flattenObject(node);
      const hash = crypto.createHash("sha1").update(JSON.stringify(flat).slice(0, 5000)).digest("hex");
      if (!seen.has(hash)) {
        seen.add(hash);
        out.push({ _source: "api_json", _sourceUrl: sourceUrl, ...flat });
      }
    }
    for (const value of Object.values(node)) {
      if (typeof value === "object") visit(value);
    }
  };

  visit(payload);
  return out;
}

async function getFrameDiagnostics(frame) {
  return frame
    .evaluate((emptyRegexSource) => {
      const emptyRegex = new RegExp(emptyRegexSource, "i");
      const q = (sel) => document.querySelectorAll(sel).length;
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const buttonLikeSelector =
        'button, [role="button"], a, input[type="button"], input[type="submit"], .btn, div, span';

      const buttons = Array.from(document.querySelectorAll(buttonLikeSelector))
        .filter(visible)
        .map((el) => (el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim())
        .filter((t) => t.length > 0)
        .slice(0, 300);

      const visibleText = (document.body?.innerText || "").replace(/\s+/g, " ").trim().slice(0, 4000);
      const hasEmptyMessage = emptyRegex.test(visibleText);

      return {
        url: window.location.href,
        table: q("table"),
        tr: q("tr"),
        roleRow: q('[role="row"]'),
        agRow: q(".ag-row"),
        pDatatableRows: q(".p-datatable-tbody tr"),
        input: q("input"),
        select: q("select"),
        button: q("button"),
        visibleButtons: buttons,
        visibleText,
        hasEmptyMessage,
      };
    }, EMPTY_TEXT_REGEX.source)
    .catch(() => ({
      url: frame.url(),
      table: 0,
      tr: 0,
      roleRow: 0,
      agRow: 0,
      pDatatableRows: 0,
      input: 0,
      select: 0,
      button: 0,
      visibleButtons: [],
      visibleText: "",
      hasEmptyMessage: false,
      frameError: true,
    }));
}

function isFrameInfoLoginLike(frameInfo) {
  if (!frameInfo) return false;
  const url = String(frameInfo.url || "").toLowerCase();
  const visibleText = String(frameInfo.visibleText || "");
  const hasPasswordSignals = Number(frameInfo.input || 0) > 0 && /password|clave/i.test(visibleText);
  const isLoginUrl = /\/login(?:\/|$|\?)/i.test(url) || /signin|session|sesion|ingresar/.test(url);
  const hasLoginText = LOGIN_TEXT_REGEX.test(visibleText);
  return isLoginUrl || hasLoginText || hasPasswordSignals;
}

async function captureDeepDiagnostics(page, moduleDef, log) {
  const screenshotPath = path.join(DATA_DIR, `${moduleDef.debugBaseName}.png`);
  const htmlPath = path.join(DATA_DIR, `${moduleDef.debugBaseName}.html`);
  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) await fs.writeFile(htmlPath, html, "utf8");

  const frameInfos = [];
  const frames = page.frames();
  log("Iframes detectados", "info", { count: Math.max(frames.length - 1, 0) });

  for (let idx = 0; idx < frames.length; idx += 1) {
    const frame = frames[idx];
    const isMain = frame === page.mainFrame();
    const stats = await getFrameDiagnostics(frame);
    frameInfos.push({
      frameIndex: idx,
      frameType: isMain ? "main_page" : "iframe",
      name: frame.name() || "",
      url: frame.url(),
      ...stats,
    });
    log(`Diagnóstico ${isMain ? "main page" : "iframe"}`, "info", {
      frameIndex: idx,
      frameType: isMain ? "main_page" : "iframe",
      frameUrl: frame.url(),
      table: stats.table,
      tr: stats.tr,
      roleRow: stats.roleRow,
      agRow: stats.agRow,
      pDatatableRows: stats.pDatatableRows,
      hasEmptyMessage: stats.hasEmptyMessage,
      visibleButtons: stats.visibleButtons.slice(0, 50),
      visibleTextPreview: (stats.visibleText || "").slice(0, 500),
    });
  }

  const scoredCandidates = frameInfos
    .map((frameInfo) => ({
      ...frameInfo,
      score:
        Number(frameInfo.tr || 0) +
        Number(frameInfo.roleRow || 0) +
        Number(frameInfo.agRow || 0) +
        Number(frameInfo.pDatatableRows || 0),
      isLoginLike: isFrameInfoLoginLike(frameInfo),
    }))
    .filter((frameInfo) => !frameInfo.isLoginLike)
    .sort((a, b) => b.score - a.score);

  const candidate = scoredCandidates[0] || frameInfos.find((f) => f.frameType === "main_page") || frameInfos[0];

  log("Frame seleccionado para extracción", "info", {
    frameIndex: candidate?.frameIndex ?? 0,
    frameType: candidate?.frameType ?? "main_page",
    frameUrl: candidate?.url || page.url(),
    score: candidate?.score ?? 0,
    isLoginLike: Boolean(candidate && isFrameInfoLoginLike(candidate)),
  });

  return {
    screenshotPath,
    htmlPath,
    frameInfos,
    selectedFrameIndex: candidate?.frameIndex ?? 0,
    selectedFrameInfo: candidate || null,
  };
}

async function clickSearchLikeElementsInFrame(frame, log, labelScope) {
  const clicked = await frame
    .evaluate((regexSource) => {
      const regex = new RegExp(regexSource, "i");
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const candidates = Array.from(
        document.querySelectorAll(
          'button, [role="button"], a, input[type="button"], input[type="submit"], .btn, div, span'
        )
      )
        .filter(visible)
        .slice(0, 2000);

      const results = [];
      for (const el of candidates) {
        const txt = (el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim();
        if (!txt || !regex.test(txt)) continue;
        try {
          el.click();
          results.push(txt);
        } catch {
          // ignore
        }
      }
      return results.slice(0, 50);
    }, SEARCH_TEXT_REGEX.source)
    .catch(() => []);

  if (clicked.length > 0) {
    log("Clicks de búsqueda ejecutados", "info", { scope: labelScope, clicked });
  } else {
    log("No se detectaron elementos clickeables de búsqueda", "info", { scope: labelScope });
  }
}

async function clickSearchActions(page, frameInfos, log) {
  await clickSearchLikeElementsInFrame(page.mainFrame(), log, "main_page");
  const frames = page.frames();
  for (const fi of frameInfos.filter((f) => f.frameType === "iframe")) {
    const frame = frames[fi.frameIndex];
    if (!frame) continue;
    await clickSearchLikeElementsInFrame(frame, log, `iframe_${fi.frameIndex}`);
  }
  await waitForSettled(page);
  await sleep(600);
}

async function maximizeRowsPerPage(frame, log, scope) {
  const changed = await frame
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const selects = Array.from(document.querySelectorAll("select")).filter(visible);
      let changedCount = 0;
      for (const select of selects) {
        const label = (
          (select.getAttribute("aria-label") || "") +
          " " +
          (select.getAttribute("name") || "") +
          " " +
          (select.getAttribute("id") || "") +
          " " +
          (select.closest("label")?.textContent || "")
        )
          .replace(/\s+/g, " ")
          .trim();
        if (!/(rows|fila|por pagina|por p[aá]gina|cantidad|mostrar|registros|items|page size)/i.test(label)) {
          continue;
        }
        const options = Array.from(select.options || [])
          .map((opt) => ({
            value: opt.value,
            text: (opt.textContent || "").trim(),
            num: Number(String(opt.value || opt.textContent || "").replace(/[^\d]/g, "")),
          }))
          .filter((o) => Number.isFinite(o.num) && o.num > 0)
          .sort((a, b) => a.num - b.num);
        if (!options.length) continue;
        const max = options[options.length - 1];
        select.value = max.value;
        select.dispatchEvent(new Event("change", { bubbles: true }));
        select.dispatchEvent(new Event("input", { bubbles: true }));
        changedCount += 1;
      }
      return changedCount;
    })
    .catch(() => 0);
  log("Selector de filas por página", "info", { scope, changed });
}

async function detectFilterCombos(frame, log, scope) {
  const filters = await frame
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const normalize = (v) => (v || "").replace(/\s+/g, " ").trim();
      const selects = Array.from(document.querySelectorAll("select")).filter(visible);
      return selects.map((el, idx) => ({
        index: idx,
        label: normalize(
          el.getAttribute("aria-label") ||
            el.getAttribute("name") ||
            el.getAttribute("id") ||
            el.closest("label")?.textContent ||
            ""
        ),
        options: Array.from(el.options || []).map((o) => ({
          value: o.value,
          text: normalize(o.textContent),
        })),
      }));
    })
    .catch(() => []);

  const eligible = filters
    .map((f) => ({
      ...f,
      options: (f.options || []).filter((o) => {
        const txt = (o.text || "").toLowerCase();
        return isTruthy(o.value) || (isTruthy(o.text) && !/todos|all|seleccione|select|--/.test(txt));
      }),
    }))
    .filter((f) => (f.options || []).length > 1)
    .filter((f) => /(pais|country|period|fecha|anio|año|mes|origen|destino)/i.test(f.label))
    .slice(0, 3);

  if (eligible.length === 0) {
    log("No se detectaron filtros útiles", "info", { scope });
    return [{ selects: [] }];
  }

  let combos = [{ selects: [] }];
  for (const f of eligible) {
    const next = [];
    for (const combo of combos) {
      for (const opt of f.options.slice(0, 100)) {
        next.push({
          selects: [
            ...combo.selects,
            { index: f.index, label: f.label, value: opt.value || opt.text, text: opt.text || opt.value },
          ],
        });
        if (next.length >= CONFIG.maxFilterCombos) break;
      }
      if (next.length >= CONFIG.maxFilterCombos) break;
    }
    combos = next.length ? next : combos;
    if (combos.length >= CONFIG.maxFilterCombos) break;
  }
  log("Combinaciones de filtros detectadas", "info", { scope, totalCombos: combos.length });
  return combos.length ? combos : [{ selects: [] }];
}

async function applyFilterCombo(frame, combo) {
  await frame
    .evaluate((comboArg) => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const selects = Array.from(document.querySelectorAll("select")).filter(visible);
      for (const sel of comboArg.selects || []) {
        const target = selects[sel.index];
        if (!target) continue;
        const byValue = Array.from(target.options).find((o) => String(o.value) === String(sel.value));
        const byText = Array.from(target.options).find(
          (o) => String(o.textContent || "").trim() === String(sel.text).trim()
        );
        const finalOption = byValue || byText;
        if (!finalOption) continue;
        target.value = finalOption.value;
        target.dispatchEvent(new Event("change", { bubbles: true }));
        target.dispatchEvent(new Event("input", { bubbles: true }));
      }
    }, combo)
    .catch(() => {});
}

async function detectNativeExportInFrame(frame) {
  return frame
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const candidates = Array.from(
        document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], .btn')
      ).filter(visible);
      for (const el of candidates) {
        const txt = (el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim();
        if (/(export|excel|csv|descargar|download|reporte)/i.test(txt)) {
          return true;
        }
      }
      return false;
    })
    .catch(() => false);
}

async function runNativeExport(page, frame, moduleDef, comboTag, log) {
  const found = await detectNativeExportInFrame(frame);
  if (!found) return null;
  const clicked = await frame
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const candidates = Array.from(
        document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], .btn')
      ).filter(visible);
      for (const el of candidates) {
        const txt = (el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim();
        if (/(export|excel|csv|descargar|download|reporte)/i.test(txt)) {
          el.click();
          return txt;
        }
      }
      return "";
    })
    .catch(() => "");
  if (!clicked) return null;

  const download = await page.waitForEvent("download", { timeout: 25000 }).catch(() => null);
  if (!download) return null;
  const filePrefix = `${moduleDef.key}_native_${safeSlug(comboTag || "base")}`;
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
  if (ext === ".xlsx") await workbook.xlsx.readFile(filePath);
  else await workbook.csv.readFile(filePath);
  const ws = workbook.worksheets[0];
  if (!ws) return [];
  const headers = ws
    .getRow(1)
    .values.slice(1)
    .map((h, i) => (isTruthy(h) ? String(h).trim() : `col_${i + 1}`));
  const rows = [];
  ws.eachRow((row, idx) => {
    if (idx === 1) return;
    const vals = row.values.slice(1);
    if (!vals.some((v) => isTruthy(v) || typeof v === "number")) return;
    const out = {};
    for (let i = 0; i < headers.length; i += 1) out[headers[i]] = vals[i] ?? "";
    out._source = "native_export";
    out._module = moduleDef.key;
    rows.push(out);
  });
  return rows;
}

async function extractDomRowsFromFrame(frame, moduleDef, comboLabel, pageNo, frameScope) {
  return frame
    .evaluate(
      ({ moduleKey, companyType, comboLabelArg, pageNoArg, frameScopeArg }) => {
        const norm = (v) => (v || "").replace(/\s+/g, " ").trim();
        const visible = (el) => {
          const st = window.getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
        };
        const rows = [];
        const push = (raw, source) => {
          const out = {
            _source: source,
            _module: moduleKey,
            _companyType: companyType,
            _filter: comboLabelArg,
            _page: pageNoArg,
            _frameScope: frameScopeArg,
          };
          for (const [k, v] of Object.entries(raw)) out[norm(k) || "col"] = norm(v);
          rows.push(out);
        };

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
            const a = tr.querySelector("a[href]");
            if (a) raw._detailHref = a.getAttribute("href") || "";
            push(raw, "dom_table");
          });
        });

        const roleRows = Array.from(document.querySelectorAll('[role="row"]')).filter(visible);
        if (roleRows.length > 1) {
          const headers = Array.from(roleRows[0].querySelectorAll('[role="columnheader"], .ag-header-cell-text')).map(
            (h, i) => norm(h.textContent) || `col_${i + 1}`
          );
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
            const a = rr.querySelector("a[href]");
            if (a) raw._detailHref = a.getAttribute("href") || "";
            push(raw, "dom_role_grid");
          }
        }

        const agRows = Array.from(document.querySelectorAll(".ag-row")).filter(visible);
        agRows.forEach((row) => {
          const cells = Array.from(row.querySelectorAll(".ag-cell"));
          if (!cells.length) return;
          const raw = {};
          cells.forEach((c, idx) => (raw[`ag_col_${idx + 1}`] = norm(c.innerText || c.textContent)));
          push(raw, "dom_ag_grid");
        });

        const pRows = Array.from(document.querySelectorAll(".p-datatable-tbody tr")).filter(visible);
        pRows.forEach((tr) => {
          const tds = Array.from(tr.querySelectorAll("td"));
          if (!tds.length) return;
          const raw = {};
          tds.forEach((td, idx) => (raw[`p_col_${idx + 1}`] = norm(td.innerText || td.textContent)));
          push(raw, "dom_primeng");
        });

        return rows;
      },
      {
        moduleKey: moduleDef.key,
        companyType: moduleDef.companyType,
        comboLabelArg: comboLabel,
        pageNoArg: pageNo,
        frameScopeArg: frameScope,
      }
    )
    .catch(() => []);
}

async function openDetailIfAnyAndExtract(context, row, log) {
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
        if (dd && dd.tagName.toLowerCase() === "dd") out[`detail_${norm(dt.textContent)}`] = norm(dd.textContent);
      });
      document.querySelectorAll("label").forEach((lb) => {
        const key = norm(lb.textContent);
        if (!key || out[`detail_${key}`]) return;
        const sib = lb.nextElementSibling;
        if (sib) out[`detail_${key}`] = norm(sib.textContent);
      });
      return out;
    });
    await p.close();
    return { ...row, ...detail, _detailUrl: absolute };
  } catch (error) {
    await p.close();
    log("No se pudo extraer detalle de registro", "warn", { href: absolute, error: summarizeError(error) });
    return row;
  }
}

async function advancePaginationInFrame(frame, log, scope) {
  const moved = await frame
    .evaluate(() => {
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const clickByRegex = (regex) => {
        const list = Array.from(
          document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], .btn')
        ).filter(visible);
        for (const el of list) {
          const txt = (el.innerText || el.textContent || el.value || "").replace(/\s+/g, " ").trim();
          if (regex.test(txt)) {
            el.click();
            return true;
          }
        }
        return false;
      };
      if (clickByRegex(/(cargar m[aá]s|mostrar m[aá]s|load more|ver m[aá]s)/i)) return "load_more";
      if (clickByRegex(/(siguiente|next|pr[oó]xima|›|»)/i)) return "next";
      return "";
    })
    .catch(() => "");

  if (moved) {
    log("Paginación detectada", "info", { scope, mode: moved });
    return true;
  }

  const before = await frame
    .evaluate(() => document.querySelectorAll("table tbody tr, .ag-row, .p-datatable-tbody tr, [role='row']").length)
    .catch(() => 0);
  await frame.page().mouse.wheel(0, 2200).catch(() => {});
  await sleep(700);
  const after = await frame
    .evaluate(() => document.querySelectorAll("table tbody tr, .ag-row, .p-datatable-tbody tr, [role='row']").length)
    .catch(() => before);
  if (after > before) {
    log("Infinite scroll detectado", "info", { scope, before, after });
    return true;
  }
  return false;
}

function getValueByKeyPatterns(raw, patterns) {
  for (const [k, v] of Object.entries(raw || {})) {
    const lk = k.toLowerCase();
    if (patterns.some((p) => p.test(lk)) && isTruthy(v)) return String(v).trim();
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
    if (!out.has(key)) out.set(key, r);
    else {
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

  metadataSheet.columns = [
    { header: "campo", key: "campo", width: 40 },
    { header: "valor", key: "valor", width: 120 },
  ];
  Object.entries(metadata).forEach(([k, v]) => {
    metadataSheet.addRow({ campo: k, valor: typeof v === "string" ? v : JSON.stringify(v) });
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

  const csvWb = new ExcelJS.Workbook();
  const csvWs = csvWb.addWorksheet("master");
  csvWs.columns = columns.map((c) => ({ header: c, key: c }));
  records.forEach((r) => csvWs.addRow(r));
  await csvWb.csv.writeFile(csvPath);

  return { jsonPath, xlsxPath, csvPath };
}

async function processModule(page, context, moduleDef, allLogs, networkBucket) {
  const log = makeLogger(moduleDef.key, allLogs);
  const moduleUrl = `${CONFIG.baseUrl}${moduleDef.urlPath}`;
  const rawRows = [];
  const normalizedRows = [];
  const usedMechanisms = new Set();
  const errors = [];

  log("Navegando módulo", "info", { moduleUrl });
  await page.goto(moduleUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
  await waitForSettled(page);

  let moduleState = await isLoginLikeState(page);
  if (moduleState.isLoginUrl || moduleState.isLoginLike) {
    log("Redirección a login detectada al entrar al módulo", "warn", {
      currentUrl: moduleState.url,
      visibleTextPreview: moduleState.visibleTextPreview,
    });
    await ensureLoggedIn(page, context, log);
    await page.goto(moduleUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await waitForSettled(page);
    moduleState = await isLoginLikeState(page);
    if (moduleState.isLoginUrl || moduleState.isLoginLike) {
      const msg = "No se pudo abrir módulo autenticado: sigue en login tras reintento";
      errors.push(msg);
      log(msg, "error", {
        currentUrl: moduleState.url,
        visibleTextPreview: moduleState.visibleTextPreview,
      });
      return {
        module: moduleDef.key,
        moduleUrl,
        rawRows,
        normalizedRows,
        diagnostics: null,
        usedMechanisms: [...usedMechanisms],
        errors,
      };
    }
  }

  const networkCollector = createNetworkCollector(page, log, networkBucket);
  await sleep(800);

  let diagnostics = await captureDeepDiagnostics(page, moduleDef, log);
  if (diagnostics.selectedFrameInfo && isFrameInfoLoginLike(diagnostics.selectedFrameInfo)) {
    log("Contexto seleccionado parece login, forzando reautenticación", "warn", {
      frameUrl: diagnostics.selectedFrameInfo.url,
      visibleTextPreview: (diagnostics.selectedFrameInfo.visibleText || "").slice(0, 400),
    });
    await ensureLoggedIn(page, context, log);
    await page.goto(moduleUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    await waitForSettled(page);
    diagnostics = await captureDeepDiagnostics(page, moduleDef, log);
    if (diagnostics.selectedFrameInfo && isFrameInfoLoginLike(diagnostics.selectedFrameInfo)) {
      const msg = "El contexto de extracción sigue siendo login luego de reautenticación";
      errors.push(msg);
      log(msg, "error", {
        frameUrl: diagnostics.selectedFrameInfo.url,
        visibleTextPreview: (diagnostics.selectedFrameInfo.visibleText || "").slice(0, 400),
      });
      return {
        module: moduleDef.key,
        moduleUrl,
        rawRows,
        normalizedRows,
        diagnostics,
        usedMechanisms: [...usedMechanisms],
        errors,
      };
    }
  }
  const frames = page.frames();
  const selectedFrame = frames[diagnostics.selectedFrameIndex] || page.mainFrame();
  const selectedScope =
    diagnostics.frameInfos.find((f) => f.frameIndex === diagnostics.selectedFrameIndex)?.frameType || "main_page";

  await clickSearchActions(page, diagnostics.frameInfos, log);
  await maximizeRowsPerPage(selectedFrame, log, selectedScope);
  await clickSearchActions(page, diagnostics.frameInfos, log);
  await waitForSettled(page);

  const filterCombos = await detectFilterCombos(selectedFrame, log, selectedScope);
  let nativeDownloadCount = 0;

  for (let comboIdx = 0; comboIdx < filterCombos.length; comboIdx += 1) {
    const combo = filterCombos[comboIdx];
    const comboLabel =
      (combo.selects || [])
        .map((s) => `${s.label || "filter"}=${s.text || s.value}`)
        .join(" | ") || "sin_filtro";
    log("Aplicando filtro", "info", { comboIdx: comboIdx + 1, comboLabel, scope: selectedScope });
    await applyFilterCombo(selectedFrame, combo);
    await clickSearchActions(page, diagnostics.frameInfos, log);
    await waitForSettled(page);

    const nativePath = await retry(
      () => runNativeExport(page, selectedFrame, moduleDef, `${comboIdx + 1}_${comboLabel}`, log),
      {
        retries: 2,
        delayMs: 1200,
        actionName: "runNativeExport",
        onRetry: (attempt, err) =>
          log("Reintento exportación nativa", "warn", { attempt, error: summarizeError(err) }),
      }
    ).catch(() => null);

    if (nativePath) {
      usedMechanisms.add("native_export");
      nativeDownloadCount += 1;
      const parsed = await parseNativeFileToRows(nativePath, moduleDef, log).catch((error) => {
        errors.push(`parse_native: ${summarizeError(error)}`);
        return [];
      });
      rawRows.push(...parsed);
    }

    if (!nativePath || rawRows.length === 0) {
      let pageNo = 1;
      let visited = 0;
      while (visited < CONFIG.maxPagesPerModule) {
        visited += 1;
        const domRows = await extractDomRowsFromFrame(selectedFrame, moduleDef, comboLabel, pageNo, selectedScope);
        if (!domRows.length) {
          log("La tabla/grilla no existe o no devuelve filas en este contexto", "warn", {
            scope: selectedScope,
            comboLabel,
            pageNo,
          });
        }
        for (const row of domRows) {
          const withDetail = await openDetailIfAnyAndExtract(context, row, log);
          rawRows.push(withDetail);
        }
        const moved = await advancePaginationInFrame(selectedFrame, log, selectedScope);
        if (!moved) break;
        pageNo += 1;
      }
      if (rawRows.length) usedMechanisms.add("dom_grid");
    }
  }

  const networkEvents = networkCollector.getAll();
  networkCollector.stop();
  const apiRows = networkEvents
    .filter((e) => e.type === "response" && e.hasJson && e.jsonPreview)
    .flatMap((e) => {
      try {
        const parsed = JSON.parse(e.jsonPreview);
        return extractRecordsFromJsonPayload(parsed, e.url);
      } catch {
        return [];
      }
    });

  if (apiRows.length > 0) {
    usedMechanisms.add("api_json");
    rawRows.push(...apiRows.map((r) => ({ ...r, _module: moduleDef.key, _frameScope: selectedScope })));
    log("Requests JSON con datos detectados", "info", { count: apiRows.length, scope: selectedScope });
  } else {
    log("No hubo requests JSON con datos útiles", "warn", { scope: selectedScope });
  }

  for (const raw of rawRows) {
    const norm = normalizeRecord(raw, moduleDef, moduleUrl);
    if (norm) normalizedRows.push(norm);
  }

  if (normalizedRows.length === 0) {
    const emptyFlags = diagnostics.frameInfos
      .map((f) => ({ frameIndex: f.frameIndex, frameType: f.frameType, hasEmptyMessage: f.hasEmptyMessage }))
      .filter((x) => x.hasEmptyMessage);
    if (emptyFlags.length > 0) {
      log("La UI muestra estado vacío", "warn", { emptyFlags });
    }
    errors.push("Sin registros normalizados tras agotar export nativa + API + DOM + iframes");
  }

  log("Resumen módulo", "info", {
    rawRows: rawRows.length,
    normalizedRows: normalizedRows.length,
    nativeDownloadCount,
    mechanisms: [...usedMechanisms],
    selectedScope,
  });

  return {
    module: moduleDef.key,
    moduleUrl,
    rawRows,
    normalizedRows,
    diagnostics,
    usedMechanisms: [...usedMechanisms],
    errors,
  };
}

async function writeNetworkDebugFile(networkBucket, logs) {
  const networkOut = {
    generatedAt: nowIso(),
    totalEvents: networkBucket.length,
    events: networkBucket.map((e) => ({
      ts: e.ts,
      type: e.type,
      method: e.method,
      url: e.url,
      status: e.status || null,
      resourceType: e.resourceType || "",
      contentType: e.contentType || "",
      relevantUrl: Boolean(e.relevantUrl),
      hasJson: Boolean(e.hasJson),
      parsedRecords: e.parsedRecords || 0,
      requestBodyPreview: (e.requestBodyPreview || "").slice(0, 1500),
      jsonPreview: (e.jsonPreview || "").slice(0, 4000),
    })),
    logs,
  };
  await fs.writeFile(NETWORK_DEBUG_PATH, JSON.stringify(networkOut, null, 2), "utf8");
}

async function main() {
  await ensureDirectories();
  const allLogs = [];
  const networkBucket = [];
  const log = makeLogger("main", allLogs);

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await createContextWithAuthState(browser, log);
  const page = await context.newPage();
  page.setDefaultTimeout(CONFIG.timeoutMs);

  const moduleResults = [];
  try {
    const initialUrl = `${CONFIG.baseUrl}${MODULES[0].urlPath}`;
    log("Inicio de ejecución", "info", { baseUrl: CONFIG.baseUrl, initialUrl });
    await page.goto(initialUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.timeoutMs });
    const authLog = makeLogger("auth", allLogs);
    await ensureLoggedIn(page, context, authLog);
    await waitForSettled(page);

    for (const moduleDef of MODULES) {
      const result = await processModule(page, context, moduleDef, allLogs, networkBucket);
      moduleResults.push(result);
    }
  } catch (error) {
    log("Fallo fatal durante ejecución", "error", { error: summarizeError(error) });
    await page
      .screenshot({ path: path.join(DATA_DIR, `fatal_${safeSlug(nowIso())}.png`), fullPage: true })
      .catch(() => {});
  } finally {
    await browser.close();
  }

  await writeNetworkDebugFile(networkBucket, allLogs);
  log("network-debug.json generado", "info", { path: NETWORK_DEBUG_PATH, events: networkBucket.length });

  const allNormalized = moduleResults.flatMap((m) => m.normalizedRows || []);
  const deduped = dedupeNormalized(allNormalized);

  const metadata = {
    extractedAt: nowIso(),
    baseUrl: CONFIG.baseUrl,
    totalModules: MODULES.length,
    totalRawRows: moduleResults.reduce((acc, m) => acc + (m.rawRows?.length || 0), 0),
    totalNormalizedRows: allNormalized.length,
    totalDedupedRows: deduped.length,
    mechanismsByModule: moduleResults.map((m) => ({ module: m.module, mechanisms: m.usedMechanisms })),
    moduleErrors: moduleResults.map((m) => ({ module: m.module, errors: m.errors })),
    diagnostics: moduleResults.map((m) => ({
      module: m.module,
      screenshot: m.diagnostics?.screenshotPath || "",
      html: m.diagnostics?.htmlPath || "",
      selectedFrameIndex: m.diagnostics?.selectedFrameIndex ?? 0,
      frameInfos: m.diagnostics?.frameInfos || [],
    })),
    networkDebugFile: NETWORK_DEBUG_PATH,
  };

  if (deduped.length === 0) {
    log("Sin filas en resultado final. Se prioriza diagnóstico profundo; no se genera master.", "warn", {
      networkDebugFile: NETWORK_DEBUG_PATH,
      diagnostics: moduleResults.map((m) => ({
        module: m.module,
        html: m.diagnostics?.htmlPath || "",
        screenshot: m.diagnostics?.screenshotPath || "",
      })),
    });
    await fs.writeFile(path.join(DATA_DIR, "diagnostic-summary.json"), JSON.stringify(metadata, null, 2), "utf8");
    return;
  }

  const outputs = await writeMasterOutputs(deduped, metadata, allLogs);
  log("Proceso finalizado con datos", "info", { totalDedupedRows: deduped.length, outputs });
}

main().catch((error) => {
  console.error(`[fatal] ${summarizeError(error)}`);
  process.exit(1);
});

