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

const STEP_DASHBOARD = path.join(DATA_DIR, "step-dashboard.png");
const STEP_MENU_OPEN = path.join(DATA_DIR, "step-argentina-menu-open.png");
const STEP_IMPORT_OPEN = path.join(DATA_DIR, "step-importaciones-detalladas-open.png");
const STEP_MENU_BEFORE_IMPORT_CLICK = path.join(
  DATA_DIR,
  "step-menu-before-click-importaciones-detalladas.png"
);
const IMPORT_CLICK_FAILED_HTML = path.join(DATA_DIR, "importaciones-detalladas-click-failed.html");
const IMPORT_CLICK_FAILED_PNG = path.join(DATA_DIR, "importaciones-detalladas-click-failed.png");

const LOGIN_TEXT_REGEX =
  /(user login|user password|forgot my password|enter|iniciar sesi[oó]n|ingresar|password)/i;

function nowIso() {
  return new Date().toISOString();
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function summarizeError(error) {
  if (!error) return "unknown_error";
  if (error instanceof Error) return `${error.name}: ${error.message}`;
  return String(error);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
    'ion-button:has-text("Enter")',
    'ion-button:has-text("Ingresar")',
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

async function findAndOpenArgentinaMenu(page, log) {
  const result = await page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
    const visible = (el) => {
      const st = window.getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
    };

    const viewportTopBand = window.innerHeight * 0.35;
    const candidates = [];

    const pushCandidate = (el, reason) => {
      if (!el || !visible(el)) return;
      const r = el.getBoundingClientRect();
      if (r.top > viewportTopBand) return;
      const text = normalize(el.innerText || el.textContent || el.getAttribute("aria-label") || "");
      const className = String(el.className || "").toLowerCase();
      const id = String(el.id || "").toLowerCase();
      const src = String(el.getAttribute("src") || "").toLowerCase();
      const styleBg = String(el.style?.backgroundImage || "").toLowerCase();
      candidates.push({
        el,
        reason,
        top: r.top,
        left: r.left,
        width: r.width,
        height: r.height,
        text,
        className,
        id,
        src,
        styleBg,
      });
    };

    // 1) img con arg/argentina/flag.
    for (const img of Array.from(document.querySelectorAll("img"))) {
      const src = String(img.getAttribute("src") || "").toLowerCase();
      const alt = String(img.getAttribute("alt") || "").toLowerCase();
      const title = String(img.getAttribute("title") || "").toLowerCase();
      if (/arg|argentina|bandera|flag/.test(`${src} ${alt} ${title}`)) {
        const clickable = img.closest('button, a, [role="button"], li, div, ion-item, ion-button') || img;
        pushCandidate(clickable, "img-argentina");
      }
    }

    // 2) nodos topbar con texto argentina.
    const textNodes = Array.from(document.querySelectorAll("button, a, [role='button'], li, div, span, ion-item, ion-button"));
    for (const node of textNodes) {
      const txt = normalize(node.innerText || node.textContent || "");
      if (txt === "argentina" || txt.includes("argentina")) {
        pushCandidate(node, "text-argentina");
      }
    }

    // 3) fallback: elementos del carrusel/top bar.
    const topNodes = Array.from(
      document.querySelectorAll(
        "header * , ion-header * , .topbar * , .toolbar * , .navbar * , [class*='flag'] , [class*='bandera']"
      )
    );
    for (const node of topNodes) {
      const className = String(node.className || "").toLowerCase();
      if (/flag|bandera|country|pais/.test(className)) {
        const clickable = node.closest('button, a, [role="button"], li, div, ion-item, ion-button') || node;
        pushCandidate(clickable, "topbar-flag-class");
      }
    }

    // Dedup por referencia.
    const unique = [];
    const seen = new Set();
    for (const c of candidates) {
      if (seen.has(c.el)) continue;
      seen.add(c.el);
      unique.push(c);
    }

    // Score heurístico: arriba, izquierda, señales argentina.
    unique.sort((a, b) => {
      const score = (x) => {
        let s = 0;
        if (x.reason === "img-argentina") s += 200;
        if (x.reason === "text-argentina") s += 120;
        if (/arg|argentina/.test(`${x.text} ${x.src} ${x.className} ${x.id}`)) s += 80;
        s += Math.max(0, 60 - x.top);
        s += Math.max(0, 60 - x.left);
        return s;
      };
      return score(b) - score(a);
    });

    const chosen = unique[0];
    if (!chosen) return { clicked: false, strategy: "", reason: "no-candidate" };

    // click normal o por bbox.
    let clicked = false;
    let strategy = "dom-click";
    try {
      chosen.el.click();
      clicked = true;
    } catch {
      strategy = "bbox-click";
      const r = chosen.el.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const ev = document.elementFromPoint(cx, cy);
      if (ev) {
        ev.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, clientX: cx, clientY: cy }));
        ev.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: cx, clientY: cy }));
        ev.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: cx, clientY: cy }));
        clicked = true;
      }
    }

    return {
      clicked,
      strategy,
      reason: chosen.reason,
      chosenMeta: {
        text: chosen.text,
        className: chosen.className.slice(0, 200),
        id: chosen.id,
        src: chosen.src.slice(0, 200),
        top: chosen.top,
        left: chosen.left,
        width: chosen.width,
        height: chosen.height,
      },
    };
  });

  if (!result.clicked) {
    throw new Error("No se pudo clickear la bandera/contenedor de Argentina en topbar");
  }

  log("bandera argentina encontrada", "info", {
    strategy: result.strategy,
    reason: result.reason,
    chosen: result.chosenMeta,
  });

  await sleep(500);
  await waitForSettled(page);
  return result;
}

async function waitArgentinaMenuOpened(page, log) {
  const opened = await page
    .waitForFunction(() => {
      const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
      const visible = (el) => {
        const st = window.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
      };
      const items = Array.from(document.querySelectorAll("button, a, [role='menuitem'], [role='button'], li, span, div"))
        .filter(visible)
        .map((el) => normalize(el.innerText || el.textContent || ""))
        .filter(Boolean);
      const hasImportaciones = items.some((t) => t.includes("importaciones"));
      const hasImportDet = items.some((t) => t.includes("importaciones detalladas"));
      const hasExportaciones = items.some((t) => t.includes("exportaciones"));
      return hasImportaciones && hasImportDet && hasExportaciones;
    }, { timeout: 10000 })
    .then(() => true)
    .catch(() => false);

  if (!opened) {
    throw new Error("No se abrió el menú desplegable de Argentina");
  }

  log("menú argentina abierto");
}

async function clickImportacionesDetalladas(page, log) {
  const clicked = await page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim().toLowerCase();
    const visible = (el) => {
      const st = window.getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
    };
    const nodes = Array.from(
      document.querySelectorAll("button, a, [role='menuitem'], [role='button'], li, span, div, ion-item, ion-button")
    ).filter(visible);
    const target = nodes.find((el) => normalize(el.innerText || el.textContent || "") === "importaciones detalladas");
    if (!target) return false;
    target.click();
    return true;
  });

  if (!clicked) {
    const fallback = page.getByText("Importaciones Detalladas", { exact: true }).first();
    if (await fallback.isVisible().catch(() => false)) {
      await fallback.click({ timeout: 7000 }).catch(() => {});
    } else {
      throw new Error('No se pudo clickear "Importaciones Detalladas" en el menú de Argentina');
    }
  }

  log("click en Importaciones Detalladas ejecutado");
  await waitForSettled(page);
}

async function waitImportacionesDetalladasLoaded(page, log) {
  const ok = await page
    .waitForFunction(() => {
      const txt = (document.body?.innerText || "").replace(/\s+/g, " ").trim().toLowerCase();
      const hasTitle = txt.includes("argentina - importaciones detalladas");
      const hasPanel = txt.includes("consulta por parámetros") || txt.includes("consulta por parametros");
      return hasTitle && hasPanel;
    }, { timeout: 15000 })
    .then(() => true)
    .catch(() => false);

  if (!ok) {
    throw new Error('No cargó la vista "Argentina - Importaciones Detalladas" con "Consulta por Parámetros"');
  }

  log("vista Importaciones Detalladas cargada");
}

async function runNavigationOnly() {
  await ensureDirectories();
  const log = makeLogger("nav");
  const authLog = makeLogger("auth");

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await createContextWithAuthState(browser, log);
  const page = await context.newPage();
  page.setDefaultTimeout(CONFIG.timeoutMs);

  try {
    await ensureLoggedIn(page, context, authLog);

    // Ir explícitamente a dashboard, no al módulo directo.
    await page.goto(`${CONFIG.baseUrl}/home/dashboard`, {
      waitUntil: "domcontentloaded",
      timeout: CONFIG.timeoutMs,
    });
    await waitForSettled(page);
    await page.screenshot({ path: STEP_DASHBOARD, fullPage: true }).catch(() => {});

    const state = await isLoginLikeState(page);
    if (state.isLoginUrl) {
      throw new Error("Luego del login se cayó nuevamente a /login");
    }

    const argentinaClick = await findAndOpenArgentinaMenu(page, log);
    await waitArgentinaMenuOpened(page, log);
    await page.screenshot({ path: STEP_MENU_OPEN, fullPage: true }).catch(() => {});

    await clickImportacionesDetalladas(page, log);
    await waitImportacionesDetalladasLoaded(page, log);
    await page.screenshot({ path: STEP_IMPORT_OPEN, fullPage: true }).catch(() => {});

    log("Navegación dashboard -> Argentina -> Importaciones Detalladas OK", "info", {
      stepDashboard: STEP_DASHBOARD,
      stepMenuOpen: STEP_MENU_OPEN,
      stepImportOpen: STEP_IMPORT_OPEN,
      argentinaFlagStrategy: argentinaClick.strategy,
      argentinaFlagReason: argentinaClick.reason,
      moduleClickStrategy: "texto exacto Importaciones Detalladas en menú abierto",
      currentUrl: page.url(),
    });
  } catch (error) {
    log("Fallo en navegación", "error", { error: summarizeError(error), currentUrl: page.url() });
    await page
      .screenshot({ path: path.join(DATA_DIR, "navigation-failed.png"), fullPage: true })
      .catch(() => {});
    throw error;
  } finally {
    await browser.close();
  }
}

runNavigationOnly().catch((error) => {
  console.error(`[fatal] ${summarizeError(error)}`);
  process.exit(1);
});

