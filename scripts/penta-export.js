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
const ARG_MENU_SCOPED_HTML = path.join(DATA_DIR, "argentina-menu-scoped.html");
const ARG_MENU_SCOPED_PNG = path.join(DATA_DIR, "argentina-menu-scoped.png");
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

async function detectArgentinaMenuScoped(page) {
  return page.evaluate(() => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const visible = (el) => {
      const st = window.getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
    };
    const bbox = (el) => {
      const r = el.getBoundingClientRect();
      return { left: r.left, top: r.top, width: r.width, height: r.height };
    };
    const expected = [
      "importaciones",
      "importaciones detalladas",
      "exportaciones",
      "exportaciones detalladas",
      "otras operaciones",
    ];
    const rowSelector =
      "button, a, [role='menuitem'], [role='button'], li, ion-item, mat-list-item, .menu-item, .dropdown-item, div";

    const candidateMenus = Array.from(
      document.querySelectorAll(
        "[role='menu'], .dropdown-menu, .menu, .mat-mdc-menu-panel, .cdk-overlay-pane, ion-popover, ion-list, .popover, .p-menu, .p-tieredmenu"
      )
    ).filter(visible);

    const scoreMenu = (menu) => {
      const box = bbox(menu);
      const rows = Array.from(menu.querySelectorAll(rowSelector))
        .filter(visible)
        .map((el) => normalize(el.innerText || el.textContent || ""))
        .filter((txt) => txt.length > 0 && txt.length < 80);
      const rowsUniq = [...new Set(rows)];
      const lowers = rowsUniq.map((t) => t.toLowerCase());
      const matches = expected.filter((e) => lowers.includes(e)).length;
      let score = matches * 100;
      if (rowsUniq.length >= 4 && rowsUniq.length <= 8) score += 40;
      if (box.width <= 600 && box.height <= 700) score += 20;
      if (box.width > window.innerWidth * 0.9 || box.height > window.innerHeight * 0.9) score -= 120;
      if (lowers.some((t) => /américa|africa|asia|europa/.test(t))) score -= 80;
      return { score, rowsUniq, box };
    };

    let best = null;
    for (const menu of candidateMenus) {
      const scored = scoreMenu(menu);
      if (!best || scored.score > best.score) {
        best = { menu, ...scored };
      }
    }
    if (!best || best.score < 120) {
      return {
        menuFound: false,
        menuBox: null,
        options: [],
        rows: [],
      };
    }

    const rowNodesRaw = Array.from(best.menu.querySelectorAll(rowSelector)).filter(visible);
    const rowNodes = rowNodesRaw.filter((el) => {
      const txt = normalize(el.innerText || el.textContent || "");
      const b = bbox(el);
      if (!txt || txt.length > 120) return false;
      return b.height >= 18 && b.height <= 80 && b.width >= 80;
    });

    rowNodes.sort((a, b) => bbox(a).top - bbox(b).top);
    const rowsData = rowNodes.map((row, index) => ({
      index: index + 1,
      text: normalize(row.innerText || row.textContent || "").slice(0, 200),
      tagName: (row.tagName || "").toLowerCase(),
      className: String(row.className || "").slice(0, 200),
      id: row.id || "",
      role: row.getAttribute("role") || "",
      boundingBox: bbox(row),
      outerHTML: (row.outerHTML || "").replace(/\s+/g, " ").trim().slice(0, 260),
    }));

    const targetByText =
      rowsData.find((r) => normalize(r.text).toLowerCase() === "importaciones detalladas") || null;
    const targetSecondRow = rowsData[1] || null;

    return {
      menuFound: true,
      menuBox: best.box,
      options: best.rowsUniq,
      rows: rowsData,
      targetByText,
      targetSecondRow,
    };
  });
}

async function waitArgentinaMenuOpened(page, log) {
  let scoped = null;
  const started = Date.now();
  while (Date.now() - started < 10000) {
    scoped = await detectArgentinaMenuScoped(page).catch(() => null);
    if (scoped?.menuFound && scoped.rows.length >= 4) break;
    await sleep(250);
  }

  if (!scoped?.menuFound) {
    throw new Error("No se detectó el contenedor real del menú desplegado de Argentina");
  }

  await page.screenshot({ path: ARG_MENU_SCOPED_PNG, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) await fs.writeFile(ARG_MENU_SCOPED_HTML, html, "utf8");

  log("menú argentina abierto", "info", {
    menuBox: scoped.menuBox,
    options: scoped.options,
    scopedRows: scoped.rows.length,
    artifacts: {
      html: ARG_MENU_SCOPED_HTML,
      png: ARG_MENU_SCOPED_PNG,
    },
  });

  return scoped;
}

async function clickImportacionesDetalladas(page, log, scopedMenu) {
  await page.screenshot({ path: STEP_MENU_BEFORE_IMPORT_CLICK, fullPage: true }).catch(() => {});

  const scoped = scopedMenu || (await detectArgentinaMenuScoped(page));
  if (!scoped?.menuFound) {
    throw new Error("No se encontró menú scopiado de Argentina para clickear opción");
  }

  log("opciones detectadas dentro del menú scopiado", "info", {
    optionTexts: scoped.rows.map((r) => r.text),
  });

  const target = scoped.targetByText || scoped.targetSecondRow;
  if (!target) {
    throw new Error("No se pudo determinar target en menú (ni texto exacto ni segunda fila)");
  }

  log("target real seleccionado", "info", {
    mode: scoped.targetByText ? "text-exact-match" : "second-row-fallback",
    target,
  });

  const clickResult = await page.evaluate(({ scopedArg, targetArg }) => {
    const normalize = (txt) => (txt || "").replace(/\s+/g, " ").trim();
    const visible = (el) => {
      const st = window.getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== "none" && st.visibility !== "hidden" && r.width > 0 && r.height > 0;
    };
    const bbox = (el) => {
      const r = el.getBoundingClientRect();
      return { left: r.left, top: r.top, width: r.width, height: r.height };
    };
    const close = (a, b, t) => Math.abs(a - b) <= t;
    const clickAt = (x, y) => {
      const node = document.elementFromPoint(x, y);
      if (!node) return false;
      node.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, clientX: x, clientY: y }));
      node.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: x, clientY: y }));
      node.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: x, clientY: y }));
      return true;
    };

    const rowSelector =
      "button, a, [role='menuitem'], [role='button'], li, ion-item, mat-list-item, .menu-item, .dropdown-item, div";
    const menuCandidates = Array.from(
      document.querySelectorAll(
        "[role='menu'], .dropdown-menu, .menu, .mat-mdc-menu-panel, .cdk-overlay-pane, ion-popover, ion-list, .popover, .p-menu, .p-tieredmenu"
      )
    ).filter(visible);
    const menu = menuCandidates.find((m) => {
      const b = bbox(m);
      return (
        close(b.left, scopedArg.menuBox.left, 20) &&
        close(b.top, scopedArg.menuBox.top, 20) &&
        close(b.width, scopedArg.menuBox.width, 40) &&
        close(b.height, scopedArg.menuBox.height, 60)
      );
    }) || menuCandidates[0];

    if (!menu) return { ok: false, strategy: "no-menu" };

    const rows = Array.from(menu.querySelectorAll(rowSelector)).filter(visible);
    const row = rows.find((el) => {
      const t = normalize(el.innerText || el.textContent || "");
      const b = bbox(el);
      return (
        t === targetArg.text &&
        close(b.top, targetArg.boundingBox.top, 10) &&
        close(b.left, targetArg.boundingBox.left, 18) &&
        close(b.height, targetArg.boundingBox.height, 12)
      );
    });

    if (row) {
      row.scrollIntoView({ block: "center", inline: "nearest" });
      try {
        row.click();
        return { ok: true, strategy: "locator-click" };
      } catch {
        const r = row.getBoundingClientRect();
        if (clickAt(r.left + r.width / 2, r.top + r.height / 2)) {
          return { ok: true, strategy: "bbox-center-click" };
        }
        const icon = row.querySelector(
          "ion-icon, i, svg, [class*='icon'], [class*='chevron'], [class*='arrow']"
        );
        if (icon && visible(icon)) {
          const ib = icon.getBoundingClientRect();
          if (clickAt(ib.left + ib.width / 2, ib.top + ib.height / 2)) {
            return { ok: true, strategy: "icon-click" };
          }
        }
      }
    }

    // Fallback posicional dentro del menú (segunda fila).
    const m = menu.getBoundingClientRect();
    const targetY = m.top + (m.height * 2.5) / 5;
    const targetX = m.left + 40;
    if (clickAt(targetX, targetY)) {
      return { ok: true, strategy: "menu-absolute-coordinate-click", targetX, targetY };
    }

    return { ok: false, strategy: "none" };
  }, { scopedArg: scoped, targetArg: target });

  log("método de click ejecutado", clickResult.ok ? "info" : "warn", clickResult);

  const loaded = await page
    .waitForFunction(() => {
      const txt = (document.body?.innerText || "").replace(/\s+/g, " ").trim().toLowerCase();
      return (
        txt.includes("argentina - importaciones detalladas") &&
        (txt.includes("consulta por parámetros") || txt.includes("consulta por parametros"))
      );
    }, { timeout: 10000 })
    .then(() => true)
    .catch(() => false);

  if (!loaded) {
    await page.screenshot({ path: IMPORT_CLICK_FAILED_PNG, fullPage: true }).catch(() => {});
    const html = await page.content().catch(() => "");
    if (html) await fs.writeFile(IMPORT_CLICK_FAILED_HTML, html, "utf8");
    throw new Error(
      `No cargó Importaciones Detalladas luego del click. Estrategia usada: ${clickResult.strategy || "none"}`
    );
  }

  log("click en Importaciones Detalladas ejecutado", "info", {
    clickMethod: clickResult.strategy || "unknown",
    finalUrl: page.url(),
  });

  await waitForSettled(page);
  return {
    clickStrategy: clickResult.strategy || "unknown",
    scopedOptions: scoped.rows.map((r) => r.text),
  };
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

    const importClick = await clickImportacionesDetalladas(page, log);
    await waitImportacionesDetalladasLoaded(page, log);
    await page.screenshot({ path: STEP_IMPORT_OPEN, fullPage: true }).catch(() => {});

    log("Navegación dashboard -> Argentina -> Importaciones Detalladas OK", "info", {
      stepDashboard: STEP_DASHBOARD,
      stepMenuOpen: STEP_MENU_OPEN,
      stepImportOpen: STEP_IMPORT_OPEN,
      argentinaFlagStrategy: argentinaClick.strategy,
      argentinaFlagReason: argentinaClick.reason,
      menuRowsDetected: importClick.rowCount,
      moduleClickStrategy: importClick.clickStrategy,
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

