import { TABLES, PREVIEW } from "./schema.js";
import {
  appendRow,
  describeError,
  diagnostics,
  focusRow,
  formatDate,
  loadLists,
  readLastRows,
  serialToISO,
  updateRow,
} from "./excel.js";

/* ------------------------------------------------------------------ *
 * Estado
 * ------------------------------------------------------------------ */

const state = {
  lists: null,
  ready: false,
  user: "",
  consulta: { entrada: [], saida: [], loaded: false, loading: false },
  /** { tableKey, row, protocol, original } quando uma linha está em edição. */
  editing: null,
};

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

/* ------------------------------------------------------------------ *
 * Utilidades de tela
 * ------------------------------------------------------------------ */

let toastTimer = null;
function toast(message, kind = "") {
  const el = $("#toast");
  if (!el) return;
  el.textContent = message;
  el.className = "toast " + kind;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    el.hidden = true;
  }, kind === "error" ? 7000 : 3500);
}

function todayISO() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function normalise(text) {
  return String(text ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function cellText(row, col) {
  const v = row[col];
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return v;
  return String(v).trim();
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[ch]));
}

/* ------------------------------------------------------------------ *
 * Formulários
 * ------------------------------------------------------------------ */

function buildForm(form, table) {
  form.innerHTML = "";

  for (const field of table.fields) {
    const wrapper = document.createElement("div");
    wrapper.className = "field" + (field.span === 2 ? " span-2" : "");

    const label = document.createElement("label");
    label.htmlFor = `${table.key}-${field.key}`;
    label.innerHTML = `${escapeHtml(field.label)}${field.required ? ' <span class="req">*</span>' : ""}`;
    wrapper.appendChild(label);

    let control;

    if (field.type === "select") {
      control = document.createElement("select");
      control.className = "input";
      const blank = document.createElement("option");
      blank.value = "";
      blank.textContent = "— selecione —";
      control.appendChild(blank);

      const items = (state.lists && state.lists[field.list]) || [];
      for (const item of items) {
        const opt = document.createElement("option");
        opt.value = item;
        opt.textContent = item;
        control.appendChild(opt);
      }
      if (items.length === 0) {
        const opt = document.createElement("option");
        opt.value = "";
        opt.textContent = "(lista não encontrada)";
        opt.disabled = true;
        control.appendChild(opt);
      }
    } else if (field.type === "textarea") {
      control = document.createElement("textarea");
      control.className = "input";
      control.rows = 3;
    } else {
      control = document.createElement("input");
      control.className = "input";
      control.type = field.type === "date" ? "date" : "text";
      if (field.type === "date") control.value = todayISO();
    }

    control.id = `${table.key}-${field.key}`;
    control.name = field.key;
    control.dataset.key = field.key;
    if (field.placeholder) control.placeholder = field.placeholder;
    if (field.required) control.required = true;

    if (field.list && field.type === "text") {
      const listId = `dl-${table.key}-${field.key}`;
      control.setAttribute("list", listId);
      const dl = document.createElement("datalist");
      dl.id = listId;
      for (const item of (state.lists && state.lists[field.list]) || []) {
        const opt = document.createElement("option");
        opt.value = item;
        dl.appendChild(opt);
      }
      wrapper.appendChild(control);
      wrapper.appendChild(dl);
    } else {
      wrapper.appendChild(control);
    }

    form.appendChild(wrapper);
  }
}

/** Lê todos os campos do formulário (inclusive os vazios). */
function collect(form, table) {
  const payload = {};
  for (const field of table.fields) {
    const el = form.querySelector(`[data-key="${field.key}"]`);
    payload[field.key] = el ? String(el.value).trim() : "";
  }
  return payload;
}

function resetDates(form, table) {
  for (const field of table.fields) {
    if (field.type !== "date") continue;
    const el = form.querySelector(`[data-key="${field.key}"]`);
    if (el) el.value = todayISO();
  }
}

/** Preenche o formulário com os valores de uma linha da planilha. */
function fillForm(form, table, row) {
  for (const field of table.fields) {
    const el = form.querySelector(`[data-key="${field.key}"]`);
    if (!el) continue;

    const raw = row[field.col];
    if (field.type === "date") {
      el.value = serialToISO(raw);
      continue;
    }

    const value = raw === null || raw === undefined ? "" : String(raw).trim();

    if (field.type === "select") {
      let found = Array.from(el.options).some((opt) => opt.value === value);
      if (!found && value) {
        const opt = document.createElement("option");
        opt.value = value;
        opt.textContent = value;
        el.appendChild(opt);
        found = true;
      }
      el.value = found ? value : "";
    } else {
      el.value = value;
    }
  }
}

function showResult(container, { ok, table, protocol, row, message, attempt, wasEdit, warnings }) {
  container.hidden = false;
  container.className = "result" + (ok ? "" : " error");

  if (!ok) {
    container.innerHTML = `<h3>Não foi possível gravar</h3><p>${escapeHtml(message)}</p>`;
    return;
  }

  const user = state.user ? `<div class="kv"><span>Lançado por</span><b>${escapeHtml(state.user)}</b></div>` : "";
  const retry = attempt > 1 ? `<p class="hint">Linha reservada na tentativa ${attempt}.</p>` : "";
  const warn = warnings && warnings.length
    ? `<p class="hint warn">${warnings.map((w) => escapeHtml(w)).join("<br>")}</p>`
    : "";

  container.innerHTML = `
    <h3>${wasEdit ? "Alteração gravada" : "Lançamento gravado"}</h3>
    <div class="kv"><span>Aba</span><b>${escapeHtml(table.sheet.trim())}</b></div>
    <div class="kv"><span>Linha</span><b>${row}</b></div>
    <div class="kv"><span>Protocolo</span><b>${escapeHtml(String(protocol ?? "—"))}</b></div>
    ${user}
    ${retry}
    ${warn}
    <div class="actions">
      <button type="button" class="ghost" data-goto="${row}">Ver na planilha</button>
    </div>`;
}

function invalidateConsulta() {
  state.consulta.loaded = false;
  if ($(".view.is-active") && $(".view.is-active").dataset.view === "consulta") {
    loadConsulta(true);
  }
}

function wireForm(form, table, resultBox) {
  form.addEventListener("submit", async (event) => {
    event.preventDefault();

    const payload = collect(form, table);
    const missing = table.fields.filter((f) => f.required && !payload[f.key]);
    if (missing.length > 0) {
      toast(`Preencha: ${missing.map((f) => f.label).join(", ")}`, "error");
      return;
    }

    const isEdit = Boolean(state.editing && state.editing.tableKey === table.key);
    const submit = form.querySelector('button[type="submit"]');

    const buttons = $$("button", document);
    const reenable = buttons.filter((b) => !b.disabled);
    reenable.forEach((b) => (b.disabled = true));
    if (submit) submit.textContent = "Gravando...";

    try {
      let res;

      if (isEdit) {
        const { row, original } = state.editing;
        res = await updateRow(table, row, payload, original);

        if (!res.ok && res.reason === "conflict") {
          const proceed = window.confirm(
            `${res.message}\n\nDeseja gravar por cima das alterações feitas por outra pessoa?`
          );
          if (proceed) res = await updateRow(table, row, payload, null);
        }
      } else {
        res = await appendRow(table, payload, (attempt, row) => {
          toast(`Linha ${row} estava em uso. Tentando a próxima (${attempt})...`);
        });
      }

      if (res.ok) {
        showResult(resultBox, {
          ok: true,
          table,
          protocol: res.protocol,
          row: res.row,
          attempt: res.attempt,
          wasEdit: isEdit,
          warnings: res.warnings,
        });
        toast(isEdit ? "Alteração gravada." : "Lançamento gravado.", "ok");

        if (isEdit) {
          clearForm(table);
          state.editing = null;
          syncEditingUI();
        } else {
          clearForm(table);
        }

        invalidateConsulta();
        resultBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
      } else {
        showResult(resultBox, { ok: false, table, message: res.message });
        toast("Não foi possível gravar.", "error");
      }
    } catch (err) {
      showResult(resultBox, { ok: false, table, message: describeError(err) });
      toast("Erro ao gravar.", "error");
    } finally {
      reenable.forEach((b) => (b.disabled = false));
      syncEditingUI();
    }
  });
}

/* ------------------------------------------------------------------ *
 * Edição de lançamento existente
 * ------------------------------------------------------------------ */

function syncEditingUI() {
  for (const table of Object.values(TABLES)) {
    const bar = $(`[data-editbar="${table.key}"]`);
    const btn = $(`#btn-${table.key}`);
    const active = Boolean(state.editing && state.editing.tableKey === table.key);

    if (bar) {
      bar.hidden = !active;
      bar.innerHTML = active
        ? `<span class="editbar-text">Editando a linha <b>${state.editing.row}</b> · Protocolo <b>${escapeHtml(String(state.editing.protocol ?? "—"))}</b></span>
           <button type="button" class="ghost" data-cancel-edit>Cancelar</button>`
        : "";
    }

    if (btn) {
      btn.textContent = active
        ? "Salvar alterações"
        : table.key === "entrada" ? "Gravar entrada" : "Gravar saída";
    }
  }
}

function startEditing(item) {
  const table = item.table;
  const form = $(`#form-${table.key}`);
  if (!form) return;

  fillForm(form, table, item.row);

  const original = {};
  for (const field of table.fields) {
    const value = item.row[field.col];
    original[field.key] = value === null || value === undefined ? "" : value;
  }

  state.editing = {
    tableKey: table.key,
    row: item.rowNumber,
    protocol: item.row["A"],
    original,
  };

  const result = $(`#res-${table.key}`);
  if (result) result.hidden = true;

  switchView(table.key);
  syncEditingUI();
  form.scrollIntoView({ behavior: "smooth", block: "start" });
  toast(`Editando a linha ${item.rowNumber}.`);
}

function exitEditing(resetForm = true) {
  const editing = state.editing;
  state.editing = null;

  if (editing && resetForm) {
    const table = TABLES[editing.tableKey];
    clearForm(table);
    const result = $(`#res-${table.key}`);
    if (result) result.hidden = true;
  }

  syncEditingUI();
}

function clearForm(table) {
  const form = $(`#form-${table.key}`);
  if (!form) return;
  form.reset();
  resetDates(form, table);
}

/* ------------------------------------------------------------------ *
 * Consulta
 * ------------------------------------------------------------------ */

function rowToItem(table, row) {
  const preview = PREVIEW[table.key];
  const get = (col) => cellText(row, col);

  const title = table.key === "entrada"
    ? (get("H") || "(sem tipo)")
    : (get("J") || get("I") || "(sem assunto)");

  const sub = table.key === "entrada"
    ? [get("M"), get("O")].filter(Boolean).join(" · ")
    : [get("K"), get("U")].filter(Boolean).join(" · ");

  const meta = [
    `Prot. ${get("A")}`,
    formatDate(row[table.dateCol]),
    table.key === "entrada" ? "" : get("I"),
  ].filter(Boolean);

  const haystack = normalise(
    preview.map((p) => get(p.col))
      .concat([get("I"), get("J"), get("K"), get("M"), get("O"), get("Q"), get("Y")])
      .join(" ")
  );

  return {
    table,
    row,
    rowNumber: row.__row,
    title,
    sub,
    meta: meta.join(" · "),
    haystack,
  };
}

function renderConsulta() {
  const q = normalise($("#q").value);
  const which = $("#q-sheet").value;
  const container = $("#consulta-lista");
  const resumo = $("#consulta-resumo");

  const tables = [];
  if (which === "ambas" || which === "entrada") tables.push(TABLES.entrada);
  if (which === "ambas" || which === "saida") tables.push(TABLES.saida);

  let items = [];
  for (const table of tables) {
    for (const row of state.consulta[table.key]) {
      items.push(rowToItem(table, row));
    }
  }

  if (q) items = items.filter((it) => it.haystack.includes(q));

  items.sort((a, b) => (b.rowNumber || 0) - (a.rowNumber || 0));

  const shown = items.slice(0, 120);

  resumo.textContent = state.consulta.loading
    ? "Carregando..."
    : `${items.length} lançamento(s) encontrado(s)${items.length > shown.length ? ` — exibindo ${shown.length}` : ""}`;

  container.innerHTML = "";

  if (!state.consulta.loading && shown.length === 0) {
    const empty = document.createElement("p");
    empty.className = "muted small";
    empty.textContent = state.consulta.loaded
      ? "Nenhum lançamento encontrado. Use o botão Atualizar para reler a planilha."
      : "Abra a aba Consulta para carregar os lançamentos.";
    container.appendChild(empty);
    return;
  }

  for (const item of shown) {
    const el = document.createElement("div");
    el.className = "item";
    el.innerHTML = `
      <div class="top"><span>${escapeHtml(item.meta)}</span><span>${escapeHtml(item.table.label)}</span></div>
      <div class="title">${escapeHtml(item.title)}</div>
      <div class="sub">${escapeHtml(item.sub || "—")}</div>
      <div class="item-actions">
        <button type="button" class="ghost small" data-view-row>Ver na planilha</button>
        <button type="button" class="primary small" data-edit-row>Editar</button>
      </div>`;

    el.querySelector("[data-view-row]").addEventListener("click", async () => {
      const ok = await focusRow(item.table, item.rowNumber);
      if (!ok) toast("Não foi possível abrir a linha.", "error");
    });

    el.querySelector("[data-edit-row]").addEventListener("click", () => startEditing(item));

    container.appendChild(el);
  }
}

async function loadConsulta(force = false) {
  if (state.consulta.loading) return;
  if (state.consulta.loaded && !force) {
    renderConsulta();
    return;
  }

  state.consulta.loading = true;
  renderConsulta();

  try {
    const [entrada, saida] = await Promise.all([
      readLastRows(TABLES.entrada, 400),
      readLastRows(TABLES.saida, 400),
    ]);
    state.consulta.entrada = entrada;
    state.consulta.saida = saida;
    state.consulta.loaded = true;
  } catch (err) {
    toast(describeError(err), "error");
  } finally {
    state.consulta.loading = false;
    renderConsulta();
  }
}

/* ------------------------------------------------------------------ *
 * Navegação
 * ------------------------------------------------------------------ */

function switchView(name) {
  $$(".tab").forEach((t) => t.classList.toggle("is-active", t.dataset.view === name));
  $$(".view").forEach((v) => v.classList.toggle("is-active", v.dataset.view === name));

  if (name === "consulta") loadConsulta();
  if (name === "sobre") renderDiagnostics();
}

let diagDone = false;
async function renderDiagnostics() {
  if (diagDone) return;
  const dl = $("#diag");
  dl.innerHTML = "<dt>Status</dt><dd>verificando...</dd>";
  try {
    const info = await diagnostics(TABLES);
    diagDone = true;

    const describe = (key, label) => {
      const f = info.found[key];
      if (!f || !f.present) return [label, "aba não encontrada"];
      const protect = f.protected ? "protegida" : "livre";
      const next = f.nextRow ? `próxima linha ${f.nextRow}` : "próxima linha ?";
      const limit = f.editableUntil ? `limite ${f.editableUntil}` : "";
      return [label, `${next} · ${protect} · ${limit}`];
    };

    const rows = [
      ["Arquivo", info.workbook],
      describe("entrada", "Entrada"),
      describe("saida", "Saída"),
      ["Listas", state.lists ? `${Object.keys(state.lists).length} carregadas` : "não carregadas"],
      ["Usuário", state.user || "não identificado"],
    ];

    dl.innerHTML = rows
      .map(([k, v]) => `<dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd>`)
      .join("");
  } catch (err) {
    dl.innerHTML = `<dt>Erro</dt><dd>${escapeHtml(describeError(err))}</dd>`;
  }
}

/* ------------------------------------------------------------------ *
 * Usuário
 * ------------------------------------------------------------------ */

function loadUser() {
  state.user = localStorage.getItem("protocolo.user") || "";
  $("#btn-user").textContent = state.user ? state.user.split(" ")[0] : "Entrar";
}

function askUser() {
  const name = window.prompt(
    "Seu nome (aparece só no painel, não é gravado na planilha):",
    state.user || ""
  );
  if (name === null) return;
  state.user = name.trim();
  if (state.user) localStorage.setItem("protocolo.user", state.user);
  else localStorage.removeItem("protocolo.user");
  loadUser();
  diagDone = false;
}

/* ------------------------------------------------------------------ *
 * Inicialização
 * ------------------------------------------------------------------ */

function wireStaticControls() {
  $("#btn-user").addEventListener("click", askUser);

  $$(".tab").forEach((t) =>
    t.addEventListener("click", () => switchView(t.dataset.view))
  );

  $("#q").addEventListener("input", () => renderConsulta());
  $("#q-sheet").addEventListener("change", () => loadConsulta(true));

  const refresh = $("#btn-refresh");
  if (refresh) refresh.addEventListener("click", () => loadConsulta(true));

  $$('button[type="reset"]').forEach((btn) =>
    btn.addEventListener("click", () => {
      if (state.editing) exitEditing(true);
    })
  );

  document.addEventListener("click", (event) => {
    const btn = event.target.closest("[data-cancel-edit]");
    if (btn) exitEditing(true);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && state.editing) {
      exitEditing(true);
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
      const active = $(".view.is-active form");
      if (active) active.requestSubmit();
    }
  });

  $("#res-entrada").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-goto]");
    if (btn) focusRow(TABLES.entrada, Number(btn.dataset.goto));
  });
  $("#res-saida").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-goto]");
    if (btn) focusRow(TABLES.saida, Number(btn.dataset.goto));
  });
}

async function start() {
  loadUser();
  wireStaticControls();

  const entradaForm = $("#form-entrada");
  const saidaForm = $("#form-saida");

  try {
    state.lists = await loadLists();
  } catch (err) {
    toast(describeError(err), "error");
    state.lists = {};
  }

  buildForm(entradaForm, TABLES.entrada);
  buildForm(saidaForm, TABLES.saida);

  wireForm(entradaForm, TABLES.entrada, $("#res-entrada"));
  wireForm(saidaForm, TABLES.saida, $("#res-saida"));

  syncEditingUI();
  state.ready = true;
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady((info) => {
    if (!info || info.host !== Office.HostType.Excel) {
      $("#brand-sub").textContent = "Fora do Excel";
      toast("Este painel só funciona dentro do Excel Online ou do Excel para desktop.", "error");
      return;
    }

    $("#brand-sub").textContent = `Subcorregedoria 5ª RPM · ${info.platform || "Excel"}`;
    start().catch((err) => {
      $("#brand-sub").textContent = "Falha ao iniciar";
      toast(describeError(err), "error");
    });
  });
} else {
  window.addEventListener("load", () => {
    $("#brand-sub").textContent = "Fora do Excel";
    toast("Abra este painel dentro do Excel Online.", "error");
  });
}

// Fallback: se o Office.js não carregar (página aberta fora do Excel).
window.addEventListener("load", () => {
  setTimeout(() => {
    if (!state.ready) {
      toast("Abra este painel dentro do Excel Online.", "error");
    }
  }, 2500);
});
