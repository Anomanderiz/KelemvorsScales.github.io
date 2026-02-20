(() => {
  "use strict";

  const SAVE_KEYS = ["STR", "DEX", "CON", "INT", "WIS", "CHA"];
  const NOVA_METHODS = ["attack", "save_half", "auto_hit"];
  const STORAGE_KEY = "kelemvor_scales_web_state_v1";

  const DEFAULT_STATE = {
    party_table: [
      { Name: "Fighter", AC: 18, HP: 40, STR: 4, DEX: 2, CON: 3, INT: 0, WIS: 1, CHA: 0 },
      { Name: "Rogue", AC: 16, HP: 35, STR: 0, DEX: 5, CON: 2, INT: 1, WIS: 2, CHA: 1 },
      { Name: "Cleric", AC: 19, HP: 38, STR: 3, DEX: 0, CON: 3, INT: 1, WIS: 4, CHA: 2 },
      { Name: "Wizard", AC: 13, HP: 30, STR: 0, DEX: 3, CON: 2, INT: 5, WIS: 2, CHA: 1 },
    ],
    attacks_table: [
      { Name: "Bite", Type: "attack", "Attack bonus": 7, DC: 0, Save: "DEX", Damage: "2d10+5", "Uses/round": 1, "Melee?": true, "Enabled?": true },
      { Name: "Claw", Type: "attack", "Attack bonus": 7, DC: 0, Save: "DEX", Damage: "2d6+5", "Uses/round": 2, "Melee?": true, "Enabled?": true },
      { Name: "Fire Breath", Type: "save", "Attack bonus": 0, DC: 15, Save: "DEX", Damage: "8d6", "Uses/round": 1, "Melee?": false, "Enabled?": true },
    ],
    party_dpr_table: [{ Member: "Fighter", DPR: 15.0 }],
    party_nova_table: [{
      Member: "Fighter",
      "Nova DPR": 25.0,
      Method: "attack",
      "Atk Bonus": 8,
      "Roll Mode": "normal",
      "Target AC": 17,
      "Crit Ratio": 1.5,
      "Save DC": 16,
      "Target Save Bonus": 0,
      "Save Success Mult": 0.5,
      Uptime: 0.9,
    }],

    mode_select: "normal",
    spread_targets: 1,
    thp_expr: "1d6+4",

    lair_enabled: false,
    lair_avg: 6.0,
    lair_targets: 2,
    lair_every_n: 2,

    rech_enabled: false,
    recharge_text: "5-6",
    rech_avg: 22.0,
    rech_targets: 1,

    rider_mode: "none",
    rider_duration: 1,
    rider_melee_only: true,

    boss_hp: 150,
    boss_ac: 16,
    resist_factor: 1.0,
    boss_regen: 0.0,

    mc_rounds: 3,
    mc_trials: 10000,
    mc_show_hist: true,

    enc_trials: 10000,
    enc_max_rounds: 12,
    enc_use_nova: true,
    dpr_cv: 0.6,
    initiative_mode: "random",

    tune_target_median: 4.0,
    tune_tpk_cap: 0.05,
  };

  const PARTY_COLUMNS = [
    { key: "Name", label: "Name", type: "text", parser: (v) => String(v).trim() },
    { key: "AC", label: "AC", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 10)) },
    { key: "HP", label: "HP", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 1)) },
    { key: "STR", label: "STR", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "DEX", label: "DEX", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "CON", label: "CON", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "INT", label: "INT", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "WIS", label: "WIS", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "CHA", label: "CHA", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
  ];

  const DPR_COLUMNS = [
    { key: "Member", label: "Member", type: "text", readOnly: true },
    { key: "DPR", label: "DPR", type: "number", step: "0.1", min: "0", parser: (v) => Math.max(0, safeFloat(v, 0.0)) },
  ];

  const NOVA_COLUMNS = [
    { key: "Member", label: "Member", type: "text", readOnly: true },
    { key: "Nova DPR", label: "Nova DPR", type: "number", readOnly: true },
    {
      key: "Method",
      label: "Method",
      type: "select",
      options: NOVA_METHODS,
      parser: (v) => normalizeNovaMethod(v),
    },
    { key: "Atk Bonus", label: "Atk Bonus", type: "number", step: "1", parser: (v) => safeInt(v, 7) },
    {
      key: "Roll Mode",
      label: "Roll Mode",
      type: "select",
      options: ["normal", "adv", "dis"],
      parser: (v) => normalizeRollMode(v),
    },
    { key: "Target AC", label: "Target AC", type: "number", readOnly: true },
    { key: "Crit Ratio", label: "Crit Ratio", type: "number", step: "0.01", parser: (v) => Math.max(0, safeFloat(v, 1.5)) },
    { key: "Save DC", label: "Save DC", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 16)) },
    { key: "Target Save Bonus", label: "Save Bonus", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Save Success Mult", label: "Save Success Mult", type: "number", step: "0.01", min: "0", max: "1", parser: (v) => clamp(safeFloat(v, 0.5), 0, 1) },
    { key: "Uptime", label: "Uptime", type: "number", step: "0.01", min: "0", max: "1", parser: (v) => clamp(safeFloat(v, 0.85), 0, 1) },
  ];

  const ATTACK_COLUMNS = [
    { key: "Name", label: "Name", type: "text", parser: (v) => String(v).trim() },
    { key: "Type", label: "Type", type: "select", options: ["attack", "save"], parser: (v) => (String(v).toLowerCase() === "save" ? "save" : "attack") },
    { key: "Attack bonus", label: "Attack bonus", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "DC", label: "DC", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Save", label: "Save", type: "select", options: SAVE_KEYS, parser: (v) => (SAVE_KEYS.includes(String(v).toUpperCase()) ? String(v).toUpperCase() : "DEX") },
    { key: "Damage", label: "Damage", type: "text", parser: (v) => String(v).trim() || "0" },
    { key: "Uses/round", label: "Uses/round", type: "number", step: "1", min: "0", parser: (v) => Math.max(0, safeInt(v, 1)) },
    { key: "Melee?", label: "Melee?", type: "checkbox", parser: (v) => Boolean(v) },
    { key: "Enabled?", label: "Enabled?", type: "checkbox", parser: (v) => Boolean(v) },
  ];

  let state = normalizeState(loadStateFromStorage() || {});
  let charts = {
    det: null,
    mc: null,
    surv: null,
    ttk: null,
  };
  let statusTimer = null;
  let chartResizeObserver = null;
  let chartResizeRaf = 0;

  const els = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheElements();
    bindTabs();
    bindGlobalActions();
    bindButtons();
    bindOptionControls();
    initResizableWidgets();
    renderAll();
    setStatus("Ready.", 1800);
  }

  function initResizableWidgets() {
    if (typeof window.ResizeObserver !== "function") {
      return;
    }

    if (chartResizeObserver) {
      chartResizeObserver.disconnect();
    }

    const chartWraps = Array.from(document.querySelectorAll(".chart-wrap"));
    if (!chartWraps.length) {
      return;
    }

    chartResizeObserver = new window.ResizeObserver(() => {
      if (chartResizeRaf) {
        return;
      }

      chartResizeRaf = window.requestAnimationFrame(() => {
        chartResizeRaf = 0;
        for (const chart of Object.values(charts)) {
          if (chart) {
            chart.resize();
          }
        }
      });
    });

    chartWraps.forEach((wrap) => chartResizeObserver.observe(wrap));
  }

  function cacheElements() {
    const ids = [
      "btnSaveLocal", "btnExportJson", "inputImportJson", "statusBar",
      "btnAddPartyRow", "partyTable", "dprTable", "novaTable",
      "btnAddAttackRow", "attacksTable",
      "optModeSelect", "optSpreadTargets", "optThpExpr", "optBossHp", "optBossAc", "optResistFactor", "optBossRegen",
      "optLairEnabled", "optLairAvg", "optLairTargets", "optLairEveryN",
      "optRechEnabled", "optRechargeText", "optRechAvg", "optRechTargets",
      "optRiderMode", "optRiderDuration", "optRiderMeleeOnly",
      "btnComputeDet", "detTable", "detChart",
      "ttdModeManual", "ttdModeNova", "btnComputeTtd", "effTable",
      "lblIncoming", "lblRoundsExact", "lblRoundsCeil",
      "mcTarget", "mcRounds", "mcTrials", "mcShowHist", "btnRunMc", "mcChart",
      "lblMcMean", "lblMcP95", "lblMcP99",
      "encTrials", "encMaxRounds", "encDprCv", "encInitiative", "encUseNova",
      "tuneMedian", "tuneTpkCap", "btnRunEncounter", "btnAutoTune",
      "lblTtkMedian", "lblTtkP1090", "lblTpk", "lblDowns",
      "survChart", "ttkChart",
      "reportText",
    ];

    for (const id of ids) {
      els[id] = document.getElementById(id);
    }
  }

  function bindTabs() {
    const tabButtons = Array.from(document.querySelectorAll(".tab"));
    const panels = {
      party: document.getElementById("panel-party"),
      boss: document.getElementById("panel-boss"),
      det: document.getElementById("panel-det"),
      ttd: document.getElementById("panel-ttd"),
      mc: document.getElementById("panel-mc"),
      enc: document.getElementById("panel-enc"),
      report: document.getElementById("panel-report"),
    };

    for (const button of tabButtons) {
      button.addEventListener("click", () => {
        const tab = button.dataset.tab;
        tabButtons.forEach((b) => b.classList.toggle("active", b === button));
        Object.entries(panels).forEach(([k, panel]) => {
          panel.classList.toggle("active", k === tab);
        });
      });
    }
  }

  function bindGlobalActions() {
    els.btnSaveLocal.addEventListener("click", () => {
      persistState();
      setStatus("Saved to browser storage.", 2000);
    });

    els.btnExportJson.addEventListener("click", () => {
      exportJson();
    });

    els.inputImportJson.addEventListener("change", async (event) => {
      const file = event.target.files && event.target.files[0];
      if (!file) {
        return;
      }
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        state = normalizeState(parsed || {});
        persistState();
        renderAll();
        setStatus(`Loaded profile from ${file.name}.`, 2600);
      } catch (error) {
        alert("Failed to import JSON profile. Check file format.");
      } finally {
        event.target.value = "";
      }
    });
  }

  function bindButtons() {
    els.btnAddPartyRow.addEventListener("click", () => {
      state.party_table.push({ Name: "New PC", AC: 16, HP: 35, STR: 0, DEX: 0, CON: 0, INT: 0, WIS: 0, CHA: 0 });
      syncPartyDependentRows(state);
      persistState();
      renderPartySection();
      refreshMcTargets();
      refreshEffTableFromMode();
      refreshReport();
      setStatus("Added party member.", 1500);
    });

    els.btnAddAttackRow.addEventListener("click", () => {
      state.attacks_table.push({
        Name: "New Attack",
        Type: "attack",
        "Attack bonus": 6,
        DC: 0,
        Save: "DEX",
        Damage: "1d6+3",
        "Uses/round": 1,
        "Melee?": true,
        "Enabled?": true,
      });
      persistState();
      renderAttackSection();
      refreshReport();
      setStatus("Added attack row.", 1500);
    });

    els.btnComputeDet.addEventListener("click", () => {
      computeDeterministic();
    });

    els.btnComputeTtd.addEventListener("click", () => {
      refreshEffTableFromMode();
      setStatus("TTD refreshed.", 1500);
    });

    els.ttdModeManual.addEventListener("change", () => {
      refreshEffTableFromMode();
    });
    els.ttdModeNova.addEventListener("change", () => {
      refreshEffTableFromMode();
    });

    els.btnRunMc.addEventListener("click", () => {
      runSingleTargetMc();
    });

    els.btnRunEncounter.addEventListener("click", () => {
      runEncounterAndRender();
    });

    els.btnAutoTune.addEventListener("click", () => {
      autoTuneBossHp();
    });
  }

  function bindOptionControls() {
    bindControl("optModeSelect", "mode_select", (el) => String(el.value || "normal"));
    bindControl("optSpreadTargets", "spread_targets", (el) => Math.max(1, safeInt(el.value, 1)));
    bindControl("optThpExpr", "thp_expr", (el) => String(el.value || "0").trim() || "0");

    bindControl("optBossHp", "boss_hp", (el) => Math.max(1, safeInt(el.value, 150)), { refreshEff: true });
    bindControl("optBossAc", "boss_ac", (el) => Math.max(1, safeInt(el.value, 16)), { syncParty: true, renderParty: true, refreshEff: true });
    bindControl("optResistFactor", "resist_factor", (el) => Math.max(0.01, safeFloat(el.value, 1.0)), { refreshEff: true });
    bindControl("optBossRegen", "boss_regen", (el) => Math.max(0, safeFloat(el.value, 0.0)), { refreshEff: true });

    bindControl("optLairEnabled", "lair_enabled", (el) => Boolean(el.checked));
    bindControl("optLairAvg", "lair_avg", (el) => Math.max(0, safeFloat(el.value, 0.0)));
    bindControl("optLairTargets", "lair_targets", (el) => Math.max(1, safeInt(el.value, 1)));
    bindControl("optLairEveryN", "lair_every_n", (el) => Math.max(1, safeInt(el.value, 1)));

    bindControl("optRechEnabled", "rech_enabled", (el) => Boolean(el.checked));
    bindControl("optRechargeText", "recharge_text", (el) => String(el.value || "5-6").trim() || "5-6");
    bindControl("optRechAvg", "rech_avg", (el) => Math.max(0, safeFloat(el.value, 0.0)));
    bindControl("optRechTargets", "rech_targets", (el) => Math.max(1, safeInt(el.value, 1)));

    bindControl("optRiderMode", "rider_mode", (el) => String(el.value || "none"));
    bindControl("optRiderDuration", "rider_duration", (el) => Math.max(1, safeInt(el.value, 1)));
    bindControl("optRiderMeleeOnly", "rider_melee_only", (el) => Boolean(el.checked));

    bindControl("mcRounds", "mc_rounds", (el) => Math.max(1, safeInt(el.value, 3)));
    bindControl("mcTrials", "mc_trials", (el) => Math.max(1000, safeInt(el.value, 10000)));
    bindControl("mcShowHist", "mc_show_hist", (el) => Boolean(el.checked));

    bindControl("encTrials", "enc_trials", (el) => Math.max(1000, safeInt(el.value, 10000)));
    bindControl("encMaxRounds", "enc_max_rounds", (el) => Math.max(1, safeInt(el.value, 12)));
    bindControl("encDprCv", "dpr_cv", (el) => clamp(safeFloat(el.value, 0.6), 0.05, 2.0));
    bindControl("encInitiative", "initiative_mode", (el) => {
      const v = String(el.value || "random");
      return ["random", "party_first", "boss_first"].includes(v) ? v : "random";
    });
    bindControl("encUseNova", "enc_use_nova", (el) => Boolean(el.checked));

    bindControl("tuneMedian", "tune_target_median", (el) => clamp(safeFloat(el.value, 4.0), 1.0, 20.0));
    bindControl("tuneTpkCap", "tune_tpk_cap", (el) => clamp(safeFloat(el.value, 0.05), 0.0, 1.0));
  }

  function bindControl(id, key, parser, options = {}) {
    const el = els[id];
    if (!el) return;
    el.addEventListener("change", () => {
      state[key] = parser(el);
      if (options.syncParty) {
        syncPartyDependentRows(state);
      }
      if (options.renderParty) {
        renderPartySection();
        refreshMcTargets();
      }
      if (options.refreshEff) {
        refreshEffTableFromMode();
      }
      persistState();
      refreshReport();
    });
  }

  function renderAll() {
    syncPartyDependentRows(state);
    syncControlsFromState();
    renderPartySection();
    renderAttackSection();
    refreshMcTargets();
    refreshEffTableFromMode();
    refreshReport();
  }

  function syncControlsFromState() {
    setControlValue(els.optModeSelect, state.mode_select);
    setControlValue(els.optSpreadTargets, state.spread_targets);
    setControlValue(els.optThpExpr, state.thp_expr);

    setControlValue(els.optBossHp, state.boss_hp);
    setControlValue(els.optBossAc, state.boss_ac);
    setControlValue(els.optResistFactor, state.resist_factor);
    setControlValue(els.optBossRegen, state.boss_regen);

    setControlChecked(els.optLairEnabled, state.lair_enabled);
    setControlValue(els.optLairAvg, state.lair_avg);
    setControlValue(els.optLairTargets, state.lair_targets);
    setControlValue(els.optLairEveryN, state.lair_every_n);

    setControlChecked(els.optRechEnabled, state.rech_enabled);
    setControlValue(els.optRechargeText, state.recharge_text);
    setControlValue(els.optRechAvg, state.rech_avg);
    setControlValue(els.optRechTargets, state.rech_targets);

    setControlValue(els.optRiderMode, state.rider_mode);
    setControlValue(els.optRiderDuration, state.rider_duration);
    setControlChecked(els.optRiderMeleeOnly, state.rider_melee_only);

    setControlValue(els.mcRounds, state.mc_rounds);
    setControlValue(els.mcTrials, state.mc_trials);
    setControlChecked(els.mcShowHist, state.mc_show_hist);

    setControlValue(els.encTrials, state.enc_trials);
    setControlValue(els.encMaxRounds, state.enc_max_rounds);
    setControlValue(els.encDprCv, state.dpr_cv);
    setControlValue(els.encInitiative, state.initiative_mode);
    setControlChecked(els.encUseNova, state.enc_use_nova);

    setControlValue(els.tuneMedian, state.tune_target_median);
    setControlValue(els.tuneTpkCap, state.tune_tpk_cap);
  }

  function renderPartySection() {
    renderEditableTable({
      mount: els.partyTable,
      columns: PARTY_COLUMNS,
      rows: state.party_table,
      showRemove: true,
      onCellChange: (rowIndex, key, value) => {
        state.party_table[rowIndex][key] = value;
        state.party_table[rowIndex] = sanitizePartyRow(state.party_table[rowIndex]);
        if (key === "Name") {
          syncPartyDependentRows(state);
          renderPartySection();
          refreshMcTargets();
          refreshEffTableFromMode();
        }
        persistState();
        refreshReport();
      },
      onRemoveRow: (rowIndex) => {
        state.party_table.splice(rowIndex, 1);
        syncPartyDependentRows(state);
        persistState();
        renderPartySection();
        refreshMcTargets();
        refreshEffTableFromMode();
        refreshReport();
      },
      emptyMessage: "No party members yet.",
    });

    renderEditableTable({
      mount: els.dprTable,
      columns: DPR_COLUMNS,
      rows: state.party_dpr_table,
      showRemove: false,
      onCellChange: (rowIndex, key, value) => {
        state.party_dpr_table[rowIndex][key] = value;
        state.party_dpr_table[rowIndex] = sanitizeDprRow(state.party_dpr_table[rowIndex]);
        syncPartyDependentRows(state);
        persistState();
        renderPartySection();
        refreshEffTableFromMode();
        refreshReport();
      },
      emptyMessage: "Add party members to create DPR rows.",
    });

    renderEditableTable({
      mount: els.novaTable,
      columns: NOVA_COLUMNS,
      rows: state.party_nova_table,
      showRemove: false,
      onCellChange: (rowIndex, key, value) => {
        state.party_nova_table[rowIndex][key] = value;
        state.party_nova_table[rowIndex] = sanitizeNovaRow(state.party_nova_table[rowIndex]);
        syncPartyDependentRows(state);
        persistState();
        renderPartySection();
        refreshEffTableFromMode();
        refreshReport();
      },
      emptyMessage: "Add party members to configure nova conversion.",
    });
  }

  function renderAttackSection() {
    renderEditableTable({
      mount: els.attacksTable,
      columns: ATTACK_COLUMNS,
      rows: state.attacks_table,
      showRemove: true,
      onCellChange: (rowIndex, key, value) => {
        state.attacks_table[rowIndex][key] = value;
        state.attacks_table[rowIndex] = sanitizeAttackRow(state.attacks_table[rowIndex]);
        persistState();
        refreshReport();
      },
      onRemoveRow: (rowIndex) => {
        state.attacks_table.splice(rowIndex, 1);
        persistState();
        renderAttackSection();
        refreshReport();
      },
      emptyMessage: "No boss attacks configured.",
    });
  }

  function renderEditableTable({ mount, columns, rows, onCellChange, onRemoveRow, showRemove, emptyMessage }) {
    mount.innerHTML = "";
    if (!rows.length) {
      const empty = document.createElement("div");
      empty.className = "table-empty";
      empty.textContent = emptyMessage || "No rows.";
      mount.appendChild(empty);
      return;
    }

    const wrap = document.createElement("div");
    wrap.className = "table-wrap";
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const hrow = document.createElement("tr");

    for (const col of columns) {
      const th = document.createElement("th");
      th.textContent = col.label;
      hrow.appendChild(th);
    }

    if (showRemove) {
      const th = document.createElement("th");
      th.textContent = "Actions";
      hrow.appendChild(th);
    }

    thead.appendChild(hrow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");

      columns.forEach((col) => {
        const td = document.createElement("td");
        const input = createCellInput(col, row[col.key]);
        if (col.readOnly) {
          input.disabled = true;
        }
        input.addEventListener("change", () => {
          if (col.readOnly) return;
          const parsed = readCellInput(col, input);
          onCellChange(rowIndex, col.key, parsed, col);
        });
        td.appendChild(input);
        tr.appendChild(td);
      });

      if (showRemove) {
        const td = document.createElement("td");
        td.className = "row-actions";
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "btn btn-danger";
        removeBtn.textContent = "Remove";
        removeBtn.addEventListener("click", () => onRemoveRow(rowIndex));
        td.appendChild(removeBtn);
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrap.appendChild(table);
    mount.appendChild(wrap);
  }

  function createCellInput(col, value) {
    if (col.type === "select") {
      const select = document.createElement("select");
      for (const option of col.options || []) {
        const opt = document.createElement("option");
        opt.value = String(option);
        opt.textContent = String(option);
        select.appendChild(opt);
      }
      select.value = value == null ? "" : String(value);
      return select;
    }

    if (col.type === "checkbox") {
      const input = document.createElement("input");
      input.type = "checkbox";
      input.checked = Boolean(value);
      return input;
    }

    const input = document.createElement("input");
    input.type = col.type === "number" ? "number" : "text";
    if (col.step != null) input.step = col.step;
    if (col.min != null) input.min = col.min;
    if (col.max != null) input.max = col.max;
    input.value = value == null ? "" : String(value);
    return input;
  }

  function readCellInput(col, input) {
    if (col.type === "checkbox") {
      return Boolean(input.checked);
    }
    const raw = input.value;
    if (typeof col.parser === "function") {
      return col.parser(raw);
    }
    if (col.type === "number") {
      return safeFloat(raw, 0);
    }
    return String(raw);
  }

  function refreshMcTargets() {
    const names = uniquePartyNames(state.party_table);
    const current = els.mcTarget.value;
    els.mcTarget.innerHTML = "";

    if (!names.length) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "(No PCs)";
      els.mcTarget.appendChild(opt);
      return;
    }

    for (const name of names) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      els.mcTarget.appendChild(opt);
    }

    if (names.includes(current)) {
      els.mcTarget.value = current;
    }
  }

  function refreshEffTableFromMode() {
    const useNova = Boolean(els.ttdModeNova.checked);
    if (useNova) {
      const { total, rows } = buildNovaEffRows();
      renderResultTable(els.effTable, rows);
      updateTtdLabels(total);
      return;
    }

    const totalDpr = state.party_dpr_table.reduce((acc, row) => acc + safeFloat(row.DPR, 0), 0);
    renderResultTable(els.effTable, state.party_dpr_table.map((row) => ({ Member: row.Member, DPR: round2(safeFloat(row.DPR, 0)) })));
    updateTtdLabels(totalDpr);
  }

  function updateTtdLabels(totalDpr) {
    const resist = Math.max(1e-6, safeFloat(state.resist_factor, 1.0));
    const regen = Math.max(0, safeFloat(state.boss_regen, 0.0));
    const effective = Math.max(0, totalDpr / resist - regen);
    const hp = Math.max(1, safeFloat(state.boss_hp, 150));
    const exact = effective > 0 ? hp / effective : Number.POSITIVE_INFINITY;

    els.lblIncoming.textContent = `Effective Incoming DPR: ${effective.toFixed(2)}`;
    els.lblRoundsExact.textContent = `Exact Rounds to Zero: ${Number.isFinite(exact) ? exact.toFixed(2) : "inf"}`;
    els.lblRoundsCeil.textContent = `Boss Defeated In: ${Number.isFinite(exact) ? String(Math.ceil(exact)) : "inf"} rounds`;
  }

  function computeDeterministic() {
    const attacks = attacksEnabledFromTable(state.attacks_table);
    const party = state.party_table.filter((r) => String(r.Name || "").trim().length > 0);
    const lairDpr = lairPerTargetDpr(state, party.length || 1);
    const rechDpr = rechargePerTargetDpr(state, party.length || 1);
    const additiveDpr = lairDpr + rechDpr;
    const thpAvg = Math.max(0, averageDamage(state.thp_expr || "0"));
    const spread = Math.max(1, safeInt(state.spread_targets, 1));

    const rows = [];
    const chartLabels = [];
    const chartValues = [];

    for (const pc of party) {
      const baseDpr = perRoundDprVsPc(pc, state.mode_select || "normal", attacks);
      const total = baseDpr / spread + additiveDpr;
      const net = Math.max(0, total - thpAvg);
      const hp = Math.max(1, safeInt(pc.HP, 1));
      const exact = net > 0 ? hp / net : Number.POSITIVE_INFINITY;
      const ceil = Number.isFinite(exact) ? Math.ceil(exact) : null;

      rows.push({
        Name: pc.Name || "?",
        AC: safeInt(pc.AC, 10),
        HP: hp,
        "DPR (attacks)": round2(baseDpr / spread),
        "DPR (total)": round2(total),
        "Net DPR (after THP)": round2(net),
        "Rounds to 0 (exact)": Number.isFinite(exact) ? exact.toFixed(2) : "inf",
        "Rounds to 0 (ceil)": Number.isFinite(exact) ? Math.ceil(exact) : "inf",
      });

      chartLabels.push(pc.Name || "?");
      chartValues.push(ceil == null ? 0 : ceil);
    }

    renderResultTable(els.detTable, rows);

    drawChart("det", els.detChart, {
      type: "bar",
      data: {
        labels: chartLabels,
        datasets: [{
          label: "Rounds to 0 (Ceil)",
          data: chartValues,
          backgroundColor: "rgba(43, 196, 176, 0.74)",
          borderColor: "rgba(43, 196, 176, 1)",
          borderWidth: 1,
        }],
      },
      options: baseChartOptions({
        plugins: { title: { display: true, text: "Time-To-Zero per Party Member" } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: "Rounds" } },
        },
      }),
    });

    setStatus("Deterministic calculation complete.", 2200);
  }

  function runSingleTargetMc() {
    const targetName = String(els.mcTarget.value || "");
    const pc = state.party_table.find((row) => String(row.Name || "") === targetName);
    if (!pc) {
      alert("Select a valid party member target.");
      return;
    }

    setStatus("Running single-target Monte Carlo...", 0);
    const attacks = attacksEnabledFromTable(state.attacks_table);
    const totals = runMcSim(pc, attacks, state);

    const mean = meanOf(totals);
    const p95 = percentile(totals, 95);
    const p99 = percentile(totals, 99);

    els.lblMcMean.textContent = `Mean Total Damage: ${mean.toFixed(1)}`;
    els.lblMcP95.textContent = `95th Percentile: ${p95.toFixed(1)}`;
    els.lblMcP99.textContent = `99th Percentile: ${p99.toFixed(1)}`;

    if (state.mc_show_hist) {
      const hist = histogram(totals, 40);
      drawChart("mc", els.mcChart, {
        type: "bar",
        data: {
          labels: hist.labels,
          datasets: [{
            label: `Damage over ${state.mc_rounds} rounds`,
            data: hist.values,
            backgroundColor: "rgba(255, 122, 69, 0.70)",
            borderColor: "rgba(255, 122, 69, 1)",
            borderWidth: 1,
            barPercentage: 1,
            categoryPercentage: 1,
          }],
        },
        options: baseChartOptions({
          plugins: {
            title: { display: true, text: `Boss -> ${targetName} Damage Distribution` },
            legend: { display: false },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 10 } },
            y: { title: { display: true, text: "Frequency" }, beginAtZero: true },
          },
        }),
      });
    } else {
      clearChart("mc");
    }

    setStatus(`Monte Carlo complete (${state.mc_trials} trials).`, 3200);
  }

  function runEncounterAndRender() {
    setStatus("Running encounter simulation...", 0);
    const metrics = runEncounterMc();
    if (metrics.error) {
      alert(metrics.error);
      setStatus("Encounter simulation failed.", 2600);
      return;
    }

    const finite = metrics.finiteTtk;
    if (!finite.length) {
      alert("Boss never died within max rounds. Increase max rounds or lower boss durability.");
      setStatus("Encounter simulation returned no boss defeats.", 2600);
      return;
    }

    const p10 = percentile(finite, 10);
    const med = percentile(finite, 50);
    const p90 = percentile(finite, 90);
    const tpk = metrics.tpkProb;
    const downs = metrics.pcsDownAtVictory;

    els.lblTtkMedian.textContent = `Median TTK: ${med.toFixed(2)} rounds`;
    els.lblTtkP1090.textContent = `TTK p10-p90: ${p10.toFixed(2)} - ${p90.toFixed(2)} rounds`;
    els.lblTpk.textContent = `TPK Probability: ${(100 * tpk).toFixed(1)}%`;
    els.lblDowns.textContent = `PCs down at victory: mean ${meanOf(downs).toFixed(2)}, p90 ${percentile(downs, 90).toFixed(0)}`;

    drawChart("surv", els.survChart, {
      type: "line",
      data: {
        labels: metrics.times,
        datasets: [{
          label: "S(t): Boss alive",
          data: metrics.survivalCurve,
          stepped: true,
          borderColor: "rgba(43, 196, 176, 1)",
          backgroundColor: "rgba(43, 196, 176, 0.25)",
          fill: true,
          pointRadius: 0,
          tension: 0,
        }],
      },
      options: baseChartOptions({
        plugins: { title: { display: true, text: "Boss Survival Curve" } },
        scales: {
          y: { min: 0, max: 1, title: { display: true, text: "Probability" } },
          x: { title: { display: true, text: "Rounds" } },
        },
      }),
    });

    const hist = histogram(finite, 30);
    drawChart("ttk", els.ttkChart, {
      type: "bar",
      data: {
        labels: hist.labels,
        datasets: [{
          label: "TTK Frequency",
          data: hist.values,
          backgroundColor: "rgba(255, 122, 69, 0.72)",
          borderColor: "rgba(255, 122, 69, 1)",
          borderWidth: 1,
          barPercentage: 1,
          categoryPercentage: 1,
        }],
      },
      options: baseChartOptions({
        plugins: { title: { display: true, text: "TTK Distribution" }, legend: { display: false } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: "Frequency" } },
          x: { ticks: { maxTicksLimit: 10 } },
        },
      }),
    });

    setStatus(`Encounter simulation complete. Trials=${state.enc_trials}.`, 3800);
  }

  function autoTuneBossHp() {
    setStatus("Auto-tuning boss HP...", 0);

    const target = safeFloat(state.tune_target_median, 4.0);
    const tpkCap = safeFloat(state.tune_tpk_cap, 0.05);
    const originalHp = safeFloat(state.boss_hp, 150);
    const originalTrials = safeInt(state.enc_trials, 10000);

    const simulateWithHp = (hp, trials) => {
      const metrics = runEncounterMc({ boss_hp: Math.max(1, Math.round(hp)), enc_trials: Math.max(1000, Math.round(trials)) });
      if (metrics.error) {
        return { med: Number.POSITIVE_INFINITY, tpk: 1.0 };
      }
      if (!metrics.finiteTtk.length) {
        return { med: Number.POSITIVE_INFINITY, tpk: metrics.tpkProb };
      }
      return { med: percentile(metrics.finiteTtk, 50), tpk: metrics.tpkProb };
    };

    const quickTrials = Math.max(3000, Math.floor(originalTrials * 0.4));
    let low = 1.0;
    let high = Math.max(10.0, originalHp);

    const medLow = simulateWithHp(low, quickTrials).med;
    let medHigh = simulateWithHp(high, quickTrials).med;

    let attempts = 0;
    while (medHigh < target && attempts < 12) {
      high *= 2;
      medHigh = simulateWithHp(high, quickTrials).med;
      attempts += 1;
    }

    if (!(medLow <= target && target <= medHigh)) {
      alert("Could not bracket target median TTK by HP alone. Try adjusting boss or party parameters.");
      setStatus("Auto-tune could not bracket target.", 3200);
      return;
    }

    for (let i = 0; i < 16; i += 1) {
      const mid = 0.5 * (low + high);
      const midRes = simulateWithHp(mid, quickTrials);
      if (midRes.med >= target) {
        high = mid;
      } else {
        low = mid;
      }
      if (Math.abs(midRes.med - target) < 0.05) {
        break;
      }
    }

    let tunedHp = Math.max(1, Math.round(high));

    const quickAtTuned = simulateWithHp(tunedHp, quickTrials);
    let capAdjusted = false;

    if (quickAtTuned.tpk > tpkCap + 1e-9) {
      const atMin = simulateWithHp(1, quickTrials);
      if (atMin.tpk <= tpkCap + 1e-9) {
        let capLow = 1;
        let capHigh = tunedHp;
        while (capLow < capHigh) {
          const midHp = Math.floor((capLow + capHigh + 1) / 2);
          const mid = simulateWithHp(midHp, quickTrials);
          if (mid.tpk <= tpkCap + 1e-9) {
            capLow = midHp;
          } else {
            capHigh = midHp - 1;
          }
        }

        let bestHp = capLow;
        let bestErr = Number.POSITIVE_INFINITY;
        for (let hp = Math.max(1, capLow - 5); hp <= capLow; hp += 1) {
          const candidate = simulateWithHp(hp, quickTrials);
          if (candidate.tpk <= tpkCap + 1e-9) {
            const err = Math.abs(candidate.med - target);
            if (err < bestErr) {
              bestErr = err;
              bestHp = hp;
            }
          }
        }
        tunedHp = bestHp;
        capAdjusted = true;
      }
    }

    state.boss_hp = tunedHp;
    state.enc_trials = originalTrials;
    syncControlsFromState();
    persistState();
    refreshReport();

    runEncounterAndRender();

    const finalMetrics = runEncounterMc();
    if (!finalMetrics.error && finalMetrics.finiteTtk.length) {
      const finalMed = percentile(finalMetrics.finiteTtk, 50);
      const finalTpk = finalMetrics.tpkProb;

      if (finalTpk > tpkCap + 1e-9) {
        alert(`Auto-tuned HP=${tunedHp}, median~${finalMed.toFixed(2)}, but TPK=${(100 * finalTpk).toFixed(1)}% is above cap ${(100 * tpkCap).toFixed(1)}%. Lower boss DPR or add party sustain.`);
      } else if (capAdjusted) {
        alert(`Auto-tune applied TPK cap and set HP=${tunedHp}. Median~${finalMed.toFixed(2)} rounds, TPK ${(100 * finalTpk).toFixed(1)}%.`);
      }

      setStatus(`Auto-tune complete. Boss HP set to ${tunedHp}.`, 4200);
      return;
    }

    setStatus("Auto-tune completed with partial results.", 4200);
  }

  function runMcSim(pcRow, attacks, opts) {
    const trials = Math.max(1000, safeInt(opts.mc_trials, 10000));
    const rounds = Math.max(1, safeInt(opts.mc_rounds, 3));
    const ac = Math.max(1, safeInt(pcRow.AC, 10));
    const saveBonuses = {};
    for (const key of SAVE_KEYS) {
      saveBonuses[key] = safeInt(pcRow[key], 0);
    }
    const thpAvg = Math.max(0, averageDamage(opts.thp_expr || "0"));
    const spreadTargets = Math.max(1, safeInt(opts.spread_targets, 1));

    const totalDamage = new Array(trials).fill(0);
    const riderRemaining = new Array(trials).fill(0);

    for (let rnd = 1; rnd <= rounds; rnd += 1) {
      for (let i = 0; i < trials; i += 1) {
        let roundDamage = 0;
        let triggeredRider = false;

        let currentAc = ac;
        let currentMode = String(opts.mode_select || "normal");

        if (opts.rider_mode === "-2 AC next round" && riderRemaining[i] > 0) {
          currentAc = Math.max(1, currentAc - 2);
        }
        if (opts.rider_mode === "grant advantage on melee next round" && riderRemaining[i] > 0) {
          currentMode = "adv";
        }

        for (const atk of attacks) {
          if (atk.uses_per_round <= 0) continue;

          for (let use = 0; use < atk.uses_per_round; use += 1) {
            if (Math.random() >= 1 / spreadTargets) continue;

            if (atk.kind === "save") {
              const bonus = saveBonuses[String(atk.save_stat || "DEX").toUpperCase()] || 0;
              const roll = randomInt(1, 20);
              const success = roll === 20 || (roll !== 1 && roll + bonus >= atk.dc);
              const dmg = rollDamageOne(atk.damage_expr);
              roundDamage += success ? 0.5 * dmg : dmg;
              continue;
            }

            const r1 = randomInt(1, 20);
            const r2 = randomInt(1, 20);
            let roll = r1;
            if (currentMode === "adv") roll = Math.max(r1, r2);
            if (currentMode === "dis") roll = Math.min(r1, r2);

            const isCrit = roll === 20;
            const isHit = isCrit || (roll !== 1 && roll + atk.attack_bonus >= currentAc);

            if (isHit) {
              const dmg = isCrit ? rollDamageCrunchyCritOne(atk.damage_expr) : rollDamageOne(atk.damage_expr);
              roundDamage += dmg;
              if (atk.is_melee || !opts.rider_melee_only) {
                triggeredRider = true;
              }
            }
          }
        }

        if (opts.lair_enabled && rnd % Math.max(1, safeInt(opts.lair_every_n, 1)) === 0) {
          const pTarget = Math.min(1, safeFloat(opts.lair_targets, 1) / spreadTargets);
          if (Math.random() < pTarget) {
            roundDamage += gammaRng(Math.max(0, safeFloat(opts.lair_avg, 0)), 0.5);
          }
        }

        if (opts.rech_enabled) {
          const pRech = parseRecharge(opts.recharge_text || "5-6");
          const pTarget = Math.min(1, safeFloat(opts.rech_targets, 1) / spreadTargets);
          if (Math.random() < pRech && Math.random() < pTarget) {
            roundDamage += gammaRng(Math.max(0, safeFloat(opts.rech_avg, 0)), 0.5);
          }
        }

        totalDamage[i] += Math.max(0, roundDamage - thpAvg);

        riderRemaining[i] = Math.max(0, riderRemaining[i] - 1);
        if (triggeredRider) {
          riderRemaining[i] = Math.max(1, safeInt(opts.rider_duration, 1));
        }
      }
    }

    return totalDamage;
  }

  function runEncounterMc(overrides = {}) {
    const opts = { ...state, ...overrides };
    const party = state.party_table.filter((r) => String(r.Name || "").trim().length > 0);
    if (!party.length) {
      return { error: "Add at least one party member in Party & DPR." };
    }

    const attacks = attacksEnabledFromTable(state.attacks_table);
    if (!attacks.length && !opts.lair_enabled && !opts.rech_enabled) {
      return { error: "Enable at least one attack, lair action, or recharge power." };
    }

    const trials = Math.max(1, safeInt(opts.enc_trials, 10000));
    const maxRounds = Math.max(1, safeInt(opts.enc_max_rounds, 12));
    const resist = Math.max(1e-6, safeFloat(opts.resist_factor, 1.0));
    const regen = Math.max(0, safeFloat(opts.boss_regen, 0.0));
    const thpAvg = Math.max(0, averageDamage(opts.thp_expr || "0"));
    const spreadTargets = Math.max(1, safeInt(opts.spread_targets, 1));
    const initMode = String(opts.initiative_mode || "random");
    const useNova = Boolean(opts.enc_use_nova);
    const dprCv = clamp(safeFloat(opts.dpr_cv, 0.6), 0.05, 2.0);

    const effList = effPartyDprs(useNova);
    if (!effList.length) {
      return { error: "No DPR rows available. Fill DPR input table first." };
    }

    const partyNames = party.map((pc, i) => String(pc.Name || `PC${i + 1}`));
    const P = partyNames.length;
    const effByName = new Map(effList.map((r) => [String(r[0]), safeFloat(r[1], 0)]));
    const effMeans = partyNames.map((name) => effByName.get(name) || 0);

    const bossHp = new Array(trials).fill(Math.max(1, safeFloat(opts.boss_hp, 150)));
    const pcsHp = Array.from({ length: trials }, () => party.map((pc) => Math.max(1, safeFloat(pc.HP, 1))));
    const pcsAlive = Array.from({ length: trials }, () => party.map(() => true));

    const pcAc = party.map((pc) => Math.max(1, safeInt(pc.AC, 10)));
    const pcSaves = party.map((pc) => SAVE_KEYS.map((k) => safeInt(pc[k], 0)));
    const saveIndex = { STR: 0, DEX: 1, CON: 2, INT: 3, WIS: 4, CHA: 5 };

    const riderRem = Array.from({ length: trials }, () => new Array(P).fill(0));

    const ttk = new Array(trials).fill(Number.POSITIVE_INFINITY);
    const tpkFlags = new Array(trials).fill(false);
    const pcsDownAtVictory = new Array(trials).fill(0);

    const bossFirstFlags = buildBossFirstFlags(trials, initMode);

    for (let rnd = 1; rnd <= maxRounds; rnd += 1) {
      let anyOngoing = false;
      for (let t = 0; t < trials; t += 1) {
        if (!Number.isFinite(ttk[t])) {
          anyOngoing = true;
          break;
        }
      }
      if (!anyOngoing) break;

      for (let t = 0; t < trials; t += 1) {
        if (Number.isFinite(ttk[t])) continue;
        if (bossFirstFlags[t]) continue;

        let dmgParty = 0;
        for (let j = 0; j < P; j += 1) {
          if (!pcsAlive[t][j]) continue;
          dmgParty += gammaRng(effMeans[j], dprCv);
        }
        dmgParty = Math.max(0, dmgParty / resist - regen);

        bossHp[t] -= dmgParty;
        if (bossHp[t] <= 0 && !Number.isFinite(ttk[t])) {
          ttk[t] = rnd;
          pcsDownAtVictory[t] = pcsAlive[t].reduce((acc, alive) => acc + (alive ? 0 : 1), 0);
        }
      }

      for (let t = 0; t < trials; t += 1) {
        if (Number.isFinite(ttk[t])) continue;

        const currentAc = pcAc.slice();
        const currentMode = new Array(P).fill(String(opts.mode_select || "normal"));
        const riderMode = String(opts.rider_mode || "none");

        if (riderMode === "-2 AC next round") {
          for (let j = 0; j < P; j += 1) {
            if (riderRem[t][j] > 0) {
              currentAc[j] = Math.max(1, currentAc[j] - 2);
            }
          }
        }
        if (riderMode === "grant advantage on melee next round") {
          for (let j = 0; j < P; j += 1) {
            if (riderRem[t][j] > 0) {
              currentMode[j] = "adv";
            }
          }
        }

        const riderTrig = new Array(P).fill(false);
        const thpPool = new Array(P).fill(thpAvg);

        const applyDamage = (target, rawDamage) => {
          const blocked = Math.min(rawDamage, thpPool[target]);
          thpPool[target] -= blocked;
          const dealt = rawDamage - blocked;
          if (dealt <= 0) return;

          pcsHp[t][target] -= dealt;
          if (pcsHp[t][target] <= 0 && pcsAlive[t][target]) {
            pcsAlive[t][target] = false;
          }
        };

        let aliveNow = aliveIndicesForTrial(pcsAlive[t]);
        if (!aliveNow.length) {
          tpkFlags[t] = true;
          continue;
        }

        const poolSize = Math.min(spreadTargets, aliveNow.length);
        const roundPool = sampleWithoutReplacement(aliveNow, poolSize);

        for (const atk of attacks) {
          if (atk.uses_per_round <= 0) continue;

          for (let use = 0; use < atk.uses_per_round; use += 1) {
            let alivePool = roundPool.filter((idx) => pcsAlive[t][idx]);
            if (!alivePool.length) {
              alivePool = aliveIndicesForTrial(pcsAlive[t]);
              if (!alivePool.length) break;
            }

            const target = alivePool[randomInt(0, alivePool.length - 1)];
            if (!pcsAlive[t][target]) continue;

            let rawDealt = 0;

            if (atk.kind === "save") {
              const saveIdx = saveIndex[String(atk.save_stat || "DEX").toUpperCase()] ?? saveIndex.DEX;
              const bonus = pcSaves[target][saveIdx];
              const roll = randomInt(1, 20);
              const success = roll === 20 || (roll !== 1 && roll + bonus >= atk.dc);
              const dmg = rollDamageOne(atk.damage_expr);
              rawDealt = success ? 0.5 * dmg : dmg;
            } else {
              const r1 = randomInt(1, 20);
              const r2 = randomInt(1, 20);
              let r = r1;
              if (currentMode[target] === "adv") r = Math.max(r1, r2);
              if (currentMode[target] === "dis") r = Math.min(r1, r2);

              const isCrit = r === 20;
              const isHit = isCrit || (r !== 1 && r + atk.attack_bonus >= currentAc[target]);

              if (isHit) {
                rawDealt = isCrit ? rollDamageCrunchyCritOne(atk.damage_expr) : rollDamageOne(atk.damage_expr);
                if (atk.is_melee || !opts.rider_melee_only) {
                  riderTrig[target] = true;
                }
              }
            }

            applyDamage(target, Math.max(0, rawDealt));
          }
        }

        if (opts.lair_enabled && rnd % Math.max(1, safeInt(opts.lair_every_n, 1)) === 0) {
          aliveNow = aliveIndicesForTrial(pcsAlive[t]);
          if (aliveNow.length) {
            const L = Math.min(Math.max(1, safeInt(opts.lair_targets, 1)), aliveNow.length);
            const targets = sampleWithoutReplacement(aliveNow, L);
            for (const idx of targets) {
              applyDamage(idx, Math.max(0, gammaRng(Math.max(0, safeFloat(opts.lair_avg, 0)), 0.5)));
            }
          }
        }

        if (opts.rech_enabled) {
          const pRech = parseRecharge(opts.recharge_text || "5-6");
          if (Math.random() < pRech) {
            aliveNow = aliveIndicesForTrial(pcsAlive[t]);
            if (aliveNow.length) {
              const R = Math.min(Math.max(1, safeInt(opts.rech_targets, 1)), aliveNow.length);
              const targets = sampleWithoutReplacement(aliveNow, R);
              for (const idx of targets) {
                applyDamage(idx, Math.max(0, gammaRng(Math.max(0, safeFloat(opts.rech_avg, 0)), 0.5)));
              }
            }
          }
        }

        for (let j = 0; j < P; j += 1) {
          riderRem[t][j] = Math.max(0, riderRem[t][j] - 1);
          if (riderTrig[j]) {
            riderRem[t][j] = Math.max(1, safeInt(opts.rider_duration, 1));
          }
        }

        if (aliveIndicesForTrial(pcsAlive[t]).length === 0) {
          tpkFlags[t] = true;
        }
      }

      for (let t = 0; t < trials; t += 1) {
        if (Number.isFinite(ttk[t])) continue;
        if (!bossFirstFlags[t]) continue;

        let dmgParty = 0;
        for (let j = 0; j < P; j += 1) {
          if (!pcsAlive[t][j]) continue;
          dmgParty += gammaRng(effMeans[j], dprCv);
        }
        dmgParty = Math.max(0, dmgParty / resist - regen);

        bossHp[t] -= dmgParty;
        if (bossHp[t] <= 0 && !Number.isFinite(ttk[t])) {
          ttk[t] = rnd;
          pcsDownAtVictory[t] = pcsAlive[t].reduce((acc, alive) => acc + (alive ? 0 : 1), 0);
        }
      }
    }

    const times = [];
    for (let t = 0; t <= maxRounds; t += 1) times.push(t);
    const survivalCurve = times.map((round) => ttk.filter((v) => v > round).length / trials);

    const finiteTtk = ttk.filter((v) => Number.isFinite(v));
    const downFinite = [];
    for (let i = 0; i < ttk.length; i += 1) {
      if (Number.isFinite(ttk[i])) {
        downFinite.push(pcsDownAtVictory[i]);
      }
    }

    return {
      ttk,
      finiteTtk,
      tpkProb: meanOf(tpkFlags.map((x) => (x ? 1 : 0))),
      pcsDownAtVictory: downFinite,
      times,
      survivalCurve,
    };
  }

  function buildBossFirstFlags(n, mode) {
    if (mode === "boss_first") {
      return new Array(n).fill(true);
    }
    if (mode === "party_first") {
      return new Array(n).fill(false);
    }
    return Array.from({ length: n }, () => Math.random() < 0.5);
  }

  function aliveIndicesForTrial(aliveRow) {
    const out = [];
    for (let i = 0; i < aliveRow.length; i += 1) {
      if (aliveRow[i]) out.push(i);
    }
    return out;
  }

  function buildNovaEffRows() {
    let total = 0;
    const rows = [];

    for (const row of state.party_nova_table) {
      const conv = computeNovaConversion(row, state.boss_ac);
      total += conv.effDpr;

      rows.push({
        Member: row.Member || "?",
        Method: conv.method,
        "P(main)%": (100 * conv.pMain).toFixed(1),
        "P(crit)%": (100 * conv.pCrit).toFixed(1),
        Factor: conv.factor.toFixed(3),
        "Eff DPR": conv.effDpr.toFixed(2),
      });
    }

    return { total, rows };
  }

  function effPartyDprs(useNova) {
    if (!useNova) {
      return state.party_dpr_table.map((row) => [row.Member || "?", safeFloat(row.DPR, 0)]);
    }

    const out = [];
    for (const row of state.party_nova_table) {
      const conv = computeNovaConversion(row, state.boss_ac);
      out.push([row.Member || "?", conv.effDpr]);
    }
    return out;
  }

  function computeNovaConversion(row, fallbackBossAc) {
    const method = normalizeNovaMethod(row.Method);
    const mode = normalizeRollMode(row["Roll Mode"]);
    const novaDpr = Math.max(0, safeFloat(row["Nova DPR"], 0));
    const uptime = clamp(safeFloat(row.Uptime, 0.85), 0, 1);

    let pMain = 1;
    let pCrit = 0;
    let baseFactor = 1;

    if (method === "attack") {
      const atkBonus = safeInt(row["Atk Bonus"], 0);
      const targetAc = Math.max(1, safeInt(row["Target AC"], fallbackBossAc));
      const critRatio = Math.max(0, safeFloat(row["Crit Ratio"], 1.5));
      const [pNon, crit, pAny] = hitProbs(targetAc, atkBonus, mode);
      pMain = pAny;
      pCrit = crit;
      baseFactor = pNon + critRatio * pCrit;
    } else if (method === "save_half") {
      const saveDc = Math.max(1, safeInt(row["Save DC"], 16));
      const targetSaveBonus = safeInt(row["Target Save Bonus"], 0);
      const onSaveMult = clamp(safeFloat(row["Save Success Mult"], 0.5), 0, 1);
      const pFail = pSaveFailWithMode(saveDc, targetSaveBonus, mode);
      pMain = pFail;
      pCrit = 0;
      baseFactor = pFail + onSaveMult * (1 - pFail);
    }

    const factor = baseFactor * uptime;
    return {
      method,
      pMain,
      pCrit,
      factor,
      effDpr: novaDpr * factor,
    };
  }

  function perRoundDprVsPc(pcRow, mode, attacks) {
    const ac = Math.max(1, safeInt(pcRow.AC, 10));
    let total = 0;

    for (const atk of attacks) {
      if (atk.uses_per_round <= 0) continue;
      let dpr = 0;
      if (atk.kind === "save") {
        dpr = expectedSaveHalfDamage(atk.dc, getSaveBonus(pcRow, atk.save_stat), atk.damage_expr);
      } else {
        dpr = expectedAttackDamage(ac, atk.attack_bonus, atk.damage_expr, mode);
      }
      total += dpr * atk.uses_per_round;
    }

    return total;
  }

  function lairPerTargetDpr(opts, partySize) {
    if (!opts.lair_enabled || partySize <= 0) return 0;
    const pHit = Math.min(1, safeFloat(opts.lair_targets, 1) / partySize);
    const cadence = Math.max(1, safeInt(opts.lair_every_n, 1));
    return (safeFloat(opts.lair_avg, 0) * pHit) / cadence;
  }

  function rechargePerTargetDpr(opts, partySize) {
    if (!opts.rech_enabled || partySize <= 0) return 0;
    const rechargeProb = parseRecharge(opts.recharge_text || "5-6");
    const pHit = Math.min(1, safeFloat(opts.rech_targets, 1) / partySize);
    return rechargeProb * safeFloat(opts.rech_avg, 0) * pHit;
  }

  function attacksEnabledFromTable(tbl) {
    const out = [];
    for (const row of tbl) {
      if (!Boolean(row["Enabled?"])) continue;

      const saveStatRaw = String(row.Save || "DEX").toUpperCase();
      const saveStat = SAVE_KEYS.includes(saveStatRaw) ? saveStatRaw : "DEX";

      out.push({
        name: String(row.Name || "Attack"),
        kind: String(row.Type || "attack").toLowerCase() === "save" ? "save" : "attack",
        attack_bonus: safeInt(row["Attack bonus"], 0),
        dc: safeInt(row.DC, 0),
        save_stat: saveStat,
        damage_expr: String(row.Damage || "1d6"),
        uses_per_round: Math.max(0, safeInt(row["Uses/round"], 1)),
        is_melee: Boolean(row["Melee?"]),
      });
    }
    return out;
  }

  function getSaveBonus(pcRow, stat) {
    const key = String(stat || "DEX").toUpperCase();
    return safeInt(pcRow[key], 0);
  }

  function expectedAttackDamage(ac, attackBonus, dmgExpr, mode = "normal") {
    const [pNonCrit, pCrit] = hitProbs(ac, attackBonus, mode);
    return pNonCrit * averageDamage(dmgExpr) + pCrit * averageCrunchyCritDamage(dmgExpr);
  }

  function expectedSaveHalfDamage(dc, saveBonus, dmgExpr) {
    return (0.5 + 0.5 * pSaveFail(dc, saveBonus)) * averageDamage(dmgExpr);
  }

  function hitProbs(ac, attackBonus, mode = "normal") {
    const classify = (roll) => {
      if (roll === 1) return [false, false];
      if (roll === 20) return [true, true];
      return [roll + attackBonus >= ac, false];
    };

    if (mode === "normal") {
      let crits = 0;
      let nonCritHits = 0;
      for (let r = 1; r <= 20; r += 1) {
        const [isHit, isCrit] = classify(r);
        if (isCrit) crits += 1;
        else if (isHit) nonCritHits += 1;
      }
      return [nonCritHits / 20, crits / 20, (nonCritHits + crits) / 20];
    }

    let crits = 0;
    let nonCritHits = 0;
    for (let r1 = 1; r1 <= 20; r1 += 1) {
      for (let r2 = 1; r2 <= 20; r2 += 1) {
        const r = mode === "adv" ? Math.max(r1, r2) : Math.min(r1, r2);
        const [isHit, isCrit] = classify(r);
        if (isCrit) crits += 1;
        else if (isHit) nonCritHits += 1;
      }
    }

    return [nonCritHits / 400, crits / 400, (nonCritHits + crits) / 400];
  }

  function pSaveFail(dc, saveBonus) {
    const target = dc - saveBonus;
    if (target <= 1) return 1 / 20;
    if (target > 20) return 19 / 20;
    return (target - 1) / 20;
  }

  function pSaveFailWithMode(dc, saveBonus, mode = "normal") {
    const normalized = normalizeRollMode(mode);
    if (normalized === "normal") {
      return pSaveFail(dc, saveBonus);
    }

    let fails = 0;
    for (let r1 = 1; r1 <= 20; r1 += 1) {
      for (let r2 = 1; r2 <= 20; r2 += 1) {
        const roll = normalized === "adv" ? Math.max(r1, r2) : Math.min(r1, r2);
        const success = roll === 20 || (roll !== 1 && roll + saveBonus >= dc);
        if (!success) fails += 1;
      }
    }
    return fails / 400;
  }

  function normalizeRollMode(value) {
    const mode = String(value || "normal");
    return ["normal", "adv", "dis"].includes(mode) ? mode : "normal";
  }

  function normalizeNovaMethod(value) {
    const method = String(value || "attack");
    return NOVA_METHODS.includes(method) ? method : "attack";
  }

  function parseRecharge(text) {
    const t = String(text || "").trim().replace(/[\u2013\u2014]/g, "-");
    if (!t) return 0;

    if (t.includes("-")) {
      const [loS, hiS] = t.split("-", 2);
      const loRaw = safeInt(loS, 0);
      const hiRaw = safeInt(hiS, 0);
      if (!Number.isFinite(loRaw) || !Number.isFinite(hiRaw)) return 0;

      let lo = Math.min(loRaw, hiRaw);
      let hi = Math.max(loRaw, hiRaw);
      lo = clamp(lo, 1, 6);
      hi = clamp(hi, 1, 6);
      return (hi - lo + 1) / 6;
    }

    const k = clamp(safeInt(t, 0), 2, 6);
    if (!Number.isFinite(k) || k <= 0) return 0;
    return (7 - k) / 6;
  }

  function parseDamageExpression(expr) {
    let s = String(expr || "").trim().toLowerCase().replace(/\s+/g, "");
    if (!s) {
      return { sign: 1, dice: [], mod: 0 };
    }

    let sign = 1;
    if (s.startsWith("-")) sign = -1;
    if (s.startsWith("-") || s.startsWith("+")) {
      s = s.slice(1);
    }

    const dice = [];
    const regex = /([+-]?\d*)d(\d+)/g;
    let match;
    while ((match = regex.exec(s)) !== null) {
      const countStr = match[1];
      const sides = safeInt(match[2], 0);
      let count = 1;
      if (countStr === "-") count = -1;
      else if (countStr !== "" && countStr !== "+") count = safeInt(countStr, 1);
      dice.push([count, sides]);
    }

    const stripped = s.replace(/([+-]?\d*)d\d+/g, "");
    const tokens = stripped.match(/[+-]?\d+/g) || [];
    const mod = tokens.reduce((acc, token) => acc + safeInt(token, 0), 0);

    return { sign, dice, mod };
  }

  function averageDamage(expr) {
    const parsed = parseDamageExpression(expr);
    let avg = parsed.mod;
    for (const [count, sides] of parsed.dice) {
      avg += count * ((sides + 1) / 2);
    }
    return Math.max(0, parsed.sign * avg);
  }

  function averageCrunchyCritDamage(expr) {
    const parsed = parseDamageExpression(expr);
    let avgDice = 0;
    for (const [count, sides] of parsed.dice) {
      avgDice += count * (sides + (sides + 1) / 2);
    }
    return Math.max(0, parsed.sign * (avgDice + parsed.mod));
  }

  function rollDamageOne(expr) {
    const parsed = parseDamageExpression(expr);
    let total = parsed.mod;
    for (const [countRaw, sides] of parsed.dice) {
      if (sides <= 0 || countRaw === 0) continue;
      const count = Math.abs(countRaw);
      let subtotal = 0;
      for (let i = 0; i < count; i += 1) {
        subtotal += randomInt(1, sides);
      }
      total += Math.sign(countRaw) * subtotal;
    }
    total *= parsed.sign;
    return Math.max(0, total);
  }

  function rollDamageCrunchyCritOne(expr) {
    const parsed = parseDamageExpression(expr);
    let total = parsed.mod;

    for (const [countRaw, sides] of parsed.dice) {
      if (sides <= 0 || countRaw === 0) continue;
      const count = Math.abs(countRaw);
      let rolled = 0;
      for (let i = 0; i < count; i += 1) {
        rolled += randomInt(1, sides);
      }
      const maxPart = count * sides;
      total += Math.sign(countRaw) * (maxPart + rolled);
    }

    total *= parsed.sign;
    return Math.max(0, total);
  }

  function gammaRng(mean, cv) {
    const m = Math.max(0, safeFloat(mean, 0));
    if (m <= 0) return 0;
    const c = Math.max(1e-6, safeFloat(cv, 0.5));
    const k = 1 / (c * c);
    const theta = m / k;
    return gammaSample(k) * theta;
  }

  function gammaSample(shape) {
    if (shape < 1) {
      const u = Math.random();
      return gammaSample(shape + 1) * Math.pow(u, 1 / shape);
    }

    const d = shape - 1 / 3;
    const c = 1 / Math.sqrt(9 * d);

    while (true) {
      const x = normalSample();
      let v = 1 + c * x;
      if (v <= 0) continue;
      v = v * v * v;
      const u = Math.random();
      if (u < 1 - 0.0331 * Math.pow(x, 4)) return d * v;
      if (Math.log(u) < 0.5 * x * x + d * (1 - v + Math.log(v))) return d * v;
    }
  }

  function normalSample() {
    let u = 0;
    let v = 0;
    while (u === 0) u = Math.random();
    while (v === 0) v = Math.random();
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  }

  function renderResultTable(mount, rows) {
    mount.innerHTML = "";
    if (!rows || !rows.length) {
      const empty = document.createElement("div");
      empty.className = "table-empty";
      empty.textContent = "No rows.";
      mount.appendChild(empty);
      return;
    }

    const headers = Object.keys(rows[0]).filter((k) => !k.startsWith("_"));

    const wrap = document.createElement("div");
    wrap.className = "table-wrap";
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const hrow = document.createElement("tr");

    for (const h of headers) {
      const th = document.createElement("th");
      th.textContent = h;
      hrow.appendChild(th);
    }

    thead.appendChild(hrow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      for (const h of headers) {
        const td = document.createElement("td");
        td.textContent = String(row[h]);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrap.appendChild(table);
    mount.appendChild(wrap);
  }

  function drawChart(name, canvas, config) {
    if (charts[name]) {
      charts[name].destroy();
      charts[name] = null;
    }

    if (!canvas || !window.Chart) return;

    const ctx = canvas.getContext("2d");
    charts[name] = new window.Chart(ctx, config);
  }

  function clearChart(name) {
    if (charts[name]) {
      charts[name].destroy();
      charts[name] = null;
    }
  }

  function baseChartOptions(overrides = {}) {
    return mergeDeep({
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: {
        legend: { labels: { color: "#ecf3f9" } },
        title: { display: false, color: "#ecf3f9" },
      },
      scales: {
        x: { ticks: { color: "#a9c2d7" }, grid: { color: "rgba(255,255,255,0.08)" } },
        y: { ticks: { color: "#a9c2d7" }, grid: { color: "rgba(255,255,255,0.08)" } },
      },
    }, overrides);
  }

  function exportJson() {
    const payload = JSON.stringify(state, null, 2);
    const blob = new Blob([payload], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "kelemvor-profile.json";
    document.body.appendChild(a);
    a.click();
    a.remove();

    URL.revokeObjectURL(url);
    setStatus("Exported JSON profile.", 2000);
  }

  function refreshReport() {
    if (!els.reportText) return;
    els.reportText.value = JSON.stringify(state, null, 2);
  }

  function setStatus(message, durationMs = 2500) {
    if (!els.statusBar) return;
    if (statusTimer) {
      clearTimeout(statusTimer);
      statusTimer = null;
    }

    els.statusBar.textContent = message;
    if (durationMs > 0) {
      statusTimer = setTimeout(() => {
        els.statusBar.textContent = "Ready.";
      }, durationMs);
    }
  }

  function loadStateFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (_) {
      return null;
    }
  }

  function persistState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (_) {
      // Ignore quota/storage errors silently.
    }
  }

  function normalizeState(raw) {
    const base = deepClone(DEFAULT_STATE);
    const src = raw && typeof raw === "object" ? raw : {};

    for (const [key, defaultValue] of Object.entries(base)) {
      if (Array.isArray(defaultValue)) {
        if (Array.isArray(src[key])) base[key] = deepClone(src[key]);
      } else if (Object.prototype.hasOwnProperty.call(src, key)) {
        base[key] = src[key];
      }
    }

    base.party_table = (base.party_table || []).map(sanitizePartyRow);
    base.attacks_table = (base.attacks_table || []).map(sanitizeAttackRow);
    base.party_dpr_table = (base.party_dpr_table || []).map(sanitizeDprRow);
    base.party_nova_table = (base.party_nova_table || []).map(sanitizeNovaRow);

    base.mode_select = ["normal", "adv", "dis"].includes(String(base.mode_select)) ? String(base.mode_select) : "normal";
    base.spread_targets = Math.max(1, safeInt(base.spread_targets, 1));
    base.thp_expr = String(base.thp_expr || "0");

    base.lair_enabled = Boolean(base.lair_enabled);
    base.lair_avg = Math.max(0, safeFloat(base.lair_avg, 6.0));
    base.lair_targets = Math.max(1, safeInt(base.lair_targets, 2));
    base.lair_every_n = Math.max(1, safeInt(base.lair_every_n, 2));

    base.rech_enabled = Boolean(base.rech_enabled);
    base.recharge_text = String(base.recharge_text || "5-6");
    base.rech_avg = Math.max(0, safeFloat(base.rech_avg, 22.0));
    base.rech_targets = Math.max(1, safeInt(base.rech_targets, 1));

    base.rider_mode = ["none", "grant advantage on melee next round", "-2 AC next round"].includes(String(base.rider_mode))
      ? String(base.rider_mode)
      : "none";
    base.rider_duration = Math.max(1, safeInt(base.rider_duration, 1));
    base.rider_melee_only = Boolean(base.rider_melee_only);

    base.boss_hp = Math.max(1, safeInt(base.boss_hp, 150));
    base.boss_ac = Math.max(1, safeInt(base.boss_ac, 16));
    base.resist_factor = Math.max(0.01, safeFloat(base.resist_factor, 1.0));
    base.boss_regen = Math.max(0, safeFloat(base.boss_regen, 0.0));

    base.mc_rounds = Math.max(1, safeInt(base.mc_rounds, 3));
    base.mc_trials = Math.max(1000, safeInt(base.mc_trials, 10000));
    base.mc_show_hist = Boolean(base.mc_show_hist);

    base.enc_trials = Math.max(1000, safeInt(base.enc_trials, 10000));
    base.enc_max_rounds = Math.max(1, safeInt(base.enc_max_rounds, 12));
    base.enc_use_nova = Boolean(base.enc_use_nova);
    base.dpr_cv = clamp(safeFloat(base.dpr_cv, 0.6), 0.05, 2.0);
    base.initiative_mode = ["random", "party_first", "boss_first"].includes(String(base.initiative_mode))
      ? String(base.initiative_mode)
      : "random";

    base.tune_target_median = clamp(safeFloat(base.tune_target_median, 4.0), 1.0, 20.0);
    base.tune_tpk_cap = clamp(safeFloat(base.tune_tpk_cap, 0.05), 0.0, 1.0);

    syncPartyDependentRows(base);
    return base;
  }

  function syncPartyDependentRows(st) {
    const names = uniquePartyNames(st.party_table);

    const dprMap = new Map();
    for (const row of st.party_dpr_table || []) {
      const name = String(row.Member || "").trim();
      if (name && !dprMap.has(name)) dprMap.set(name, row);
    }

    const novaMap = new Map();
    for (const row of st.party_nova_table || []) {
      const name = String(row.Member || "").trim();
      if (name && !novaMap.has(name)) novaMap.set(name, row);
    }

    st.party_dpr_table = names.map((name) => {
      const row = dprMap.get(name) || {};
      return {
        Member: name,
        DPR: Math.max(0, safeFloat(row.DPR, 10.0)),
      };
    });

    const bossAc = Math.max(1, safeInt(st.boss_ac, 16));
    st.party_nova_table = names.map((name) => {
      const nRow = novaMap.get(name) || {};
      const dpr = st.party_dpr_table.find((r) => r.Member === name);
      return {
        Member: name,
        "Nova DPR": Math.max(0, safeFloat(dpr ? dpr.DPR : 10.0, 10.0)),
        Method: normalizeNovaMethod(nRow.Method),
        "Atk Bonus": safeInt(nRow["Atk Bonus"], 7),
        "Roll Mode": normalizeRollMode(nRow["Roll Mode"]),
        "Target AC": bossAc,
        "Crit Ratio": Math.max(0, safeFloat(nRow["Crit Ratio"], 1.5)),
        "Save DC": Math.max(1, safeInt(nRow["Save DC"], 16)),
        "Target Save Bonus": safeInt(nRow["Target Save Bonus"], 0),
        "Save Success Mult": clamp(safeFloat(nRow["Save Success Mult"], 0.5), 0, 1),
        Uptime: clamp(safeFloat(nRow.Uptime, 0.85), 0, 1),
      };
    });
  }

  function uniquePartyNames(rows) {
    const out = [];
    const seen = new Set();
    for (const row of rows || []) {
      const name = String(row.Name || "").trim();
      if (!name || seen.has(name)) continue;
      seen.add(name);
      out.push(name);
    }
    return out;
  }

  function sanitizePartyRow(row) {
    const out = { ...row };
    out.Name = String(out.Name || "").trim();
    out.AC = Math.max(1, safeInt(out.AC, 10));
    out.HP = Math.max(1, safeInt(out.HP, 1));
    for (const key of SAVE_KEYS) {
      out[key] = safeInt(out[key], 0);
    }
    return out;
  }

  function sanitizeDprRow(row) {
    return {
      Member: String(row.Member || "").trim(),
      DPR: Math.max(0, safeFloat(row.DPR, 0)),
    };
  }

  function sanitizeNovaRow(row) {
    return {
      Member: String(row.Member || "").trim(),
      "Nova DPR": Math.max(0, safeFloat(row["Nova DPR"], 0)),
      Method: normalizeNovaMethod(row.Method),
      "Atk Bonus": safeInt(row["Atk Bonus"], 7),
      "Roll Mode": normalizeRollMode(row["Roll Mode"]),
      "Target AC": Math.max(1, safeInt(row["Target AC"], 16)),
      "Crit Ratio": Math.max(0, safeFloat(row["Crit Ratio"], 1.5)),
      "Save DC": Math.max(1, safeInt(row["Save DC"], 16)),
      "Target Save Bonus": safeInt(row["Target Save Bonus"], 0),
      "Save Success Mult": clamp(safeFloat(row["Save Success Mult"], 0.5), 0, 1),
      Uptime: clamp(safeFloat(row.Uptime, 0.85), 0, 1),
    };
  }

  function sanitizeAttackRow(row) {
    return {
      Name: String(row.Name || "Attack").trim() || "Attack",
      Type: String(row.Type || "attack").toLowerCase() === "save" ? "save" : "attack",
      "Attack bonus": safeInt(row["Attack bonus"], 0),
      DC: safeInt(row.DC, 0),
      Save: SAVE_KEYS.includes(String(row.Save || "DEX").toUpperCase()) ? String(row.Save).toUpperCase() : "DEX",
      Damage: String(row.Damage || "1d6").trim() || "1d6",
      "Uses/round": Math.max(0, safeInt(row["Uses/round"], 1)),
      "Melee?": Boolean(row["Melee?"]),
      "Enabled?": Boolean(row["Enabled?"]),
    };
  }

  function histogram(values, bins = 30) {
    const arr = values.map((v) => safeFloat(v, 0)).filter((v) => Number.isFinite(v));
    if (!arr.length) return { labels: [], values: [] };

    const lo = Math.min(...arr);
    const hi = Math.max(...arr);

    if (lo === hi) {
      return { labels: [lo.toFixed(2)], values: [arr.length] };
    }

    const width = (hi - lo) / bins;
    const counts = new Array(bins).fill(0);

    for (const v of arr) {
      let idx = Math.floor((v - lo) / width);
      idx = clamp(idx, 0, bins - 1);
      counts[idx] += 1;
    }

    const labels = counts.map((_, i) => {
      const a = lo + i * width;
      const b = a + width;
      return `${a.toFixed(1)}-${b.toFixed(1)}`;
    });

    return { labels, values: counts };
  }

  function meanOf(values) {
    if (!values.length) return 0;
    let sum = 0;
    for (const v of values) sum += safeFloat(v, 0);
    return sum / values.length;
  }

  function percentile(values, p) {
    if (!values.length) return Number.NaN;
    const sorted = [...values].sort((a, b) => a - b);
    const rank = (clamp(p, 0, 100) / 100) * (sorted.length - 1);
    const lo = Math.floor(rank);
    const hi = Math.ceil(rank);
    if (lo === hi) return sorted[lo];
    const weight = rank - lo;
    return sorted[lo] + (sorted[hi] - sorted[lo]) * weight;
  }

  function sampleWithoutReplacement(arr, k) {
    const copy = arr.slice();
    for (let i = copy.length - 1; i > 0; i -= 1) {
      const j = randomInt(0, i);
      const tmp = copy[i];
      copy[i] = copy[j];
      copy[j] = tmp;
    }
    return copy.slice(0, k);
  }

  function randomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  function safeInt(x, defaultValue = 0) {
    const n = Number.parseInt(String(x), 10);
    if (Number.isFinite(n)) return n;
    const f = Number.parseFloat(String(x));
    if (Number.isFinite(f)) return Math.trunc(f);
    return defaultValue;
  }

  function safeFloat(x, defaultValue = 0) {
    const n = Number.parseFloat(String(x));
    return Number.isFinite(n) ? n : defaultValue;
  }

  function clamp(v, lo, hi) {
    return Math.max(lo, Math.min(hi, v));
  }

  function round2(v) {
    return Math.round(safeFloat(v, 0) * 100) / 100;
  }

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function setControlValue(el, value) {
    if (!el) return;
    el.value = String(value);
  }

  function setControlChecked(el, checked) {
    if (!el) return;
    el.checked = Boolean(checked);
  }

  function mergeDeep(base, override) {
    if (Array.isArray(base) || Array.isArray(override) || typeof base !== "object" || typeof override !== "object") {
      return override;
    }

    const out = { ...base };
    for (const [key, value] of Object.entries(override)) {
      if (value && typeof value === "object" && !Array.isArray(value) && base[key] && typeof base[key] === "object" && !Array.isArray(base[key])) {
        out[key] = mergeDeep(base[key], value);
      } else {
        out[key] = value;
      }
    }
    return out;
  }
})();
