(() => {
  "use strict";

  const SAVE_KEYS = ["STR", "DEX", "CON", "INT", "WIS", "CHA"];
  const NOVA_METHODS = ["attack", "save_half", "auto_hit"];
  const STORAGE_KEY = "kelemvor_scales_web_state_v1";
  const IS_WORKER = typeof document === "undefined" && typeof self !== "undefined";

  const DEFAULT_STATE = {
    party_table: [
      { Name: "Fighter", AC: 18, HP: 52, STR: 4, DEX: 2, CON: 3, INT: 0, WIS: 1, CHA: 0 },
      { Name: "Rogue",   AC: 16, HP: 38, STR: 0, DEX: 5, CON: 2, INT: 1, WIS: 2, CHA: 1 },
      { Name: "Cleric",  AC: 18, HP: 44, STR: 3, DEX: 0, CON: 3, INT: 1, WIS: 4, CHA: 2 },
      { Name: "Wizard",  AC: 13, HP: 32, STR: 0, DEX: 3, CON: 2, INT: 5, WIS: 2, CHA: 1 },
    ],
    attacks_table: [
      { Name: "Bite",        Type: "attack", "Attack bonus": 7, DC: 0,  Save: "DEX", Damage: "2d10+5", "Uses/round": 1, "Melee?": true,  "Enabled?": true },
      { Name: "Claw",        Type: "attack", "Attack bonus": 7, DC: 0,  Save: "DEX", Damage: "2d6+5",  "Uses/round": 2, "Melee?": true,  "Enabled?": true },
      { Name: "Fire Breath", Type: "save",   "Attack bonus": 0, DC: 15, Save: "DEX", Damage: "8d6",    "Uses/round": 1, "Melee?": false, "Enabled?": true },
    ],
    party_dpr_table: [
      { Member: "Fighter", Damage: "1d8+5" },
      { Member: "Rogue",   Damage: "3d6+5" },
      { Member: "Cleric",  Damage: "1d8+3" },
      { Member: "Wizard",  Damage: "4d6" },
    ],
    party_nova_table: [
      { Member: "Fighter", Method: "attack",    "Attacks": 2, "Atk Bonus": 8, "Roll Mode": "normal", "Target AC": 16, "Crit": 20, "Save DC": 16, "Target Save Bonus": 0, "Save Success Mult": 0.5, Uptime: 0.9 },
      { Member: "Rogue",   Method: "attack",    "Attacks": 1, "Atk Bonus": 7, "Roll Mode": "normal", "Target AC": 16, "Crit": 20, "Save DC": 16, "Target Save Bonus": 0, "Save Success Mult": 0.5, Uptime: 0.85 },
      { Member: "Cleric",  Method: "attack",    "Attacks": 1, "Atk Bonus": 6, "Roll Mode": "normal", "Target AC": 16, "Crit": 20, "Save DC": 15, "Target Save Bonus": 0, "Save Success Mult": 0.5, Uptime: 0.8 },
      { Member: "Wizard",  Method: "save_half", "Attacks": 1, "Atk Bonus": 0, "Roll Mode": "normal", "Target AC": 16, "Crit": 20, "Save DC": 16, "Target Save Bonus": 2,  "Save Success Mult": 0.5, Uptime: 0.85 },
    ],

    mode_select: "normal",
    spread_targets: 1,
    thp_expr: "0",

    lair_enabled: false,
    lair_avg: 6.0,
    lair_formula: "1d10+1",
    lair_targets: 2,
    lair_every_n: 2,

    rech_enabled: false,
    recharge_text: "5-6",
    rech_avg: 22.0,
    rech_formula: "4d10",
    rech_targets: 1,
    phase_table: [],

    rider_mode: "none",
    rider_duration: 1,
    rider_melee_only: true,

    boss_hp: 200,
    boss_ac: 16,
    resist_factor: 1.0,
    boss_regen: 0.0,
    boss_dpr_mult: 1.0,

    mc_rounds: 3,
    mc_trials: 10000,
    mc_show_hist: true,

    enc_trials: 10000,
    enc_max_rounds: 12,
    enc_use_nova: true,
    dpr_cv: 0.6,
    initiative_mode: "random",

    tune_target_median: 5.0,
    tune_tpk_cap: 0.05,
    tune_kill_rate: 0.75,
    tune_band_max: 3.0,

    pacing_rounds: 5,
    tight_pacing_enabled: true,
    tight_cap_pct: 0.24,
    tight_cap_resources: 3,
    tight_anti_slog_mult: 1.35,

    minion_count: 0,
    minion_ac: 14,
    minion_hp: 15,
    minion_replenish: false,

    minion_atk_enabled: false,
    minion_atk_bonus: 4,
    minion_atk_avg: 5,

    minion_save_enabled: false,
    minion_save_stat: "DEX",
    minion_save_dc: 13,
    minion_save_avg: 7,
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
    { key: "Damage", label: "Dmg/Attack", type: "text", parser: (v) => String(v).trim() || "1d6" },
  ];

  const NOVA_COLUMNS = [
    { key: "Member", label: "Member", type: "text", readOnly: true },
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
    { key: "Attacks", label: "Attacks", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 1)) },
    { key: "Target AC", label: "Target AC", type: "number", readOnly: true },
    { key: "Crit", label: "Crit ≥", type: "number", step: "1", min: "2", max: "20", parser: (v) => Math.max(2, Math.min(20, safeInt(v, 20))) },
    { key: "Save DC", label: "Save DC", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 16)) },
    { key: "Target Save Bonus", label: "Save Bonus", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Save Success Mult", label: "Save Success Mult", type: "number", step: "0.01", min: "0", max: "1", parser: (v) => clamp(safeFloat(v, 0.5), 0, 1) },
    { key: "Uptime", label: "Uptime", type: "number", step: "0.01", min: "0", max: "1", parser: (v) => clamp(safeFloat(v, 0.85), 0, 1) },
  ];

  const ATTACK_COLUMNS = [
    { key: "Name", label: "Name", type: "text", parser: (v) => String(v).trim() },
    { key: "Type", label: "Type", type: "select", options: ["attack", "save"], parser: (v) => (String(v).toLowerCase() === "save" ? "save" : "attack") },
    { key: "Attack bonus", label: "Atk Bonus", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Crit", label: "Crit ≥", type: "number", step: "1", min: "2", max: "20", parser: (v) => Math.max(2, Math.min(20, safeInt(v, 20))) },
    { key: "DC", label: "DC", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Save", label: "Save", type: "select", options: SAVE_KEYS, parser: (v) => (SAVE_KEYS.includes(String(v).toUpperCase()) ? String(v).toUpperCase() : "DEX") },
    { key: "Damage", label: "Damage", type: "text", parser: (v) => String(v).trim() || "0" },
    { key: "Uses/round", label: "Uses/round", type: "number", step: "1", min: "0", parser: (v) => Math.max(0, safeInt(v, 1)) },
    { key: "Uses/encounter", label: "Uses/enc", type: "number", step: "1", min: "0", parser: (v) => Math.max(0, safeInt(v, 0)) },
    { key: "Melee?", label: "Melee?", type: "checkbox", parser: (v) => Boolean(v) },
    { key: "Enabled?", label: "Enabled?", type: "checkbox", parser: (v) => Boolean(v) },
  ];

  const MECHANIC_TYPES = ["auto", "attack", "save"];
  const PHASE_COLUMNS = [
    { key: "Round", label: "Round", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 1)) },
    { key: "Name", label: "Name", type: "text", parser: (v) => String(v).trim() },
    { key: "Type", label: "Type", type: "select", options: MECHANIC_TYPES, parser: (v) => normalizeMechanicType(v) },
    { key: "Attack bonus", label: "Atk Bonus", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Crit", label: "Crit ≥", type: "number", step: "1", min: "2", max: "20", parser: (v) => Math.max(2, Math.min(20, safeInt(v, 20))) },
    { key: "DC", label: "DC", type: "number", step: "1", parser: (v) => safeInt(v, 0) },
    { key: "Save", label: "Save", type: "select", options: SAVE_KEYS, parser: (v) => (SAVE_KEYS.includes(String(v).toUpperCase()) ? String(v).toUpperCase() : "DEX") },
    { key: "Damage", label: "Damage", type: "text", parser: (v) => String(v).trim() || "0" },
    { key: "Targets", label: "Targets", type: "number", step: "1", min: "1", parser: (v) => Math.max(1, safeInt(v, 1)) },
    { key: "Enabled?", label: "Enabled?", type: "checkbox", parser: (v) => Boolean(v) },
  ];

  let state = normalizeState(IS_WORKER ? {} : (loadStateFromStorage() || {}));
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

  if (!IS_WORKER) {
    document.addEventListener("DOMContentLoaded", init);
  }

  function init() {
    cacheElements();
    bindTabs();
    bindGlobalActions();
    bindButtons();
    bindOptionControls();
    initResizableWidgets();
    renderAll();
    refreshDerivedCv();
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
      "btnAddAttackRow", "btnAddLimitedAttackRow", "attacksTable", "btnAddPhaseMechanic", "phaseTable",
      "optModeSelect", "optSpreadTargets", "optThpExpr", "optBossHp", "optBossAc", "optResistFactor", "optBossRegen", "optBossDprMult",
      "optLairEnabled", "optLairAvg", "optLairFormula", "optLairTargets", "optLairEveryN",
      "optRechEnabled", "optRechargeText", "optRechAvg", "optRechFormula", "optRechTargets",
      "optRiderMode", "optRiderDuration", "optRiderMeleeOnly",
      "btnComputeDet", "detTable", "detChart",
      "ttdModeManual", "ttdModeNova", "btnComputeTtd", "effTable",
      "lblIncoming", "lblRoundsExact", "lblRoundsCeil",
      "mcTarget", "mcRounds", "mcTrials", "mcShowHist", "btnRunMc", "mcChart",
      "lblMcMean", "lblMcP95", "lblMcP99",
      "encTrials", "encMaxRounds", "encDprCv", "encInitiative", "encUseNova",
      "tuneMedian", "tuneTpkCap", "tuneKillRate", "tuneBandMax",
      "btnRunEncounter", "btnAutoTuneAll",
      "lblTtkMedian", "lblTtkP1090", "lblTpk", "lblDowns",
      "survChart", "ttkChart",
      "reportText",
      "pacingRounds",
      "pacingBossHp", "pacingBossHpSub", "pacingFirstDown", "pacingPartyWipe",
      "pacingTargetDpr", "pacingTargetDprSub", "pacingBalance", "pacingTable",
      "optMinionCount", "optMinionAc", "optMinionHp", "optMinionReplenish",
      "optMinionAtkEnabled", "optMinionAtkBonus", "optMinionAtkAvg",
      "optMinionSaveEnabled", "optMinionSaveStat", "optMinionSaveDc", "optMinionSaveAvg",
      "minionTtdCard", "minionTtdTable",
    ];

    for (const id of ids) {
      els[id] = document.getElementById(id);
    }
  }

  function bindTabs() {
    const tabButtons = Array.from(document.querySelectorAll(".tab"));
    const panels = {
      party:  document.getElementById("panel-party"),
      boss:   document.getElementById("panel-boss"),
      det:    document.getElementById("panel-det"),
      ttd:    document.getElementById("panel-ttd"),
      mc:     document.getElementById("panel-mc"),
      enc:    document.getElementById("panel-enc"),
      report: document.getElementById("panel-report"),
    };

    for (const button of tabButtons) {
      button.addEventListener("click", () => {
        const tab = button.dataset.tab;
        tabButtons.forEach((b) => b.classList.toggle("active", b === button));
        Object.entries(panels).forEach(([k, panel]) => {
          if (panel) panel.classList.toggle("active", k === tab);
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
        "Uses/encounter": 0,
        "Melee?": true,
        "Enabled?": true,
      });
      persistState();
      renderAttackSection();
      refreshReport();
      setStatus("Added attack row.", 1500);
    });

    els.btnAddLimitedAttackRow.addEventListener("click", () => {
      state.attacks_table.push({
        Name: "Limited Ability",
        Type: "save",
        "Attack bonus": 0,
        DC: 15,
        Save: "DEX",
        Damage: "4d8",
        "Uses/round": 1,
        "Uses/encounter": 1,
        "Melee?": false,
        "Enabled?": true,
      });
      persistState();
      renderAttackSection();
      refreshReport();
      setStatus("Added limited ability row.", 1500);
    });

    els.btnAddPhaseMechanic.addEventListener("click", () => {
      state.phase_table.push({
        Round: 3,
        Name: "Raidwide",
        Type: "save",
        "Attack bonus": 0,
        Crit: 20,
        DC: 15,
        Save: "DEX",
        Damage: "6d6",
        Targets: 99,
        "Enabled?": true,
      });
      persistState();
      renderPhaseSection();
      refreshReport();
      setStatus("Added round mechanic.", 1500);
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

    els.btnRunEncounter.addEventListener("click", () => {
      runButtonAction(els.btnRunEncounter, runEncounterAndRender, "Full analysis failed.");
    });

    els.btnAutoTuneAll.addEventListener("click", () => {
      runButtonAction(els.btnAutoTuneAll, autoTuneAll, "Encounter tuning failed.");
    });

    els.btnRunMc.addEventListener("click", () => {
      runButtonAction(els.btnRunMc, runSingleTargetMc, "Single-target Monte Carlo failed.");
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
    bindControl("optBossDprMult", "boss_dpr_mult", (el) => clamp(safeFloat(el.value, 1.0), 0, 20));

    bindControl("optLairEnabled", "lair_enabled", (el) => Boolean(el.checked));
    bindControl("optLairAvg", "lair_avg", (el) => {
      const avg = Math.max(0, safeFloat(el.value, 0.0));
      state.lair_formula = damageFormulaForAverage(avg, 6);
      setControlValue(els.optLairFormula, state.lair_formula);
      return avg;
    });
    bindControl("optLairFormula", "lair_formula", (el) => {
      const formula = String(el.value || "").trim();
      if (formula) {
        state.lair_avg = round2(averageDamage(formula));
        setControlValue(els.optLairAvg, state.lair_avg);
      }
      return formula;
    });
    bindControl("optLairTargets", "lair_targets", (el) => Math.max(1, safeInt(el.value, 1)));
    bindControl("optLairEveryN", "lair_every_n", (el) => Math.max(1, safeInt(el.value, 1)));

    bindControl("optRechEnabled", "rech_enabled", (el) => Boolean(el.checked));
    bindControl("optRechargeText", "recharge_text", (el) => String(el.value || "5-6").trim() || "5-6");
    bindControl("optRechAvg", "rech_avg", (el) => {
      const avg = Math.max(0, safeFloat(el.value, 0.0));
      state.rech_formula = damageFormulaForAverage(avg, 6);
      setControlValue(els.optRechFormula, state.rech_formula);
      return avg;
    });
    bindControl("optRechFormula", "rech_formula", (el) => {
      const formula = String(el.value || "").trim();
      if (formula) {
        state.rech_avg = round2(averageDamage(formula));
        setControlValue(els.optRechAvg, state.rech_avg);
      }
      return formula;
    });
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

    bindControl("tuneMedian",    "tune_target_median", (el) => clamp(safeFloat(el.value, 5.0), 1.0, 20.0));
    bindControl("tuneTpkCap",    "tune_tpk_cap",       (el) => clamp(safeFloat(el.value, 0.05), 0.0, 1.0));
    bindControl("tuneKillRate",  "tune_kill_rate",     (el) => clamp(safeFloat(el.value, 0.75), 0.50, 0.95));
    bindControl("tuneBandMax",   "tune_band_max",      (el) => Math.max(0.5, safeFloat(el.value, 3.0)));

    bindControl("pacingRounds", "pacing_rounds", (el) => {
      const rounds = clamp(safeInt(el.value, 5), 5, 10);
      state.tune_target_median = rounds;
      return rounds;
    });

    bindControl("optMinionCount",       "minion_count",       (el) => Math.max(0, safeInt(el.value, 0)),    { refreshEff: true });
    bindControl("optMinionAc",          "minion_ac",          (el) => Math.max(1, safeInt(el.value, 14)),   { refreshEff: true });
    bindControl("optMinionHp",          "minion_hp",          (el) => Math.max(1, safeFloat(el.value, 15)), { refreshEff: true });
    bindControl("optMinionReplenish",   "minion_replenish",   (el) => Boolean(el.checked),                  { refreshEff: true });
    bindControl("optMinionAtkEnabled",  "minion_atk_enabled", (el) => Boolean(el.checked),                  { refreshEff: true });
    bindControl("optMinionAtkBonus",    "minion_atk_bonus",   (el) => clamp(safeInt(el.value, 4), -10, 20), { refreshEff: true });
    bindControl("optMinionAtkAvg",      "minion_atk_avg",     (el) => Math.max(0, safeFloat(el.value, 5)),  { refreshEff: true });
    bindControl("optMinionSaveEnabled", "minion_save_enabled",(el) => Boolean(el.checked),                  { refreshEff: true });
    bindControl("optMinionSaveStat",    "minion_save_stat",   (el) => String(el.value || "DEX"),             { refreshEff: true });
    bindControl("optMinionSaveDc",      "minion_save_dc",     (el) => clamp(safeInt(el.value, 13), 1, 30),  { refreshEff: true });
    bindControl("optMinionSaveAvg",     "minion_save_avg",    (el) => Math.max(0, safeFloat(el.value, 7)),  { refreshEff: true });
  }

  function runButtonAction(button, action, failureMessage) {
    setBtnLoading(button, true);
    setTimeout(async () => {
      try {
        await action();
      } catch (error) {
        reportActionError(error, failureMessage);
      } finally {
        setBtnLoading(button, false);
      }
    }, 0);
  }

  function runGuardedAction(action, failureMessage) {
    try {
      action();
    } catch (error) {
      reportActionError(error, failureMessage);
    }
  }

  function reportActionError(error, failureMessage) {
    console.error(failureMessage, error);
    const detail = error && error.message ? error.message : String(error);
    setStatus(`${failureMessage} ${detail}`, 5000);
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
    renderPhaseSection();
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
    setControlValue(els.optBossDprMult, state.boss_dpr_mult);

    setControlChecked(els.optLairEnabled, state.lair_enabled);
    setControlValue(els.optLairAvg, state.lair_avg);
    setControlValue(els.optLairFormula, state.lair_formula);
    setControlValue(els.optLairTargets, state.lair_targets);
    setControlValue(els.optLairEveryN, state.lair_every_n);

    setControlChecked(els.optRechEnabled, state.rech_enabled);
    setControlValue(els.optRechargeText, state.recharge_text);
    setControlValue(els.optRechAvg, state.rech_avg);
    setControlValue(els.optRechFormula, state.rech_formula);
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

    setControlValue(els.tuneMedian,    state.tune_target_median);
    setControlValue(els.tuneTpkCap,    state.tune_tpk_cap);
    setControlValue(els.tuneKillRate,  state.tune_kill_rate);
    setControlValue(els.tuneBandMax,   state.tune_band_max);

    setControlValue(els.pacingRounds, state.pacing_rounds);

    setControlValue(els.optMinionCount,    state.minion_count);
    setControlValue(els.optMinionAc,       state.minion_ac);
    setControlValue(els.optMinionHp,       state.minion_hp);
    setControlChecked(els.optMinionReplenish, state.minion_replenish);
    setControlChecked(els.optMinionAtkEnabled, state.minion_atk_enabled);
    setControlValue(els.optMinionAtkBonus, state.minion_atk_bonus);
    setControlValue(els.optMinionAtkAvg,   state.minion_atk_avg);
    setControlChecked(els.optMinionSaveEnabled, state.minion_save_enabled);
    if (els.optMinionSaveStat) els.optMinionSaveStat.value = state.minion_save_stat || "DEX";
    setControlValue(els.optMinionSaveDc,   state.minion_save_dc);
    setControlValue(els.optMinionSaveAvg,  state.minion_save_avg);
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
        refreshDerivedCv();
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
        refreshDerivedCv();
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
      updateTtdLabels(total, true);
      return;
    }

    const totalDpr = state.party_dpr_table.reduce((acc, row) => acc + averageDamage(row.Damage || "1d6"), 0);
    renderResultTable(els.effTable, state.party_dpr_table.map((row) => ({ Member: row.Member, "Avg/Attack": round2(averageDamage(row.Damage || "1d6")), "Formula": row.Damage || "1d6" })));
    updateTtdLabels(totalDpr, false);
  }

  function updateTtdLabels(totalDpr, useNova = Boolean(state.enc_use_nova)) {
    const resist    = Math.max(1e-6, safeFloat(state.resist_factor, 1.0));
    const regen     = Math.max(0, safeFloat(state.boss_regen, 0.0));
    // effective = how much boss HP the party removes per round (after resist/regen).
    const effective = Math.max(0, totalDpr / resist - regen);
    const hp        = Math.max(1, safeFloat(state.boss_hp, 150));

    // Minion phase: pass raw totalDpr — resist/regen are boss properties, not minion properties.
    // computeMinionPhase applies resist/regen only to boss-overflow damage.
    const minionPhase = computeMinionPhase(totalDpr, useNova);
    renderMinionTtdTable(minionPhase);

    let exact;
    let suffix = "";
    if (minionPhase) {
      if (!Number.isFinite(minionPhase.minionPhaseRounds)) {
        // Replenishing minions block the party — boss never dies under pure focus-fire
        exact  = Number.POSITIVE_INFINITY;
        suffix = " (replenishing minions screen boss — party DPR too low to one-round the pack)";
      } else if (minionPhase.replenishing && minionPhase.canClearPerRound) {
        // Replenishing but party clears each round; minionPhaseRounds = rounds to kill boss
        exact  = minionPhase.minionPhaseRounds;
        suffix = " (replenishing pack cleared each round; overflow kills boss)";
      } else {
        // Non-replenishing: minion phase clears, then party focuses boss
        const remainingBossHp = minionPhase.bossHpAfterMinionPhase;
        const bossPhaseRounds = effective > 0 ? remainingBossHp / effective : Number.POSITIVE_INFINITY;
        exact  = minionPhase.minionPhaseRounds + bossPhaseRounds;
        suffix = ` (incl. ${minionPhase.minionPhaseRounds}R minion phase)`;
      }
    } else {
      exact = effective > 0 ? hp / effective : Number.POSITIVE_INFINITY;
    }

    els.lblIncoming.textContent    = `Effective Party DPR: ${effective.toFixed(2)}`;
    els.lblRoundsExact.textContent = `Exact Rounds to Zero: ${Number.isFinite(exact) ? exact.toFixed(2) : "inf"}`;
    els.lblRoundsCeil.textContent  = `Boss Defeated In: ${Number.isFinite(exact) ? String(Math.ceil(exact)) : "inf"} rounds${suffix}`;
  }

  function computeDeterministic() {
    const attacks = attacksEnabledFromTable(state.attacks_table);
    const mechanics = phaseMechanicsEnabledFromTable(state.phase_table);
    const party = state.party_table.filter((r) => String(r.Name || "").trim().length > 0);
    const dprMult = bossDprMultiplier(state);
    const lairRechDpr = lairPerTargetDpr(state, party.length || 1) + rechargePerTargetDpr(state, party.length || 1);
    const thpAvg = Math.max(0, averageDamage(state.thp_expr || "0"));
    const spread = Math.max(1, safeInt(state.spread_targets, 1));
    const horizonRounds = Math.max(1, safeInt(state.pacing_rounds || state.enc_max_rounds, 1));

    const rows = [];
    const chartLabels = [];
    const chartValues = [];

    for (const pc of party) {
      const baseDpr = perRoundDprVsPc(pc, state.mode_select || "normal", attacks, horizonRounds);
      const phaseDpr = phaseMechanicsPerTargetDpr(pc, state.mode_select || "normal", mechanics, party.length, horizonRounds);
      const minionDpr = minionDprVsPc(pc);
      const total = (baseDpr / spread + lairRechDpr + minionDpr + phaseDpr) * dprMult;
      const net = Math.max(0, total - thpAvg);
      const hp = Math.max(1, safeInt(pc.HP, 1));
      const exact = net > 0 ? hp / net : Number.POSITIVE_INFINITY;
      const ceil = Number.isFinite(exact) ? Math.ceil(exact) : null;

      rows.push({
        Name: pc.Name || "?",
        AC: safeInt(pc.AC, 10),
        HP: hp,
        "DPR (attacks)": round2(baseDpr / spread),
        "DPR (mechanics)": round2(phaseDpr),
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
          backgroundColor: "rgba(191,26,47,0.72)",
          borderColor: "rgba(191,26,47,1)",
          borderWidth: 1,
        }],
      },
      options: baseChartOptions({
        plugins: { title: { display: true, text: "Time-To-Zero per Party Member" } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: "Rounds", color: "#aa9080" } },
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
            backgroundColor: "rgba(200,152,40,0.72)",
            borderColor: "rgba(200,152,40,1)",
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

  async function runEncounterAndRender() {
    setStatus("Running encounter simulation…", 0);
    const pacing = computePacingResult();
    _lastPacingResult = pacing;
    renderPacingResult(pacing);
    if (state.enc_max_rounds < pacing.targetRounds) {
      state.enc_max_rounds = pacing.targetRounds;
      syncControlsFromState();
      persistState();
    }
    updatePacingActionButtons(pacing);
    const runner = createEncounterWorkerRunner();
    let metrics;
    try {
      metrics = await runner.run({});
    } finally {
      runner.close();
    }
    if (metrics.error) {
      alert(metrics.error);
      setStatus("Encounter simulation failed.", 2600);
      return null;
    }

    const finite = metrics.finiteTtk;
    if (!finite.length) {
      alert("Boss never died within max rounds. Increase max rounds or lower boss durability.");
      setStatus("Encounter simulation returned no boss defeats.", 2600);
      return metrics;
    }

    const p10 = percentile(finite, 10);
    const med = percentile(finite, 50);
    const p90 = percentile(finite, 90);
    const tpk = metrics.tpkProb;
    const downs = metrics.pcsDownAtVictory;
    const killRate = (100 * finite.length / Math.max(1, metrics.ttk.length)).toFixed(1);

    els.lblTtkMedian.textContent = `Median TTK: ${med.toFixed(2)} rounds (${killRate}% kill rate)`;
    els.lblTtkP1090.textContent = `TTK p10-p90: ${p10.toFixed(2)} – ${p90.toFixed(2)} rounds (among kills)`;
    els.lblTpk.textContent = `TPK Probability: ${(100 * tpk).toFixed(1)}%`;
    els.lblDowns.textContent = `PCs down at victory: mean ${meanOf(downs).toFixed(2)}, p90 ${percentile(downs, 90).toFixed(0)}`;

    drawChart("surv", els.survChart, {
      type: "line",
      data: {
        labels: metrics.times,
        datasets: [{
          label: "S(t): Boss alive (kills only)",
          data: metrics.survivalCurve,
          stepped: "before",
          borderColor: "rgba(191,26,47,1)",
          borderWidth: 2,
          backgroundColor: "rgba(191,26,47,0.18)",
          fill: true,
          pointRadius: 0,
          tension: 0,
        }],
      },
      options: baseChartOptions({
        plugins: { title: { display: true, text: "Boss Survival Curve (kills only)" } },
        scales: {
          y: {
            min: 0, max: 1,
            title: { display: true, text: "Probability", color: "#aa9080" },
            ticks: { callback: (v) => `${(v * 100).toFixed(0)}%` },
          },
          x: {
            title: { display: true, text: "Rounds", color: "#aa9080" },
            ticks: { maxTicksLimit: 13 },
          },
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
          backgroundColor: "rgba(191,26,47,0.70)",
          borderColor: "rgba(191,26,47,1)",
          borderWidth: 1,
          barPercentage: 1,
          categoryPercentage: 1,
        }],
      },
      options: baseChartOptions({
        plugins: { title: { display: true, text: "TTK Distribution" }, legend: { display: false } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: "Frequency", color: "#aa9080" } },
          x: { ticks: { maxTicksLimit: 10 } },
        },
      }),
    });

    setStatus(`Encounter simulation complete. Trials=${state.enc_trials}.`, 3800);
    return metrics;
  }

  function autoTuneAllLegacy() {
    setStatus("Auto-tuning encounter… Stage 1: boss DPR.", 0);

    // targetRound is mutable — the tuner may adjust it within [5, 10] for band compliance.
    let targetRound      = clamp(safeInt(state.pacing_rounds, 5), 5, 10);
    const origTarget     = targetRound;
    const killRateTarget = clamp(safeFloat(state.tune_kill_rate, 0.75), 0.50, 0.95);
    const tpkCap         = safeFloat(state.tune_tpk_cap, 0.05);
    const bandTarget     = Math.max(0.5, safeFloat(state.tune_band_max, 3.0));
    const originalTrials = safeInt(state.enc_trials, 10000);
    const quickTrials    = Math.max(3000, Math.floor(originalTrials * 0.4));
    const changes        = [];

    // ── Stage 1: Boss DPR multiplier from pacing ───────────────────────────
    const pacingResult = computePacingResult();
    if (Number.isFinite(pacingResult.targetBossDprMult) && pacingResult.targetBossDprMult > 0) {
      const newMult = clamp(round2(pacingResult.targetBossDprMult), 0, 20);
      if (Math.abs(newMult - state.boss_dpr_mult) > 0.01) {
        changes.push(`Boss DPR mult: ${state.boss_dpr_mult.toFixed(2)}x → ${newMult.toFixed(2)}x`);
        state.boss_dpr_mult = newMult;
      }
    }

    // ── Stage 2: Minion HP for reliable clearing (replenishing packs only) ─
    const minionCount = Math.max(0, safeInt(state.minion_count, 0));
    if (minionCount > 0 && Boolean(state.minion_replenish)) {
      const mph = pacingResult.minionPhaseInfo;
      if (!mph || mph.partyDprVsMinions <= 0) {
        // Party does 0 DPR vs minions — encounter is impossible.
        const partial = changes.length ? `Applied so far:\n${changes.map(c => `  • ${c}`).join("\n")}\n\n` : "";
        alert(`${partial}⚠ Impossible encounter: Party cannot damage minions at AC ${state.minion_ac}.\nReduce minion AC, add melee/ranged attacks, or remove the minion pack.\n\nAuto-tune aborted.`);
        setStatus("Auto-tune aborted — minions are unkillable.", 4000);
        return;
      }
      const cv     = clamp(safeFloat(state.dpr_cv, 0.6), 0.05, 2.0);
      // Gamma quantile: HP pool the party can burn through with the given reliability.
      const poolHp = mph.partyDprVsMinions * Math.max(0.05, 1 + invNorm01(1 - killRateTarget) * cv);
      const newHp  = Math.max(1, Math.round(poolHp / minionCount));
      // Sanity check: if HP per minion would exceed the one-round-clear limit, the count is too high.
      const clearLimit = Math.floor(mph.partyDprVsMinions / minionCount);
      if (clearLimit < 1) {
        const partial = changes.length ? `Applied so far:\n${changes.map(c => `  • ${c}`).join("\n")}\n\n` : "";
        alert(`${partial}⚠ Too many minions: ${minionCount} minions with any HP cannot all be cleared in one round.\nParty deals ${mph.partyDprVsMinions.toFixed(0)} DPR vs minions — need fewer than ${Math.floor(mph.partyDprVsMinions)} minions for even 1 HP each.\nReduce minion count or remove replenish, then re-tune.\n\nAuto-tune aborted.`);
        setStatus("Auto-tune aborted — too many replenishing minions.", 4000);
        return;
      }
      if (newHp !== Math.round(state.minion_hp)) {
        changes.push(`Minion HP: ${Math.round(state.minion_hp)} → ${newHp} (one-round-clearable)`);
        state.minion_hp = newHp;
      }
    }

    // ── Stage 3: Boss HP — binary search for median TTK = targetRound ──────
    // The kill-rate objective (% dying by round N) sets the median well below the target
    // when variance is high.  Instead we target the median directly: find the highest HP
    // where the P50 of finite-TTK trials ≈ targetRound.  Kill rate becomes a post-hoc check.
    setStatus("Auto-tuning encounter… Stage 3: boss HP.", 0);

    const runSim = (hp, trials) => {
      const m = runEncounterMc({ boss_hp: Math.max(1, Math.round(hp)), enc_trials: Math.max(1000, Math.round(trials)) });
      if (m.error) return null;
      const finite = m.ttk.filter(v => Number.isFinite(v));
      const med  = percentile(m.ttk, 50);
      const p10  = finite.length > 0 ? percentile(finite, 10) : 0;
      const p90  = finite.length > 0 ? percentile(finite, 90) : 0;
      const kr   = m.ttk.filter(v => Number.isFinite(v) && v <= targetRound).length / m.ttk.length;
      return { median: med, band: p90 - p10, tpk: m.tpkProb, killRate: kr, ttk: m.ttk, finite };
    };

    // Binary-search HP so that median(finiteTTK) ≈ tgtRound.
    // Higher HP → longer fight → higher median.  "lo" is always a HP where median < tgtRound.
    const hpSearch = (tgtRound) => {
      let lo = 1.0, hi = Math.max(10.0, safeFloat(state.boss_hp, 150));
      // Expand upper bracket until median exceeds target.
      for (let a = 0; a < 14; a++) {
        const r = runSim(hi, quickTrials);
        if (!r || r.median >= tgtRound) break;
        hi *= 2;
      }
      for (let i = 0; i < 22; i++) {
        const mid = (lo + hi) / 2;
        const r = runSim(mid, quickTrials);
        if (!r || r.median < tgtRound) { lo = mid; } else { hi = mid; }
        if (hi - lo < 1.0) break;
      }
      return Math.max(1, Math.round(lo));
    };

    const oldBossHp = safeFloat(state.boss_hp, 150);
    let tunedHp = hpSearch(targetRound);

    // ── Stage 3.5: Band tightening ─────────────────────────────────────────
    const getBand = (hp) => {
      const r = runSim(hp, quickTrials);
      return r ? r.band : Infinity;
    };

    let actualBand = getBand(tunedHp);

    // Step B: Reduce DPR CV proportionally if band exceeds target.
    if (!Boolean(state.enc_use_nova) && actualBand > bandTarget + 0.05) {
      const currentCv = clamp(safeFloat(state.dpr_cv, 0.6), 0.05, 2.0);
      const neededCv  = Math.max(0.10, currentCv * bandTarget / actualBand);
      if (neededCv < currentCv - 0.01) {
        const oldCv = currentCv;
        state.dpr_cv = Math.round(neededCv * 100) / 100;
        tunedHp    = hpSearch(targetRound);
        actualBand = getBand(tunedHp);
        changes.push(`DPR CV: ${oldCv.toFixed(2)} → ${state.dpr_cv.toFixed(2)} (tightening variance)`);
      }
    }

    // Step C: If band is still too wide, adjust target rounds within [5, 10].
    // Band scales linearly with target rounds (band ≈ rounds × CV_party × 2.563).
    // To satisfy band ≤ bandTarget: maxRounds = bandTarget / (CV_party × 2.563).
    if (actualBand > bandTarget + 0.05) {
      const cvPartyEst      = actualBand / (Math.max(1, targetRound) * 2.563);
      const maxRoundsForBand = cvPartyEst > 0
        ? Math.floor(bandTarget / (cvPartyEst * 2.563))
        : targetRound;
      const newTarget = clamp(maxRoundsForBand, 5, 10);
      if (newTarget !== targetRound) {
        const oldTarget = targetRound;
        targetRound     = newTarget;
        state.pacing_rounds = newTarget;
        tunedHp    = hpSearch(targetRound);
        actualBand = getBand(tunedHp);
        changes.push(`Target rounds: ${oldTarget} → ${newTarget} (needed for band ≤${bandTarget.toFixed(1)}R)`);
      }
      // If still over band even at 5 rounds, leave as-is and report it — it's a physics limit.
    }

    // ── Stage 4: TPK cap check ─────────────────────────────────────────────
    let capAdjusted = false;
    const atTuned = runSim(tunedHp, quickTrials);
    if (atTuned && atTuned.tpk > tpkCap + 1e-9) {
      const oldMult = bossDprMultiplier(state);
      const atZeroPressure = runEncounterMc({
        boss_hp: Math.max(1, Math.round(tunedHp)),
        enc_trials: quickTrials,
        boss_dpr_mult: 0,
      });

      if (!atZeroPressure.error && atZeroPressure.tpkProb <= tpkCap + 1e-9 && oldMult > 0) {
        let multLo = 0;
        let multHi = oldMult;
        for (let i = 0; i < 18; i += 1) {
          const mid = (multLo + multHi) / 2;
          const r = runEncounterMc({
            boss_hp: Math.max(1, Math.round(tunedHp)),
            enc_trials: quickTrials,
            boss_dpr_mult: mid,
          });
          if (!r.error && r.tpkProb <= tpkCap + 1e-9) {
            multLo = mid;
          } else {
            multHi = mid;
          }
        }

        const newMult = round2(clamp(multLo, 0, 20));
        if (newMult < oldMult - 0.01) {
          state.boss_dpr_mult = newMult;
          tunedHp = hpSearch(targetRound);
          capAdjusted = true;
          changes.push(`Boss DPR mult: ${oldMult.toFixed(2)}x -> ${newMult.toFixed(2)}x (TPK cap)`);
        }
      }

      const afterPressure = runSim(tunedHp, quickTrials);
      const atOne = afterPressure && afterPressure.tpk > tpkCap + 1e-9 ? runSim(1, quickTrials) : null;
      if (atOne && atOne.tpk <= tpkCap + 1e-9) {
        let capLo = 1, capHi = tunedHp;
        while (capLo < capHi) {
          const midHp = Math.floor((capLo + capHi + 1) / 2);
          const r = runSim(midHp, quickTrials);
          if (!r || r.tpk <= tpkCap + 1e-9) { capLo = midHp; } else { capHi = midHp - 1; }
        }
        tunedHp     = capLo;
        capAdjusted = true;
      }
    }

    if (Math.abs(tunedHp - Math.round(oldBossHp)) > 0) {
      changes.push(`Boss HP: ${Math.round(oldBossHp)} → ${tunedHp}`);
    }
    if (targetRound !== origTarget) {
      // Already pushed above; also update the UI control.
    }
    state.boss_hp    = tunedHp;
    state.enc_trials = originalTrials;
    syncControlsFromState();
    persistState();
    refreshReport();

    // ── Stage 5: Final render + summary ────────────────────────────────────
    const finalMetrics = runEncounterAndRender();
    if (finalMetrics && !finalMetrics.error) {
      const finiteTtk = finalMetrics.ttk.filter(v => Number.isFinite(v));
      const finalMed  = finiteTtk.length ? percentile(finiteTtk, 50).toFixed(1) : "N/A";
      const finalP10  = finiteTtk.length ? percentile(finiteTtk, 10).toFixed(1) : "N/A";
      const finalP90  = finiteTtk.length ? percentile(finiteTtk, 90).toFixed(1) : "N/A";
      const finalBand = finiteTtk.length
        ? (percentile(finiteTtk, 90) - percentile(finiteTtk, 10)).toFixed(1) : "N/A";
      const finalKr   = (100 * finiteTtk.filter(v => v <= targetRound).length
        / Math.max(1, finalMetrics.ttk.length)).toFixed(1);
      const finalTpk  = finalMetrics.tpkProb;

      if (changes.length === 0) {
        setStatus("Auto-tune: parameters already optimal.", 3500);
        return;
      }

      const tpkLine = finalTpk > tpkCap + 1e-9
        ? `\n⚠ TPK ${(100 * finalTpk).toFixed(1)}% exceeds cap — reduce boss damage.`
        : `\nTPK: ${(100 * finalTpk).toFixed(1)}% ✓`;
      const krWarn = parseFloat(finalKr) < killRateTarget * 100 - 2
        ? `\n⚠ Kill rate by round ${targetRound}: ${finalKr}% (below ${(killRateTarget * 100).toFixed(0)}% target)`
        : `\nKill rate by round ${targetRound}: ${finalKr}% ✓`;
      const capLine  = capAdjusted ? "\n(Boss pressure constrained to keep TPK within cap.)" : "";
      const bandWarn = parseFloat(finalBand) > bandTarget + 0.5
        ? `\n⚠ Band ${finalBand}R exceeds target ${bandTarget}R — party damage variance is a physics limit.`
        : "";

      alert(
        `Auto-tune complete — ${changes.length} change(s):\n` +
        changes.map(c => `  • ${c}`).join("\n") +
        `\n\nMedian TTK: ${finalMed}R  |  p10–p90: ${finalP10}–${finalP90}R  |  Band: ${finalBand}R` +
        krWarn + tpkLine + capLine + bandWarn
      );
      setStatus(`Auto-tune complete. ${changes.length} parameter(s) updated.`, 5000);
    } else {
      setStatus("Auto-tune completed with partial results.", 4200);
    }
  }

  async function autoTuneAll() {
    setStatus("Auto-tuning encounter... solving HP and TPK together.", 0);

    const originalHp = Math.max(1, safeInt(state.boss_hp, 150));
    const originalMult = bossDprMultiplier(state);
    const originalTrials = Math.max(1000, safeInt(state.enc_trials, 10000));
    const quickTrials = Math.max(2500, Math.min(12000, Math.floor(originalTrials * 0.35)));
    const targetRound = clamp(safeInt(state.pacing_rounds, 5), 5, 10);
    const tpkCap = clamp(safeFloat(state.tune_tpk_cap, 0.05), 0, 1);
    const bandTarget = Math.max(0.5, safeFloat(state.tune_band_max, 3.0));
    const wipeTarget = targetRound + 1.5;
    const wipeFloor = targetRound + 0.75;
    const changes = [];
    const runner = createEncounterWorkerRunner();

    try {
    let pacingResult = computePacingResult();
    const minionCount = Math.max(0, safeInt(state.minion_count, 0));
    if (minionCount > 0 && Boolean(state.minion_replenish)) {
      const mph = pacingResult.minionPhaseInfo;
      if (!mph || mph.partyDprVsMinions <= 0) {
        alert(`Impossible encounter: party expected DPR against minion AC ${state.minion_ac} is 0.\nReduce minion AC, remove replenish, or add party accuracy/damage.`);
        setStatus("Auto-tune aborted: minions are unkillable.", 4000);
        return;
      }

      const newHp = Math.max(1, Math.floor((mph.partyDprVsMinions * 0.8) / minionCount));
      if (newHp < Math.round(state.minion_hp)) {
        changes.push(`Minion HP: ${Math.round(state.minion_hp)} -> ${newHp}`);
        state.minion_hp = newHp;
        pacingResult = computePacingResult();
      }
    }

    const summarize = (metrics, hp, mult, tgtRound, wipeByGrace) => {
      if (!metrics || metrics.error) return null;
      const finite = metrics.ttk.filter((v) => Number.isFinite(v));
      const median = percentile(metrics.ttk, 50);
      const p10 = finite.length ? percentile(finite, 10) : Infinity;
      const p90 = finite.length ? percentile(finite, 90) : Infinity;
      const band = Number.isFinite(p10) && Number.isFinite(p90) ? p90 - p10 : Infinity;
      const killByTarget = metrics.ttk.filter((v) => Number.isFinite(v) && v <= tgtRound).length / Math.max(1, metrics.ttk.length);
      const projectedWipe = projectedPartyWipeRound(mult, tgtRound);
      return {
        hp: Math.max(1, Math.round(hp)),
        mult: clamp(safeFloat(mult, 1), 0, 20),
        metrics,
        median,
        finite,
        p10,
        p90,
        band,
        tpk: metrics.tpkProb,
        killByTarget,
        projectedWipe,
        wipeByGrace,
      };
    };

    const runAt = async (hp, mult, trials, tgtRound) => {
      const trialCount = Math.max(1000, Math.round(trials));
      const tunedMult = clamp(safeFloat(mult, 1), 0, 20);
      const metrics = await runner.run({
        boss_hp: Math.max(1, Math.round(hp)),
        boss_dpr_mult: tunedMult,
        enc_trials: trialCount,
        enc_max_rounds: Math.max(safeInt(state.enc_max_rounds, 12), tgtRound + 3),
      });
      const wipeMetrics = await runner.run({
        boss_hp: 1000000000,
        boss_dpr_mult: tunedMult,
        enc_trials: Math.max(500, Math.floor(trialCount * 0.35)),
        enc_max_rounds: Math.max(1, Math.ceil(tgtRound + 2)),
        tight_pacing_enabled: false,
      });
      const wipeByGrace = wipeMetrics && !wipeMetrics.error ? clamp(safeFloat(wipeMetrics.tpkProb, 0), 0, 1) : 0;
      return summarize(metrics, hp, tunedMult, tgtRound, wipeByGrace);
    };

    const medianDistance = (median, tgtRound) => Number.isFinite(median) ? Math.abs(median - tgtRound) : 999;
    const hpScore = (result, tgtRound) => {
      if (!result) return Infinity;
      return medianDistance(result.median, tgtRound)
        + Math.abs(result.killByTarget - 0.5) * 0.35
        + Math.max(0, result.tpk - tpkCap) * 100;
    };

    const tuneHpForMult = async (mult, tgtRound, trials) => {
      const low = await runAt(1, mult, trials, tgtRound);
      if (!low) return null;
      if (low.median > tgtRound || !Number.isFinite(low.median)) {
        return { ...low, infeasible: "too-lethal-at-1hp" };
      }

      const deterministicHp = Math.max(1, safeInt(pacingResult.recommendedBossHp, originalHp));
      let hi = Math.max(10, originalHp, deterministicHp);
      let high = await runAt(hi, mult, trials, tgtRound);
      let guard = 0;
      while (high && high.median < tgtRound && guard < 12) {
        hi = Math.ceil(hi * 1.65 + 10);
        high = await runAt(hi, mult, trials, tgtRound);
        guard += 1;
      }
      if (!high) return null;

      let best = hpScore(low, tgtRound) <= hpScore(high, tgtRound) ? low : high;
      let loHp = 1;
      let hiHp = hi;

      for (let i = 0; i < 14 && hiHp - loHp > 1; i += 1) {
        const mid = Math.max(1, Math.floor((loHp + hiHp) / 2));
        const r = await runAt(mid, mult, trials, tgtRound);
        if (!r) return null;
        if (hpScore(r, tgtRound) < hpScore(best, tgtRound)) best = r;
        if (r.median < tgtRound) loHp = mid + 1;
        else hiHp = mid;
      }

      for (let hp = Math.max(1, best.hp - 3); hp <= best.hp + 3; hp += 1) {
        const r = await runAt(hp, mult, Math.max(1000, Math.floor(trials * 0.6)), tgtRound);
        if (r && hpScore(r, tgtRound) < hpScore(best, tgtRound)) best = r;
      }

      return best;
    };

    const tuneAtPressure = (mult, tgtRound, trials) => tuneHpForMult(mult, tgtRound, trials);
    const pressureScore = (result, tgtRound) => {
      if (!result || result.infeasible) return Infinity;
      const projectedWipe = result.projectedWipe;
      const wipePenalty = Number.isFinite(projectedWipe)
        ? Math.abs(projectedWipe - wipeTarget) * 8 + Math.max(0, wipeFloor - projectedWipe) * 120
        : 80;
      const graceWipePenalty = Math.abs(result.wipeByGrace - 0.5) * 45 + Math.max(0, 0.35 - result.wipeByGrace) * 90;
      const tpkPenalty = result.tpk > tpkCap ? (result.tpk - tpkCap) * 1200 : 0;
      const medianPenalty = medianDistance(result.median, tgtRound) * 25;
      const bandPenalty = Number.isFinite(result.band) ? Math.max(0, result.band - bandTarget) * 0.4 : 10;
      return medianPenalty + wipePenalty + graceWipePenalty + tpkPenalty + bandPenalty;
    };

    const seedMult = clamp(
      Number.isFinite(pacingResult.targetBossDprMult) && pacingResult.targetBossDprMult > 0
        ? pacingResult.targetBossDprMult
        : originalMult,
      0,
      20
    );

    setStatus("Auto-tuning encounter... bracketing boss pressure.", 0);
    let best = null;
    const remember = (result) => {
      if (!result || result.infeasible) return;
      if (!best || pressureScore(result, targetRound) < pressureScore(best, targetRound)) best = result;
    };

    const zero = await tuneAtPressure(0, targetRound, quickTrials);
    remember(zero);
    if (zero && zero.infeasible) {
      alert("Impossible encounter: even at zero boss damage and 1 boss HP, the party cannot reach the target pacing. Check replenishing minions, party DPR, resistance, or max rounds.");
      setStatus("Auto-tune aborted: target pacing is structurally impossible.", 5000);
      return;
    }

    const seed = await tuneAtPressure(seedMult, targetRound, quickTrials);
    remember(seed);

    let loMult = 0;
    let hiMult = Math.max(0.25, seedMult);
    if (seed && !seed.infeasible && seed.tpk <= tpkCap && seed.wipeByGrace < 0.45) {
      loMult = seed.mult;
      hiMult = Math.max(seed.mult * 1.5 + 0.25, 0.5);
      for (let i = 0; i < 8 && hiMult <= 20; i += 1) {
        const r = await tuneAtPressure(hiMult, targetRound, quickTrials);
        remember(r);
        if (!r || r.infeasible || r.tpk >= tpkCap || r.wipeByGrace >= 0.45) break;
        loMult = hiMult;
        hiMult = hiMult * 1.5 + 0.25;
      }
    }

    setStatus("Auto-tuning encounter... binary searching TPK-safe pressure.", 0);
    for (let i = 0; i < 8; i += 1) {
      const mid = (loMult + hiMult) / 2;
      const r = await tuneAtPressure(mid, targetRound, quickTrials);
      remember(r);
      if (!r || r.infeasible || r.tpk > tpkCap || (Number.isFinite(r.projectedWipe) && r.projectedWipe < wipeFloor) || r.wipeByGrace > 0.6) hiMult = mid;
      else loMult = mid;
    }

    if (!best) {
      alert("Auto-tune could not find a valid HP/DPR solution. Check that the party can damage the boss and that max rounds is at least target + 3.");
      setStatus("Auto-tune failed.", 4200);
      return;
    }

    setStatus("Auto-tuning encounter... final verification.", 0);
    let final = await tuneAtPressure(best.mult, targetRound, Math.max(3000, Math.floor(originalTrials * 0.6))) || best;
    let safetyPasses = 0;
    while (final && final.tpk > tpkCap && final.mult > 0.01 && safetyPasses < 4) {
      final = await tuneAtPressure(final.mult * 0.85, targetRound, Math.max(3000, Math.floor(originalTrials * 0.5))) || final;
      safetyPasses += 1;
    }

    const tunedHp = Math.max(1, Math.round(final.hp));
    const tunedMult = round2(clamp(final.mult, 0, 20));
    if (tunedHp !== originalHp) changes.push(`Boss HP: ${originalHp} -> ${tunedHp}`);
    if (Math.abs(tunedMult - originalMult) > 0.01) changes.push(`Boss DPR mult: ${originalMult.toFixed(2)}x -> ${tunedMult.toFixed(2)}x`);

    state.pacing_rounds = targetRound;
    state.boss_hp = tunedHp;
    state.boss_dpr_mult = tunedMult;
    state.enc_max_rounds = Math.max(safeInt(state.enc_max_rounds, 12), targetRound + 3);
    state.enc_trials = originalTrials;
    syncControlsFromState();
    persistState();
    refreshReport();

    const finalMetrics = await runEncounterAndRender();
    if (!finalMetrics || finalMetrics.error) {
      setStatus("Auto-tune applied values, but final simulation failed.", 4200);
      return;
    }

    const finiteTtk = finalMetrics.ttk.filter((v) => Number.isFinite(v));
    const finalMed = percentile(finalMetrics.ttk, 50);
    const finalP10 = finiteTtk.length ? percentile(finiteTtk, 10) : Infinity;
    const finalP90 = finiteTtk.length ? percentile(finiteTtk, 90) : Infinity;
    const finalBand = Number.isFinite(finalP10) && Number.isFinite(finalP90) ? finalP90 - finalP10 : Infinity;
    const finalTpk = finalMetrics.tpkProb;
    const finalKillByTarget = finalMetrics.ttk.filter((v) => Number.isFinite(v) && v <= targetRound).length / Math.max(1, finalMetrics.ttk.length);
    const finalProjectedWipe = projectedPartyWipeRound(state.boss_dpr_mult, targetRound);
    const finalWipeMetrics = await runner.run({
      boss_hp: 1000000000,
      boss_dpr_mult: state.boss_dpr_mult,
      enc_trials: Math.max(1000, Math.floor(originalTrials * 0.25)),
      enc_max_rounds: targetRound + 2,
      tight_pacing_enabled: false,
    });
    const finalWipeByGrace = finalWipeMetrics && !finalWipeMetrics.error ? clamp(safeFloat(finalWipeMetrics.tpkProb, 0), 0, 1) : 0;
    const lines = changes.length ? changes.map((c) => `  - ${c}`).join("\n") : "  - Already at tuned values";
    const tpkLine = finalTpk <= tpkCap
      ? `TPK ${(100 * finalTpk).toFixed(1)}% <= cap ${(100 * tpkCap).toFixed(1)}%`
      : `TPK ${(100 * finalTpk).toFixed(1)}% exceeds cap ${(100 * tpkCap).toFixed(1)}%; reduce boss/minion damage or target a shorter fight`;
    const bandLine = finalBand > bandTarget + 0.5
      ? `\nBand ${finalBand.toFixed(1)}R exceeds ${bandTarget.toFixed(1)}R; damage variance is too high for tighter pacing without changing mechanics.`
      : "";

    alert(
      `Auto-tune complete:\n${lines}\n\n` +
      `Target: ${targetRound}R | Median: ${Number.isFinite(finalMed) ? finalMed.toFixed(1) : "inf"}R | ` +
      `p10-p90: ${Number.isFinite(finalP10) ? finalP10.toFixed(1) : "inf"}-${Number.isFinite(finalP90) ? finalP90.toFixed(1) : "inf"}R\n` +
      `${tpkLine} | Projected party wipe: ${Number.isFinite(finalProjectedWipe) ? finalProjectedWipe.toFixed(1) : "inf"}R | Wipe by R${targetRound + 2}: ${(100 * finalWipeByGrace).toFixed(1)}% | Kill by target: ${(100 * finalKillByTarget).toFixed(1)}%` +
      bandLine +
      tightPacingAdviceText()
    );
    setStatus(`Auto-tune complete. HP=${state.boss_hp}, DPR=${state.boss_dpr_mult.toFixed(2)}x.`, 5000);
    } finally {
      runner.close();
    }
  }

  function runMcSim(pcRow, attacks, opts) {
    const mechanics = phaseMechanicsEnabledFromTable(opts.phase_table);
    const trials = Math.max(1000, safeInt(opts.mc_trials, 10000));
    const rounds = Math.max(1, safeInt(opts.mc_rounds, 3));
    const ac = Math.max(1, safeInt(pcRow.AC, 10));
    const saveBonuses = {};
    for (const key of SAVE_KEYS) {
      saveBonuses[key] = safeInt(pcRow[key], 0);
    }
    const thpAvg = Math.max(0, averageDamage(opts.thp_expr || "0"));
    const spreadTargets = Math.max(1, safeInt(opts.spread_targets, 1));
    const dprMult = bossDprMultiplier(opts);

    const totalDamage = new Array(trials).fill(0);
    const riderRemaining = new Array(trials).fill(0);
    const attackUsesRemaining = Array.from({ length: trials }, () => attacks.map((atk) => attackUseCapacity(atk)));

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

        for (let a = 0; a < attacks.length; a += 1) {
          const atk = attacks[a];
          if (atk.uses_per_round <= 0) continue;

          const availableUses = Math.min(atk.uses_per_round, attackUsesRemaining[i][a]);
          for (let use = 0; use < availableUses; use += 1) {
            attackUsesRemaining[i][a] -= 1;
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

            const isCrit = roll >= (atk.crit_threshold ?? 20);
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
            roundDamage += rollDamageOne(opts.lair_formula || damageFormulaForAverage(opts.lair_avg, 6));
          }
        }

        if (opts.rech_enabled) {
          const pRech = parseRecharge(opts.recharge_text || "5-6");
          const pTarget = Math.min(1, safeFloat(opts.rech_targets, 1) / spreadTargets);
          if (Math.random() < pRech && Math.random() < pTarget) {
            roundDamage += rollDamageOne(opts.rech_formula || damageFormulaForAverage(opts.rech_avg, 6));
          }
        }

        for (const mech of mechanicsForRound(mechanics, rnd)) {
          const pTarget = Math.min(1, safeFloat(mech.targets, 1) / spreadTargets);
          if (Math.random() < pTarget) {
            roundDamage += rollMechanicDamageVsPc(mech, pcRow, currentMode, currentAc, saveBonuses);
          }
        }

        totalDamage[i] += Math.max(0, roundDamage * dprMult - thpAvg);

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
    const mechanics = phaseMechanicsEnabledFromTable(opts.phase_table);
    const hasMinionDamage = Math.max(0, safeInt(opts.minion_count, 0)) > 0
      && (Boolean(opts.minion_atk_enabled) || Boolean(opts.minion_save_enabled));
    if (!attacks.length && !mechanics.length && !opts.lair_enabled && !opts.rech_enabled && !hasMinionDamage) {
      return { error: "Enable at least one attack, round mechanic, lair action, recharge power, or damaging minion pack." };
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
    const dprMult = bossDprMultiplier(opts);
    const bossMaxHp = Math.max(1, safeFloat(opts.boss_hp, 150));
    const tightPacing = Boolean(opts.tight_pacing_enabled);
    const tightTargetRound = clamp(safeInt(opts.pacing_rounds, 5), 5, 10);
    const tightCapPct = clamp(safeFloat(opts.tight_cap_pct, 0.24), 0.10, 0.40);
    const tightCapResources = Math.max(0, safeInt(opts.tight_cap_resources, 3));
    const tightAntiSlogMult = clamp(safeFloat(opts.tight_anti_slog_mult, 1.35), 1.0, 2.5);

    const effList = effPartyDprs(useNova);
    if (!effList.length) {
      return { error: "No DPR rows available. Fill DPR input table first." };
    }

    const partyNames = party.map((pc, i) => String(pc.Name || `PC${i + 1}`));
    const P = partyNames.length;
    const effByName = new Map(effList.map((r) => [String(r[0]), safeFloat(r[1], 0)]));
    const effMeans = partyNames.map((name) => effByName.get(name) || 0);

    const bossHp = new Array(trials).fill(bossMaxHp);
    const bossDamageThisRound = new Array(trials).fill(0);
    const tightResourcesLeft = new Array(trials).fill(tightCapResources);
    const pcsHp = Array.from({ length: trials }, () => party.map((pc) => Math.max(1, safeFloat(pc.HP, 1))));
    const pcsAlive = Array.from({ length: trials }, () => party.map(() => true));

    const pcAc = party.map((pc) => Math.max(1, safeInt(pc.AC, 10)));
    const pcSaves = party.map((pc) => SAVE_KEYS.map((k) => safeInt(pc[k], 0)));
    const saveIndex = { STR: 0, DEX: 1, CON: 2, INT: 3, WIS: 4, CHA: 5 };

    const riderRem = Array.from({ length: trials }, () => new Array(P).fill(0));
    const attackUsesRemaining = Array.from({ length: trials }, () => attacks.map((atk) => attackUseCapacity(atk)));

    const ttk = new Array(trials).fill(Number.POSITIVE_INFINITY);
    const tpkFlags = new Array(trials).fill(false);
    const pcsDownAtVictory = new Array(trials).fill(0);

    const bossFirstFlags = buildBossFirstFlags(trials, initMode);

    // ── Minion pack setup ──────────────────────────────────────────────────
    const mcMinionCount   = Math.max(0, safeInt(opts.minion_count, 0));
    const mcMinionHpEach  = Math.max(1, safeFloat(opts.minion_hp, 15));
    const mcMinionAc      = Math.max(1, safeInt(opts.minion_ac, 14));
    const mcHasMinionPack = mcMinionCount > 0;
    const mcMinionHitRatio = mcHasMinionPack
      ? minionHitRatio(mcMinionAc, Math.max(1, safeInt(opts.boss_ac, 16)))
      : 1;
    const mcMinionReplenish = mcHasMinionPack && Boolean(opts.minion_replenish);
    const mcMinionPoolHp = mcHasMinionPack
      ? new Array(trials).fill(mcMinionCount * mcMinionHpEach)
      : null;
    const mcMinionHasDmg = mcHasMinionPack
      && (Boolean(opts.minion_atk_enabled) || Boolean(opts.minion_save_enabled));
    // ──────────────────────────────────────────────────────────────────────

    const applyPartyPhase = (t, roundNumber) => {
      let bossRawDamage = 0;

      if (useNova) {
        const bossAc = Math.max(1, safeInt(opts.boss_ac, 16));
        for (let j = 0; j < P; j += 1) {
          if (!pcsAlive[t][j]) continue;
          const minionHpLeft = mcHasMinionPack ? mcMinionPoolHp[t] : 0;
          const split = sampleNovaActionSplit(partyNames[j], minionHpLeft, bossAc, mcMinionAc);
          if (mcHasMinionPack) {
            mcMinionPoolHp[t] = split.minionPoolHp;
          }
          bossRawDamage += split.bossRaw;
        }
      } else {
        let rawPartyDpr = 0;
        for (let j = 0; j < P; j += 1) {
          if (!pcsAlive[t][j]) continue;
          rawPartyDpr += gammaRng(effMeans[j], dprCv);
        }

        if (mcHasMinionPack && mcMinionPoolHp[t] > 0) {
          const dprToMinions = rawPartyDpr * mcMinionHitRatio;
          const prevHp = mcMinionPoolHp[t];
          mcMinionPoolHp[t] = Math.max(0, prevHp - dprToMinions);
          if (mcMinionPoolHp[t] === 0 && dprToMinions > prevHp) {
            bossRawDamage += rawPartyDpr * ((dprToMinions - prevHp) / dprToMinions);
          }
        } else {
          bossRawDamage += rawPartyDpr;
        }
      }

      if (bossRawDamage > 0) {
        let bossDamage = Math.max(0, bossRawDamage / resist - regen);
        if (tightPacing && roundNumber > tightTargetRound && tightAntiSlogMult > 1) {
          bossDamage *= Math.pow(tightAntiSlogMult, roundNumber - tightTargetRound);
        }
        if (tightPacing && roundNumber <= tightTargetRound && tightResourcesLeft[t] > 0) {
          const roundCap = bossMaxHp * tightCapPct;
          const remainingCap = Math.max(0, roundCap - bossDamageThisRound[t]);
          if (bossDamage > remainingCap) {
            bossDamage = remainingCap;
            tightResourcesLeft[t] -= 1;
          }
        }
        bossDamageThisRound[t] += bossDamage;
        bossHp[t] -= bossDamage;
      }

      if (bossHp[t] <= 0 && !Number.isFinite(ttk[t])) {
        ttk[t] = roundNumber;
        pcsDownAtVictory[t] = pcsAlive[t].reduce((acc, alive) => acc + (alive ? 0 : 1), 0);
      }
    };

    for (let rnd = 1; rnd <= maxRounds; rnd += 1) {
      bossDamageThisRound.fill(0);
      // Replenishing minions: reset each trial's pool at the top of every round.
      if (mcMinionReplenish) {
        for (let t = 0; t < trials; t += 1) {
          if (!Number.isFinite(ttk[t])) {
            mcMinionPoolHp[t] = mcMinionCount * mcMinionHpEach;
          }
        }
      }

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
        applyPartyPhase(t, rnd);
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

        for (let a = 0; a < attacks.length; a += 1) {
          const atk = attacks[a];
          if (atk.uses_per_round <= 0) continue;

          const availableUses = Math.min(atk.uses_per_round, attackUsesRemaining[t][a]);
          for (let use = 0; use < availableUses; use += 1) {
            attackUsesRemaining[t][a] -= 1;
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

              const isCrit = r >= (atk.crit_threshold ?? 20);
              const isHit = isCrit || (r !== 1 && r + atk.attack_bonus >= currentAc[target]);

              if (isHit) {
                rawDealt = isCrit ? rollDamageCrunchyCritOne(atk.damage_expr) : rollDamageOne(atk.damage_expr);
                if (atk.is_melee || !opts.rider_melee_only) {
                  riderTrig[target] = true;
                }
              }
            }

            applyDamage(target, Math.max(0, rawDealt * dprMult));
          }
        }

        for (const mech of mechanicsForRound(mechanics, rnd)) {
          aliveNow = aliveIndicesForTrial(pcsAlive[t]);
          if (!aliveNow.length) break;

          const targetCount = Math.min(Math.max(1, safeInt(mech.targets, 1)), aliveNow.length);
          const targets = sampleWithoutReplacement(aliveNow, targetCount);
          for (const target of targets) {
            const rawDealt = rollMechanicDamageForTarget(mech, target, currentAc, currentMode, pcSaves, saveIndex);
            applyDamage(target, Math.max(0, rawDealt * dprMult));
          }
        }

        if (opts.lair_enabled && rnd % Math.max(1, safeInt(opts.lair_every_n, 1)) === 0) {
          aliveNow = aliveIndicesForTrial(pcsAlive[t]);
          if (aliveNow.length) {
            const L = Math.min(Math.max(1, safeInt(opts.lair_targets, 1)), aliveNow.length);
            const targets = sampleWithoutReplacement(aliveNow, L);
            for (const idx of targets) {
              applyDamage(idx, Math.max(0, rollDamageOne(opts.lair_formula || damageFormulaForAverage(opts.lair_avg, 6)) * dprMult));
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
                applyDamage(idx, Math.max(0, rollDamageOne(opts.rech_formula || damageFormulaForAverage(opts.rech_avg, 6)) * dprMult));
              }
            }
          }
        }

        // Minion pack attacks: living minions deal damage to each alive PC based on their AC/saves.
        if (mcHasMinionPack && mcMinionHasDmg && mcMinionPoolHp[t] > 0) {
          aliveNow = aliveIndicesForTrial(pcsAlive[t]);
          if (aliveNow.length) {
            const minionsLeft = Math.ceil(mcMinionPoolHp[t] / mcMinionHpEach);
            for (let m = 0; m < minionsLeft; m += 1) {
              aliveNow = aliveIndicesForTrial(pcsAlive[t]);
              if (!aliveNow.length) break;
              const idx = aliveNow[randomInt(0, aliveNow.length - 1)];
              const raw = sampleMinionDamageVsPc(idx, pcAc, pcSaves, saveIndex, opts);
              applyDamage(idx, Math.max(0, raw * dprMult));
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
        applyPartyPhase(t, rnd);
      }
    }

    const times = [];
    for (let t = 0; t <= maxRounds; t += 1) times.push(t);
    const finiteTtk = ttk.filter((v) => Number.isFinite(v));
    const nKills = finiteTtk.length;
    const survivalCurve = times.map((round) =>
      nKills > 0 ? finiteTtk.filter((v) => v > round).length / nKills : 0
    );

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

  // ── Pacing Calculator ─────────────────────────────────────────────────────

  let _lastPacingResult = null;

  function computePacingResult() {
    const targetRounds = clamp(safeInt(state.pacing_rounds, 5), 5, 10);

    // Effective party DPR on the boss
    const useNova = Boolean(state.enc_use_nova);
    const effList = effPartyDprs(useNova);
    const totalPartyDpr = effList.reduce((acc, r) => acc + safeFloat(r[1], 0), 0);
    const resistFactor = Math.max(1e-6, safeFloat(state.resist_factor, 1.0));
    const bossRegen = Math.max(0, safeFloat(state.boss_regen, 0.0));
    const effectivePartyDprOnBoss = Math.max(0, totalPartyDpr / resistFactor - bossRegen);

    // Recommended boss HP — minion-aware
    const minionPhaseInfo = computeMinionPhase(totalPartyDpr, useNova);
    let recommendedBossHp;
    if (!minionPhaseInfo) {
      recommendedBossHp = effectivePartyDprOnBoss > 0
        ? Math.max(1, Math.round(effectivePartyDprOnBoss * targetRounds))
        : 0;
    } else if (minionPhaseInfo.replenishing && !minionPhaseInfo.canClearPerRound) {
      // Boss is permanently screened — party can never reach it.
      recommendedBossHp = 0;
    } else if (minionPhaseInfo.replenishing && minionPhaseInfo.canClearPerRound) {
      // Party clears minions every round; only overflow reaches boss.
      recommendedBossHp = minionPhaseInfo.overflowBossPerRound > 0
        ? Math.max(1, Math.round(minionPhaseInfo.overflowBossPerRound * targetRounds))
        : 0;
    } else {
      // Non-replenishing: some rounds absorb minions, then boss is exposed.
      const minionRounds = Math.min(minionPhaseInfo.minionPhaseRounds, targetRounds);
      const bossRoundsAfter = Math.max(0, targetRounds - minionRounds);
      const totalBossDmg = minionPhaseInfo.totalBossDmgDuringPhase
        + effectivePartyDprOnBoss * bossRoundsAfter;
      recommendedBossHp = totalBossDmg > 0 ? Math.max(1, Math.round(totalBossDmg)) : 0;
    }

    // Per-PC TTD at current boss DPR
    const attacks = attacksEnabledFromTable(state.attacks_table);
    const mechanics = phaseMechanicsEnabledFromTable(state.phase_table);
    const party = state.party_table.filter((r) => String(r.Name || "").trim().length > 0);
    const thpAvg = Math.max(0, averageDamage(state.thp_expr || "0"));
    const spread = Math.max(1, safeInt(state.spread_targets, 1));
    const lairRechDpr = lairPerTargetDpr(state, party.length || 1) + rechargePerTargetDpr(state, party.length || 1);
    const dprMult = bossDprMultiplier(state);

    let firstDownRound = Infinity;
    let lastDownRound = 0;
    let targetAttackDpr = 0;
    let targetBossDprMult = null;
    let toughestPcName = null;
    const wipeBudgetRound = targetRounds * 1.35;
    const pcRows = [];

    for (const pc of party) {
      const pcHp = Math.max(1, safeInt(pc.HP, 1));
      const rawAttackDpr = perRoundDprVsPc(pc, state.mode_select || "normal", attacks, targetRounds);
      const rawPhaseDpr = phaseMechanicsPerTargetDpr(pc, state.mode_select || "normal", mechanics, party.length, targetRounds);
      const pcMinionDpr = minionDprVsPc(pc);
      const rawIncomingPerPc = rawAttackDpr / spread + lairRechDpr + pcMinionDpr + rawPhaseDpr;
      const totalDprPerPc = rawIncomingPerPc * dprMult;
      const netDprPerPc = Math.max(0, totalDprPerPc - thpAvg);
      const pcTtd = netDprPerPc > 0 ? pcHp / netDprPerPc : Infinity;

      if (Number.isFinite(pcTtd)) {
        firstDownRound = Math.min(firstDownRound, pcTtd);
        lastDownRound = Math.max(lastDownRound, pcTtd);
      }

      // Target total boss pressure so deterministic party wipe lands after the
      // intended fight length. This keeps pressure meaningful without making
      // "boss dies on target round" equivalent to a deterministic TPK.
      const neededNetDpr = pcHp / wipeBudgetRound;
      const targetMultForPc = rawIncomingPerPc > 0
        ? Math.max(0, (neededNetDpr + thpAvg) / rawIncomingPerPc)
        : Infinity;
      const neededAttackDpr = Number.isFinite(targetMultForPc) ? rawIncomingPerPc * targetMultForPc : 0;
      if (Number.isFinite(targetMultForPc) && (targetBossDprMult == null || targetMultForPc > targetBossDprMult)) {
        targetBossDprMult = targetMultForPc;
        targetAttackDpr = neededAttackDpr;
        toughestPcName = pc.Name || "?";
      }

      pcRows.push({
        PC: pc.Name || "?",
        HP: pcHp,
        AC: safeInt(pc.AC, 10),
        "Base DPR/target": round2(rawIncomingPerPc),
        "Script DPR": round2(rawPhaseDpr),
        "Minion DPR": round2(pcMinionDpr * dprMult),
        "Scaled DPR/target": round2(totalDprPerPc),
        "Net DPR": round2(netDprPerPc),
        "Target Mult": Number.isFinite(targetMultForPc) ? `${targetMultForPc.toFixed(2)}x` : "N/A",
        "TTD (exact)": Number.isFinite(pcTtd) ? pcTtd.toFixed(1) : "∞",
        "TTD (rounds)": Number.isFinite(pcTtd) ? String(Math.ceil(pcTtd)) : "∞",
      });
    }

    return {
      targetRounds,
      recommendedBossHp,
      minionPhaseInfo,
      effectivePartyDprOnBoss,
      firstDownRound: Number.isFinite(firstDownRound) ? firstDownRound : null,
      lastDownRound: lastDownRound > 0 ? lastDownRound : null,
      targetAttackDpr: round2(targetAttackDpr),
      targetBossDprMult: targetBossDprMult == null ? null : round2(targetBossDprMult),
      bossDprMult: dprMult,
      toughestPcName,
      currentBossHp: safeFloat(state.boss_hp, 200),
      pcRows,
      lairRechDpr,
      thpAvg,
    };
  }

  function runComputePacing() {
    const result = computePacingResult();
    _lastPacingResult = result;
    renderPacingResult(result);
    updatePacingActionButtons(result);
    setStatus("Pacing computed.", 2000);
  }

  function projectedPartyWipeRound(mult, horizonRounds = 5) {
    const attacks = attacksEnabledFromTable(state.attacks_table);
    const mechanics = phaseMechanicsEnabledFromTable(state.phase_table);
    const party = state.party_table.filter((r) => String(r.Name || "").trim().length > 0);
    if (!party.length) return Infinity;

    const dprMult = clamp(safeFloat(mult, 1), 0, 20);
    const thpAvg = Math.max(0, averageDamage(state.thp_expr || "0"));
    const spread = Math.max(1, safeInt(state.spread_targets, 1));
    const rounds = Math.max(1, safeFloat(horizonRounds, 5));
    const lairRechDpr = lairPerTargetDpr(state, party.length || 1) + rechargePerTargetDpr(state, party.length || 1);

    let lastDown = 0;
    let anyFinite = false;
    for (const pc of party) {
      const pcHp = Math.max(1, safeInt(pc.HP, 1));
      const attackDpr = perRoundDprVsPc(pc, state.mode_select || "normal", attacks, rounds);
      const phaseDpr = phaseMechanicsPerTargetDpr(pc, state.mode_select || "normal", mechanics, party.length, rounds);
      const minionDpr = minionDprVsPc(pc);
      const rawIncoming = attackDpr / spread + phaseDpr + minionDpr + lairRechDpr;
      const net = Math.max(0, rawIncoming * dprMult - thpAvg);
      if (net <= 0) return Infinity;
      const ttd = pcHp / net;
      if (Number.isFinite(ttd)) {
        lastDown = Math.max(lastDown, ttd);
        anyFinite = true;
      }
    }

    return anyFinite ? lastDown : Infinity;
  }

  function projectedPartyWipeProb(mult, horizonRounds, trials) {
    const m = runEncounterMc({
      boss_hp: 1000000000,
      boss_dpr_mult: clamp(safeFloat(mult, 1), 0, 20),
      enc_trials: Math.max(500, Math.round(trials)),
      enc_max_rounds: Math.max(1, Math.ceil(safeFloat(horizonRounds, 7))),
      tight_pacing_enabled: false,
    });
    if (!m || m.error) return 0;
    return clamp(safeFloat(m.tpkProb, 0), 0, 1);
  }

  function createEncounterWorkerRunner() {
    if (IS_WORKER || typeof Worker === "undefined" || typeof window === "undefined") {
      return {
        run: async (overrides = {}) => {
          await yieldToBrowser();
          const result = runEncounterMc(overrides);
          await yieldToBrowser();
          return result;
        },
        close: () => {},
      };
    }

    let worker;
    try {
      worker = new Worker(new URL("app.js", window.location.href).toString());
    } catch (_) {
      return {
        run: async (overrides = {}) => {
          await yieldToBrowser();
          const result = runEncounterMc(overrides);
          await yieldToBrowser();
          return result;
        },
        close: () => {},
      };
    }

    let seq = 0;
    const pending = new Map();

    worker.onmessage = (event) => {
      const data = event.data || {};
      const entry = pending.get(data.id);
      if (!entry) return;
      pending.delete(data.id);
      if (data.error) entry.reject(new Error(data.error));
      else entry.resolve(data.result);
    };

    worker.onerror = (event) => {
      const err = new Error(event.message || "Encounter worker failed.");
      for (const entry of pending.values()) {
        entry.reject(err);
      }
      pending.clear();
    };

    return {
      run: (overrides = {}) => new Promise((resolve, reject) => {
        const id = ++seq;
        pending.set(id, { resolve, reject });
        worker.postMessage({
          id,
          type: "runEncounterMc",
          state: deepClone(state),
          overrides,
        });
      }),
      close: () => {
        for (const entry of pending.values()) {
          entry.reject(new Error("Encounter worker closed."));
        }
        pending.clear();
        worker.terminate();
      },
    };
  }

  function yieldToBrowser() {
    return new Promise((resolve) => setTimeout(resolve, 0));
  }

  function renderPhaseSection() {
    renderEditableTable({
      mount: els.phaseTable,
      columns: PHASE_COLUMNS,
      rows: state.phase_table,
      showRemove: true,
      onCellChange: (rowIndex, key, value) => {
        state.phase_table[rowIndex][key] = value;
        state.phase_table[rowIndex] = sanitizePhaseRow(state.phase_table[rowIndex]);
        persistState();
        renderPhaseSection();
        refreshReport();
      },
      onRemoveRow: (rowIndex) => {
        state.phase_table.splice(rowIndex, 1);
        persistState();
        renderPhaseSection();
        refreshReport();
      },
      emptyMessage: "No scripted round mechanics yet.",
    });
  }

  function updatePacingActionButtons(result) {
    void result;
  }

  function applyPacingHp() {
    if (!_lastPacingResult || _lastPacingResult.recommendedBossHp <= 0) return;
    state.boss_hp = _lastPacingResult.recommendedBossHp;
    syncControlsFromState();
    persistState();
    refreshEffTableFromMode();
    refreshReport();
    runComputePacing();
    setStatus(`Boss HP set to ${state.boss_hp}.`, 2500);
  }

  function applyPacingDpr() {
    if (!_lastPacingResult || !Number.isFinite(_lastPacingResult.targetBossDprMult)) return;
    state.boss_dpr_mult = clamp(_lastPacingResult.targetBossDprMult, 0, 20);
    syncControlsFromState();
    persistState();
    refreshReport();
    runComputePacing();
    setStatus(`Boss DPR multiplier set to ${state.boss_dpr_mult.toFixed(2)}x.`, 2500);
  }

  function applyBossKitMultiplier() {
    const mult = bossDprMultiplier(state);
    if (!Number.isFinite(mult) || mult <= 0 || Math.abs(mult - 1) < 0.005) {
      setStatus("No non-1x boss DPR multiplier to apply.", 2500);
      return;
    }

    state.attacks_table = state.attacks_table.map((row) => ({
      ...row,
      Damage: scaleDamageExpression(row.Damage, mult),
    })).map(sanitizeAttackRow);
    state.phase_table = state.phase_table.map((row) => ({
      ...row,
      Damage: scaleDamageExpression(row.Damage, mult),
    })).map(sanitizePhaseRow);

    state.lair_avg = round2(Math.max(0, safeFloat(state.lair_avg, 0) * mult));
    state.rech_avg = round2(Math.max(0, safeFloat(state.rech_avg, 0) * mult));
    state.lair_formula = damageFormulaForAverage(state.lair_avg, 6);
    state.rech_formula = damageFormulaForAverage(state.rech_avg, 6);
    state.boss_dpr_mult = 1;

    syncControlsFromState();
    persistState();
    renderAttackSection();
    renderPhaseSection();
    refreshReport();
    runComputePacing();
    setStatus(`Baked ${mult.toFixed(2)}x DPR into boss kit formulas and reset multiplier to 1x.`, 3500);
  }

  function renderPacingResult(r) {
    // Metric values
    setText(els.pacingBossHp, r.recommendedBossHp > 0 ? String(r.recommendedBossHp) : "N/A");
    {
      const mph = r.minionPhaseInfo;
      let hpSub;
      if (!mph) {
        hpSub = `party eff. DPR ${r.effectivePartyDprOnBoss.toFixed(1)} × ${r.targetRounds} rounds`;
      } else if (mph.replenishing && !mph.canClearPerRound) {
        hpSub = `boss permanently screened by replenishing minions`;
      } else if (mph.replenishing && mph.canClearPerRound) {
        hpSub = `${mph.overflowBossPerRound.toFixed(1)} overflow/round × ${r.targetRounds} rounds (minions cleared each round)`;
      } else {
        hpSub = `${mph.minionPhaseRounds}R minion phase + ${r.effectivePartyDprOnBoss.toFixed(1)} eff. DPR × ${Math.max(0, r.targetRounds - mph.minionPhaseRounds)} remaining rounds`;
      }
      setText(els.pacingBossHpSub, hpSub);
    }

    if (r.firstDownRound != null) {
      const fd = Math.ceil(r.firstDownRound);
      setText(els.pacingFirstDown, String(fd));
      applyMetricClass(els.pacingFirstDown, fd < r.targetRounds ? "danger" : fd > r.targetRounds * 1.5 ? "success" : "");
    } else {
      setText(els.pacingFirstDown, "∞");
      applyMetricClass(els.pacingFirstDown, "success");
    }

    if (r.lastDownRound != null) {
      const ld = Math.ceil(r.lastDownRound);
      setText(els.pacingPartyWipe, String(ld));
      applyMetricClass(els.pacingPartyWipe, ld < r.targetRounds ? "danger" : ld > r.targetRounds * 2 ? "success" : "warning");
    } else {
      setText(els.pacingPartyWipe, "∞");
      applyMetricClass(els.pacingPartyWipe, "success");
    }

    const bossDprMult = Number.isFinite(r.bossDprMult) ? r.bossDprMult : 1;
    setText(els.pacingTargetDpr, Number.isFinite(r.targetBossDprMult) ? `${r.targetBossDprMult.toFixed(2)}x` : "N/A");
    const currentDprHint = r.toughestPcName
      ? `current ${bossDprMult.toFixed(2)}x; ${r.toughestPcName} wipe budget ${(r.targetRounds * 1.35).toFixed(1)}R`
      : "no boss damage";
    setText(els.pacingTargetDprSub, currentDprHint);

    // Balance analysis
    renderPacingBalance(r);

    // Per-PC table
    renderResultTable(els.pacingTable, r.pcRows);
  }

  function renderPacingBalance(r) {
    if (!els.pacingBalance) return;
    const el = els.pacingBalance;
    const N = r.targetRounds;
    const bossHp = r.currentBossHp;
    const recHp = r.recommendedBossHp;
    const partyWipe = r.lastDownRound;
    const firstDown = r.firstDownRound;

    let cls = "bal-ok";
    let badge = "";
    let analysis = "";

    const mph = r.minionPhaseInfo;
    if (mph && mph.replenishing && !mph.canClearPerRound) {
      cls = "bal-hard";
      badge = `<span class="pacing-badge badge-hard">Boss Screened</span>`;
      analysis = `Replenishing minion pack permanently shields the boss — the party can't one-round the pack. ` +
        `The party needs more than ${round2(mph.partyDprVsMinions)} effective DPR vs minion AC to break through. ` +
        `Reduce minion count/HP or raise party DPR.`;
    } else if (r.effectivePartyDprOnBoss <= 0) {
      cls = "bal-hard";
      badge = `<span class="pacing-badge badge-hard">Cannot Kill Boss</span>`;
      analysis = "Party effective DPR on the boss is zero — check resistance factor, regen, and party DPR settings.";
    } else if (partyWipe != null && partyWipe < N * 0.7) {
      cls = "bal-hard";
      badge = `<span class="pacing-badge badge-hard">Very Hard</span>`;
      analysis = `Party wipes at round ${fmt1(partyWipe)} — well before the target ${N} rounds. ` +
        `Reduce boss DPR toward ${fmtMult(r.targetBossDprMult)} or raise party sustain.`;
    } else if (partyWipe != null && partyWipe < N) {
      cls = "bal-warn";
      badge = `<span class="pacing-badge badge-warn">Hard</span>`;
      analysis = `Party wipes at round ${fmt1(partyWipe)} — slightly before target. ` +
        `Use boss DPR multiplier ${fmtMult(r.targetBossDprMult)} for ${r.toughestPcName} to last the full ${N} rounds.`;
    } else if (partyWipe == null || !Number.isFinite(partyWipe) || partyWipe > N * 2.5) {
      cls = "bal-easy";
      badge = `<span class="pacing-badge badge-easy">Very Easy</span>`;
      analysis = `Boss dies in ${N} rounds and the party takes minimal casualties. ` +
        `To add tension, increase boss DPR to ${fmtMult(r.targetBossDprMult)}.`;
    } else if (partyWipe != null && partyWipe < N * 1.4) {
      cls = "bal-ok";
      badge = `<span class="pacing-badge badge-ok">Balanced</span>`;
      analysis = `Boss dies in ${N} rounds; party starts falling at round ${fmt1(firstDown)} and last PC at ${fmt1(partyWipe)}. ` +
        `Recommended boss HP: ${recHp} (current: ${Math.round(bossHp)}).`;
    } else {
      cls = "bal-easy";
      badge = `<span class="pacing-badge badge-easy">Easy</span>`;
      analysis = `Party survives well past the target encounter length. Consider increasing boss DPR or HP. ` +
        `Target boss DPR multiplier: ${fmtMult(r.targetBossDprMult)}.`;
    }

    const advice = tightPacingAdviceHtml();
    el.className = `pacing-balance ${cls}`;
    el.innerHTML = `${badge} ${analysis}${advice}`;
  }

  function tightPacingAdviceHtml() {
    if (!state.tight_pacing_enabled) return "";
    const capPct = Math.round(clamp(safeFloat(state.tight_cap_pct, 0.24), 0.10, 0.40) * 100);
    const resources = Math.max(0, safeInt(state.tight_cap_resources, 3));
    const target = clamp(safeInt(state.pacing_rounds, 5), 5, 10);
    const antiSlog = clamp(safeFloat(state.tight_anti_slog_mult, 1.35), 1.0, 2.5);
    return `<div class="pacing-advice"><strong>Tight pacing table rule:</strong> Until round ${target}, the boss can burn ${resources} Legendary Resistance/Action pacing resource(s) to prevent HP loss beyond ${capPct}% max HP in a round. After round ${target}, remove that protection and make the boss collapse: damage that reaches boss HP is modeled at ${antiSlog.toFixed(2)}x per late round.</div>`;
  }

  function tightPacingAdviceText() {
    if (!state.tight_pacing_enabled) return "";
    const capPct = Math.round(clamp(safeFloat(state.tight_cap_pct, 0.24), 0.10, 0.40) * 100);
    const resources = Math.max(0, safeInt(state.tight_cap_resources, 3));
    const target = clamp(safeInt(state.pacing_rounds, 5), 5, 10);
    const antiSlog = clamp(safeFloat(state.tight_anti_slog_mult, 1.35), 1.0, 2.5);
    return `\n\nTable rule required for these numbers:\n- Until round ${target}, the boss may burn ${resources} Legendary Resistance/Action pacing resource(s) to prevent HP loss beyond ${capPct}% max HP in a round.\n- After round ${target}, remove that protection and make the boss collapse: damage that reaches boss HP is modeled at ${antiSlog.toFixed(2)}x per late round.\n- If you do not run those mechanics, expect a wider p10-p90 spread than the tool reports.`;
  }

  function fmt1(n) {
    return n == null ? "∞" : Number.isFinite(n) ? n.toFixed(1) : "∞";
  }

  function fmtMult(n) {
    return n == null || !Number.isFinite(n) ? "N/A" : `${n.toFixed(2)}x`;
  }

  function setText(el, text) {
    if (el) el.textContent = text;
  }

  function applyMetricClass(el, cls) {
    if (!el) return;
    el.classList.remove("danger", "warning", "success");
    if (cls) el.classList.add(cls);
  }

  // ── End Pacing Calculator ─────────────────────────────────────────────────

  // ── Minion helpers ────────────────────────────────────────────────────────

  // Scale factor for party effective DPR when attacking minion AC vs boss AC.
  // Uses the average attack bonus from the nova table as a proxy.
  function minionHitRatio(minionAc, bossAc) {
    const rows = state.party_nova_table;
    const avgAtk = rows.length
      ? rows.reduce((acc, r) => acc + safeInt(r["Atk Bonus"], 7), 0) / rows.length
      : 7;
    const pHitMinion = clamp((21 + avgAtk - Math.max(1, minionAc)) / 20, 0.05, 0.95);
    const pHitBoss   = clamp((21 + avgAtk - Math.max(1, bossAc))   / 20, 0.05, 0.95);
    return pHitBoss > 0.001 ? pHitMinion / pHitBoss : 1;
  }

  // Expected DPR of ONE minion against ONE specific PC (no partySize scaling).
  function minionOneDprVsPc(pcRow) {
    let dpr = 0;
    if (state.minion_atk_enabled) {
      const ac       = Math.max(1, safeInt(pcRow.AC, 10));
      const atkBonus = safeInt(state.minion_atk_bonus, 4);
      const avgDmg   = Math.max(0, safeFloat(state.minion_atk_avg, 5));
      const pHit     = clamp((21 + atkBonus - ac) / 20, 0.05, 0.95);
      dpr += pHit * avgDmg;
    }
    if (state.minion_save_enabled) {
      const saveBonus = getSaveBonus(pcRow, state.minion_save_stat);
      const dc        = Math.max(1, safeInt(state.minion_save_dc, 13));
      const avgDmg    = Math.max(0, safeFloat(state.minion_save_avg, 7));
      const pFail     = pSaveFail(dc, saveBonus);
      dpr += (0.5 + 0.5 * pFail) * avgDmg;
    }
    return dpr;
  }

  // Full minion pack DPR against one specific PC, distributed across the full party size.
  function minionDprVsPc(pcRow) {
    const count = Math.max(0, safeInt(state.minion_count, 0));
    if (count <= 0) return 0;
    const partySize = state.party_table.filter((r) => String(r.Name || "").trim().length > 0).length || 1;
    return (count / partySize) * minionOneDprVsPc(pcRow);
  }

  // Average DPR per minion, averaged over party composition (used for table display).
  function minionDprEachAvg() {
    const count = Math.max(0, safeInt(state.minion_count, 0));
    if (count <= 0) return 0;
    const members = state.party_table.filter((r) => String(r.Name || "").trim().length > 0);
    if (!members.length) return 0;
    return members.reduce((acc, pc) => acc + minionOneDprVsPc(pc), 0) / members.length;
  }

  function sampleMinionDamageVsPc(pcIndex, pcAc, pcSaves, saveIndex, opts) {
    let raw = 0;

    if (opts.minion_atk_enabled) {
      const ac = Math.max(1, safeInt(pcAc[pcIndex], 10));
      const atkBonus = safeInt(opts.minion_atk_bonus, 4);
      const roll = randomInt(1, 20);
      const hit = roll === 20 || (roll !== 1 && roll + atkBonus >= ac);
      if (hit) {
        raw += Math.max(0, safeFloat(opts.minion_atk_avg, 5));
      }
    }

    if (opts.minion_save_enabled) {
      const stat = String(opts.minion_save_stat || "DEX").toUpperCase();
      const saveIdx = saveIndex[stat] ?? saveIndex.DEX;
      const bonus = pcSaves[pcIndex][saveIdx];
      const dc = Math.max(1, safeInt(opts.minion_save_dc, 13));
      const roll = randomInt(1, 20);
      const success = roll === 20 || (roll !== 1 && roll + bonus >= dc);
      const avg = Math.max(0, safeFloat(opts.minion_save_avg, 7));
      raw += success ? 0.5 * avg : avg;
    }

    return raw;
  }

  // Round-by-round minion-phase simulation (deterministic, focus-fire model).
  // Returns null when no minions are configured.
  function computeMinionPhase(rawPartyDpr, useNova = Boolean(state.enc_use_nova)) {
    const count    = Math.max(0, safeInt(state.minion_count, 0));
    const hpEach   = Math.max(1, safeFloat(state.minion_hp, 15));
    const mAc      = Math.max(1, safeInt(state.minion_ac, 14));
    const dprEach  = minionDprEachAvg();
    if (count <= 0) return null;

    const partyDprVsMinions = expectedPartyDprVsEnemy(useNova, mAc);
    const ratio             = rawPartyDpr > 1e-9
      ? partyDprVsMinions / rawPartyDpr
      : Number.POSITIVE_INFINITY;
    const resist            = Math.max(1e-6, safeFloat(state.resist_factor, 1.0));
    const regen             = Math.max(0, safeFloat(state.boss_regen, 0.0));
    const replenish         = Boolean(state.minion_replenish);
    const totalPool         = count * hpEach;

    const rows = [];
    let bossHp = Math.max(1, safeFloat(state.boss_hp, 150));

    // ── Replenishing: full minion count respawns at the start of every round ─
    if (replenish) {
      const minionTeamDpr = count * dprEach;
      const canClear      = partyDprVsMinions >= totalPool;

      if (!canClear) {
        // Party can't one-round the pack — boss is permanently screened.
        for (let r = 1; r <= 5; r++) {
          rows.push({
            Round: r,
            "Minions (replenish)": count,
            "DPR vs Minions":      round2(partyDprVsMinions),
            "Overflow to Boss":    0,
            "Boss HP Left":        round2(bossHp),
            "Minion Team DPR":     round2(minionTeamDpr),
          });
        }
        return { rows, minionPhaseRounds: Infinity, bossHpAfterMinionPhase: bossHp,
                 partyDprVsMinions, ratio, replenishing: true, canClearPerRound: false };
      }

      // Party clears the full pack each round; excess spills to boss.
      const overflowBoss = Math.max(0, ((partyDprVsMinions - totalPool) / ratio) / resist - regen);
      let round = 0;
      while (bossHp > 0 && round < 200) {
        round++;
        bossHp = Math.max(0, bossHp - overflowBoss);
        rows.push({
          Round: round,
          "Minions (replenish)": "0 (cleared)",
          "DPR vs Minions":      round2(totalPool),
          "Overflow to Boss":    round2(overflowBoss),
          "Boss HP Left":        round2(bossHp),
          "Minion Team DPR":     round2(minionTeamDpr),
        });
      }
      return { rows, minionPhaseRounds: round, bossHpAfterMinionPhase: 0,
               overflowBossPerRound: overflowBoss, partyDprVsMinions, ratio,
               replenishing: true, canClearPerRound: true };
    }

    // ── Non-replenishing: pool depletes over multiple rounds ──────────────
    let minionPoolHp = totalPool;
    let round = 0;
    const initialBossHp = bossHp;

    while (minionPoolHp > 0 && round < 60) {
      round++;
      const minionsAtStart = Math.ceil(minionPoolHp / hpEach);
      const minionTeamDpr  = minionsAtStart * dprEach;
      let dprToBoss = 0;

      if (partyDprVsMinions >= minionPoolHp) {
        const overflow = partyDprVsMinions - minionPoolHp;
        minionPoolHp   = 0;
        dprToBoss      = ratio > 0 ? Math.max(0, (overflow / ratio) / resist - regen) : 0;
      } else {
        minionPoolHp -= partyDprVsMinions;
      }

      bossHp -= dprToBoss;

      rows.push({
        Round: round,
        "Minions Alive":    minionPoolHp > 0 ? Math.ceil(minionPoolHp / hpEach) : 0,
        "DPR vs Minions":   round2(Math.min(partyDprVsMinions, minionsAtStart * hpEach)),
        "Overflow to Boss": round2(dprToBoss),
        "Boss HP Left":     round2(Math.max(0, bossHp)),
        "Minion Team DPR":  round2(minionTeamDpr),
      });
    }

    const totalBossDmgDuringPhase = Math.max(0, initialBossHp - Math.max(0, bossHp));
    return { rows, minionPhaseRounds: round, bossHpAfterMinionPhase: Math.max(0, bossHp),
             totalBossDmgDuringPhase, partyDprVsMinions, ratio, replenishing: false };
  }

  // Render the minion round-by-round card in the TTD panel.
  function renderMinionTtdTable(phase) {
    if (!els.minionTtdCard || !els.minionTtdTable) return;
    if (!phase) {
      els.minionTtdCard.style.display = "none";
      return;
    }
    els.minionTtdCard.style.display = "";
    renderResultTable(els.minionTtdTable, phase.rows);
  }

  // ── End Minion helpers ────────────────────────────────────────────────────

  function dprRowFor(memberName) {
    return state.party_dpr_table.find(r => r.Member === memberName) || { Damage: "1d6" };
  }

  function buildNovaEffRows() {
    let total = 0;
    const rows = [];

    for (const row of state.party_nova_table) {
      const dprRow = dprRowFor(row.Member);
      const conv = computeNovaConversion(row, dprRow, state.boss_ac);
      total += conv.effDpr;

      rows.push({
        Member: row.Member || "?",
        "Dmg/Atk": round2(conv.avgDmg),
        Attacks: conv.attacks,
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
      return state.party_dpr_table.map((row) => [row.Member || "?", averageDamage(row.Damage || "1d6")]);
    }

    const out = [];
    for (const row of state.party_nova_table) {
      const conv = computeNovaConversion(row, dprRowFor(row.Member), state.boss_ac);
      out.push([row.Member || "?", conv.effDpr]);
    }
    return out;
  }

  function expectedPartyDprVsEnemy(useNova, targetAc, targetSaveBonus = null) {
    if (!useNova) {
      const bossAc = Math.max(1, safeInt(state.boss_ac, 16));
      const ratio = minionHitRatio(Math.max(1, safeInt(targetAc, bossAc)), bossAc);
      return state.party_dpr_table.reduce((acc, row) => {
        return acc + averageDamage(row.Damage || "1d6") * ratio;
      }, 0);
    }

    let total = 0;
    for (const row of state.party_nova_table) {
      const saveBonus = targetSaveBonus == null
        ? safeInt(row["Target Save Bonus"], 0)
        : safeInt(targetSaveBonus, 0);
      total += computeNovaConversion(row, dprRowFor(row.Member), targetAc, saveBonus).effDpr;
    }
    return total;
  }

  function computeNovaConversion(novaRow, dprRow, fallbackBossAc, overrideSaveBonus = null) {
    const row    = novaRow;
    const method = normalizeNovaMethod(row.Method);
    const mode   = normalizeRollMode(row["Roll Mode"]);
    const avgDmg = Math.max(0, averageDamage((dprRow && dprRow.Damage) || "1d6"));
    const attacks = Math.max(1, safeInt(row["Attacks"], 1));
    const uptime = clamp(safeFloat(row.Uptime, 0.85), 0, 1);

    let pMain = 1;
    let pCrit = 0;
    let baseFactor = 1;

    if (method === "attack") {
      const atkBonus      = safeInt(row["Atk Bonus"], 0);
      const targetAc      = Math.max(1, safeInt(fallbackBossAc, safeInt(row["Target AC"], 16)));
      const critThreshold = Math.max(2, Math.min(20, safeInt(row["Crit"], 20)));
      const [pNon, crit, pAny] = hitProbs(targetAc, atkBonus, mode, critThreshold);
      pMain = pAny;
      pCrit = crit;
      const dmgExpr = (dprRow && dprRow.Damage) || "1d6";
      const critRatio = avgDmg > 0 ? averageCrunchyCritDamage(dmgExpr) / avgDmg : 1;
      baseFactor = pNon + critRatio * pCrit;
    } else if (method === "save_half") {
      const saveDc = Math.max(1, safeInt(row["Save DC"], 16));
      const targetSaveBonus = overrideSaveBonus == null
        ? safeInt(row["Target Save Bonus"], 0)
        : safeInt(overrideSaveBonus, 0);
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
      avgDmg,
      attacks,
      effDpr: avgDmg * attacks * factor,
    };
  }

  function sampleNovaActionSplit(memberName, minionPoolHp, bossAc, minionAc) {
    const row = state.party_nova_table.find((r) => String(r.Member || "") === String(memberName || ""));
    if (!row) {
      return { bossRaw: 0, minionPoolHp };
    }

    const dprRow = dprRowFor(row.Member);
    const attacks = Math.max(1, safeInt(row["Attacks"], 1));
    const uptime = clamp(safeFloat(row.Uptime, 0.85), 0, 1);
    let remainingMinionHp = Math.max(0, safeFloat(minionPoolHp, 0));
    let bossRaw = 0;

    if (Math.random() > uptime) {
      return { bossRaw, minionPoolHp: remainingMinionHp };
    }

    for (let i = 0; i < attacks; i += 1) {
      const targetMinion = remainingMinionHp > 0;
      const targetAc = targetMinion ? minionAc : bossAc;
      const dmg = sampleNovaAttackDamage(row, dprRow, targetAc);

      if (targetMinion) {
        remainingMinionHp = Math.max(0, remainingMinionHp - dmg);
      } else {
        bossRaw += dmg;
      }
    }

    return { bossRaw, minionPoolHp: remainingMinionHp };
  }

  function sampleNovaAttackDamage(novaRow, dprRow, targetAc) {
    const method = normalizeNovaMethod(novaRow.Method);
    const expr = (dprRow && dprRow.Damage) || "1d6";

    if (method === "auto_hit") {
      return rollDamageOne(expr);
    }

    if (method === "save_half") {
      const dc = Math.max(1, safeInt(novaRow["Save DC"], 16));
      const saveBonus = safeInt(novaRow["Target Save Bonus"], 0);
      const saveRoll = rollD20Mode(normalizeRollMode(novaRow["Roll Mode"]));
      const success = saveRoll === 20 || (saveRoll !== 1 && saveRoll + saveBonus >= dc);
      const dmg = rollDamageOne(expr);
      const saveMult = clamp(safeFloat(novaRow["Save Success Mult"], 0.5), 0, 1);
      return success ? dmg * saveMult : dmg;
    }

    const roll = rollD20Mode(normalizeRollMode(novaRow["Roll Mode"]));
    const critThreshold = Math.max(2, Math.min(20, safeInt(novaRow["Crit"], 20)));
    const atkBonus = safeInt(novaRow["Atk Bonus"], 0);
    const isCrit = roll >= critThreshold;
    const isHit = isCrit || (roll !== 1 && roll + atkBonus >= Math.max(1, safeInt(targetAc, 16)));
    if (!isHit) return 0;

    return isCrit ? rollDamageCrunchyCritOne(expr) : rollDamageOne(expr);
  }

  // Derives the per-PC DPR CV for the manual-DPR fallback model.
  // Nova-mode encounter MC rolls each captured attack directly; manual mode
  // still draws gammaRng(effMean_j, dprCv) for each PC independently.
  // When N such draws are summed, the party total has:
  //   CV_party_sim = dprCv × sqrt(Σ effMean²) / Σ effMean
  //
  // This function computes CV_physics (the party variance from actual attack physics)
  // then applies the correction factor so that CV_party_sim = CV_physics:
  //   dprCv_returned = CV_physics × Σ effMean / sqrt(Σ effMean²)
  //
  // Includes hit/miss variance and damage-dice variance per attack.
  function computePartyDprCv() {
    const rows = state.party_nova_table.filter(r => String(r.Member || "").trim());
    if (!rows.length) return 0.6;
    let totalDpr = 0;
    let totalVar = 0;
    let sumDprSq = 0;
    for (const row of rows) {
      const dprRow  = dprRowFor(row.Member);
      const avgDmg  = Math.max(0, averageDamage(dprRow.Damage || "1d6"));
      const varDmg  = varianceDamage(dprRow.Damage || "1d6");
      const nova    = computeNovaConversion(row, dprRow, safeInt(state.boss_ac, 16));
      const attacks = nova.attacks;
      const uptime  = clamp(safeFloat(row.Uptime, 0.85), 0, 1);
      const eff     = nova.effDpr;
      if (eff <= 0) continue;

      let varPerAtk;
      if (nova.method === "attack") {
        const p = clamp(nova.pMain, 0.05, 0.95);
        // Var[single attack] = p(1-p)·avgDmg² + p·varDmg
        varPerAtk = p * (1 - p) * avgDmg * avgDmg + p * varDmg;
      } else {
        const pFail = clamp(nova.pMain, 0.05, 0.95);
        const varMult = 0.25 * pFail * (1 - pFail);
        const eMult2  = 0.25 + 0.75 * pFail;
        varPerAtk = avgDmg * avgDmg * varMult + eMult2 * varDmg;
      }
      totalDpr += eff;
      totalVar += attacks * uptime * varPerAtk;
      sumDprSq += eff * eff;
    }
    if (totalDpr <= 0) return 0.6;
    // cvPhysics is the true party-level CV from attack structure.
    const cvPhysics = Math.sqrt(totalVar) / totalDpr;
    // Correction: simulation draws per-PC with the same CV, so simulated party CV =
    // dprCv × sqrt(Σm²)/Σm.  To match cvPhysics, scale by Σm/sqrt(Σm²) ≥ 1.
    const corrFactor = sumDprSq > 0 ? totalDpr / Math.sqrt(sumDprSq) : 1.0;
    return clamp(cvPhysics * corrFactor, 0.05, 2.0);
  }

  function refreshDerivedCv() {
    if (!Boolean(state.enc_use_nova)) return;
    const derived = round2(computePartyDprCv());
    if (Math.abs(derived - state.dpr_cv) > 0.005) {
      state.dpr_cv = derived;
      if (els.encDprCv) setControlValue(els.encDprCv, state.dpr_cv);
    }
  }

  function perRoundDprVsPc(pcRow, mode, attacks, horizonRounds = 1) {
    const ac = Math.max(1, safeInt(pcRow.AC, 10));
    let total = 0;

    for (const atk of attacks) {
      const uses = effectiveUsesPerRound(atk, horizonRounds);
      if (uses <= 0) continue;
      let dpr = 0;
      if (atk.kind === "save") {
        dpr = expectedSaveHalfDamage(atk.dc, getSaveBonus(pcRow, atk.save_stat), atk.damage_expr);
      } else {
        dpr = expectedAttackDamage(ac, atk.attack_bonus, atk.damage_expr, mode, atk.crit_threshold);
      }
      total += dpr * uses;
    }

    return total;
  }

  function effectiveUsesPerRound(atk, horizonRounds) {
    const perRound = Math.max(0, safeInt(atk && atk.uses_per_round, 0));
    if (perRound <= 0) return 0;
    const encounterUses = Math.max(0, safeInt(atk && atk.uses_encounter, 0));
    if (encounterUses <= 0) return perRound;
    return Math.min(perRound, encounterUses / Math.max(1, safeFloat(horizonRounds, 1)));
  }

  function attackUseCapacity(atk) {
    const encounterUses = Math.max(0, safeInt(atk && atk.uses_encounter, 0));
    return encounterUses > 0 ? encounterUses : Number.POSITIVE_INFINITY;
  }

  function phaseMechanicsPerTargetDpr(pcRow, mode, mechanics, partySize, horizonRounds) {
    const rounds = Math.max(1, safeFloat(horizonRounds, 1));
    const pSize = Math.max(1, safeInt(partySize, 1));
    let total = 0;

    for (const mech of mechanics || []) {
      if (safeInt(mech.round, 1) > rounds) continue;
      const targetShare = Math.min(1, Math.max(1, safeInt(mech.targets, 1)) / pSize);
      total += expectedMechanicDamageVsPc(mech, pcRow, mode) * targetShare / rounds;
    }

    return total;
  }

  function expectedMechanicDamageVsPc(mech, pcRow, mode) {
    const kind = normalizeMechanicType(mech && mech.kind);
    if (kind === "auto") {
      return averageDamage(mech.damage_expr);
    }
    if (kind === "save") {
      return expectedSaveHalfDamage(mech.dc, getSaveBonus(pcRow, mech.save_stat), mech.damage_expr);
    }
    return expectedAttackDamage(Math.max(1, safeInt(pcRow.AC, 10)), mech.attack_bonus, mech.damage_expr, mode, mech.crit_threshold);
  }

  function mechanicsForRound(mechanics, round) {
    const r = Math.max(1, safeInt(round, 1));
    return (mechanics || []).filter((mech) => safeInt(mech.round, 1) === r);
  }

  function rollMechanicDamageVsPc(mech, pcRow, mode, ac, saveBonuses) {
    const kind = normalizeMechanicType(mech && mech.kind);
    if (kind === "auto") {
      return rollDamageOne(mech.damage_expr);
    }
    if (kind === "save") {
      const bonus = saveBonuses[String(mech.save_stat || "DEX").toUpperCase()] || 0;
      const roll = randomInt(1, 20);
      const success = roll === 20 || (roll !== 1 && roll + bonus >= mech.dc);
      const dmg = rollDamageOne(mech.damage_expr);
      return success ? 0.5 * dmg : dmg;
    }

    const r1 = randomInt(1, 20);
    const r2 = randomInt(1, 20);
    let roll = r1;
    if (mode === "adv") roll = Math.max(r1, r2);
    if (mode === "dis") roll = Math.min(r1, r2);
    const isCrit = roll >= (mech.crit_threshold ?? 20);
    const isHit = isCrit || (roll !== 1 && roll + mech.attack_bonus >= ac);
    if (!isHit) return 0;
    return isCrit ? rollDamageCrunchyCritOne(mech.damage_expr) : rollDamageOne(mech.damage_expr);
  }

  function rollMechanicDamageForTarget(mech, target, currentAc, currentMode, pcSaves, saveIndex) {
    const kind = normalizeMechanicType(mech && mech.kind);
    if (kind === "auto") {
      return rollDamageOne(mech.damage_expr);
    }
    if (kind === "save") {
      const saveIdx = saveIndex[String(mech.save_stat || "DEX").toUpperCase()] ?? saveIndex.DEX;
      const bonus = pcSaves[target][saveIdx];
      const roll = randomInt(1, 20);
      const success = roll === 20 || (roll !== 1 && roll + bonus >= mech.dc);
      const dmg = rollDamageOne(mech.damage_expr);
      return success ? 0.5 * dmg : dmg;
    }

    const r1 = randomInt(1, 20);
    const r2 = randomInt(1, 20);
    let r = r1;
    if (currentMode[target] === "adv") r = Math.max(r1, r2);
    if (currentMode[target] === "dis") r = Math.min(r1, r2);
    const isCrit = r >= (mech.crit_threshold ?? 20);
    const isHit = isCrit || (r !== 1 && r + mech.attack_bonus >= currentAc[target]);
    if (!isHit) return 0;
    return isCrit ? rollDamageCrunchyCritOne(mech.damage_expr) : rollDamageOne(mech.damage_expr);
  }

  function bossDprMultiplier(opts) {
    return clamp(safeFloat(opts && opts.boss_dpr_mult, 1.0), 0, 20);
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

  function phaseMechanicsEnabledFromTable(tbl) {
    const out = [];
    for (const row of tbl || []) {
      if (!Boolean(row["Enabled?"])) continue;

      const saveStatRaw = String(row.Save || "DEX").toUpperCase();
      const saveStat = SAVE_KEYS.includes(saveStatRaw) ? saveStatRaw : "DEX";

      out.push({
        name: String(row.Name || "Mechanic"),
        round: Math.max(1, safeInt(row.Round, 1)),
        kind: normalizeMechanicType(row.Type),
        attack_bonus: safeInt(row["Attack bonus"], 0),
        crit_threshold: Math.max(2, Math.min(20, safeInt(row["Crit"], 20))),
        dc: safeInt(row.DC, 0),
        save_stat: saveStat,
        damage_expr: String(row.Damage || "1d6"),
        targets: Math.max(1, safeInt(row.Targets, 1)),
      });
    }
    return out;
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
        uses_encounter: Math.max(0, safeInt(row["Uses/encounter"], 0)),
        is_melee: Boolean(row["Melee?"]),
        crit_threshold: Math.max(2, Math.min(20, safeInt(row["Crit"], 20))),
      });
    }
    return out;
  }

  function getSaveBonus(pcRow, stat) {
    const key = String(stat || "DEX").toUpperCase();
    return safeInt(pcRow[key], 0);
  }

  function expectedAttackDamage(ac, attackBonus, dmgExpr, mode = "normal", critThreshold = 20) {
    const [pNonCrit, pCrit] = hitProbs(ac, attackBonus, mode, critThreshold);
    return pNonCrit * averageDamage(dmgExpr) + pCrit * averageCrunchyCritDamage(dmgExpr);
  }

  function expectedSaveHalfDamage(dc, saveBonus, dmgExpr) {
    return (0.5 + 0.5 * pSaveFail(dc, saveBonus)) * averageDamage(dmgExpr);
  }

  function hitProbs(ac, attackBonus, mode = "normal", critThreshold = 20) {
    const safeCrit = Math.max(2, Math.min(20, critThreshold));
    const classify = (roll) => {
      if (roll === 1) return [false, false];
      if (roll >= safeCrit) return [true, true];
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

  // Inverse standard-normal CDF (rational approximation, A&S 26.2.17).
  // Valid for p in (0, 1); sufficient precision for p in [0.05, 0.95].
  function invNorm01(p) {
    const safe = Math.max(1e-9, Math.min(1 - 1e-9, p));
    const q    = safe <= 0.5 ? safe : 1 - safe;
    const t    = Math.sqrt(-2 * Math.log(q));
    const z    = t - (2.515517 + 0.802853 * t + 0.010328 * t * t)
                   / (1 + 1.432788 * t + 0.189269 * t * t + 0.001308 * t * t * t);
    return safe <= 0.5 ? -z : z;
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

  function normalizeMechanicType(value) {
    const kind = String(value || "auto").toLowerCase();
    return MECHANIC_TYPES.includes(kind) ? kind : "auto";
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

  // Variance of a damage expression's dice terms. Var(XdN) = X·(N²-1)/12.
  // Uses Math.abs(count) because variance is always additive (Var[X-Y] = Var[X]+Var[Y]).
  function varianceDamage(expr) {
    const parsed = parseDamageExpression(expr);
    let variance = 0;
    for (const [count, sides] of parsed.dice) {
      if (sides <= 1) continue;
      variance += Math.abs(count) * (sides * sides - 1) / 12;
    }
    return Math.max(0, variance);
  }

  function averageCrunchyCritDamage(expr) {
    const parsed = parseDamageExpression(expr);
    let avgDice = 0;
    for (const [count, sides] of parsed.dice) {
      // Positive dice: crunchy crit = max value + average roll per die.
      // Negative dice (e.g. "2d6-1d4"): no crit bonus — treat as normal average.
      if (count > 0) {
        avgDice += count * (sides + (sides + 1) / 2);
      } else {
        avgDice += count * (sides + 1) / 2;
      }
    }
    return Math.max(0, parsed.sign * (avgDice + parsed.mod));
  }

  function scaleDamageExpression(expr, multiplier) {
    const avg = averageDamage(expr);
    const targetAvg = Math.max(0, avg * Math.max(0, safeFloat(multiplier, 1)));
    const preferredSides = dominantDieSides(expr) || 6;
    return damageFormulaForAverage(targetAvg, preferredSides);
  }

  function dominantDieSides(expr) {
    const parsed = parseDamageExpression(expr);
    if (!parsed.dice.length) return 6;

    const bySides = new Map();
    for (const [count, sides] of parsed.dice) {
      if (sides <= 0) continue;
      bySides.set(sides, (bySides.get(sides) || 0) + Math.abs(count));
    }

    let bestSides = 6;
    let bestCount = 0;
    for (const [sides, count] of bySides.entries()) {
      if (count > bestCount) {
        bestSides = sides;
        bestCount = count;
      }
    }
    return bestSides;
  }

  function damageFormulaForAverage(targetAvg, preferredSides = 6) {
    const avg = Math.max(0, safeFloat(targetAvg, 0));
    if (avg <= 0.05) return "0";

    const sideCandidates = uniqueNumbers([preferredSides, 6, 8, 10, 12, 4, 20, 5, 3, 2, 1])
      .filter((n) => n >= 1 && n <= 100);
    let best = null;

    for (const sides of sideCandidates) {
      const dieAvg = (sides + 1) / 2;
      const maxCount = Math.max(1, Math.min(40, Math.ceil(avg / dieAvg) + 4));
      for (let count = 1; count <= maxCount; count += 1) {
        const baseAvg = count * dieAvg;
        const mod = Math.max(0, Math.round(avg - baseAvg));
        const candidateAvg = baseAvg + mod;
        const err = Math.abs(candidateAvg - avg);
        const sidePenalty = sides === preferredSides ? 0 : 0.02;
        const complexity = count * 0.01 + (mod > 0 ? 0.005 : 0) + sidePenalty;
        const score = err + complexity;
        if (!best || score < best.score) {
          best = { count, sides, mod, score };
        }
      }
    }

    if (!best) return String(Math.max(0, Math.round(avg)));
    const dice = `${best.count}d${best.sides}`;
    return best.mod > 0 ? `${dice}+${best.mod}` : dice;
  }

  function uniqueNumbers(values) {
    const out = [];
    const seen = new Set();
    for (const value of values) {
      const n = Math.max(0, safeInt(value, 0));
      if (seen.has(n)) continue;
      seen.add(n);
      out.push(n);
    }
    return out;
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
        legend: { labels: { color: "#f0ebe0" } },
        title: { display: false, color: "#ecf3f9" },
      },
      scales: {
        x: { ticks: { color: "#aa9080" }, grid: { color: "rgba(191,26,47,0.12)" } },
        y: { ticks: { color: "#aa9080" }, grid: { color: "rgba(255,255,255,0.08)" } },
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
    const payload = {
      ...state,
      _tight_pacing_advice: tightPacingAdviceText().trim(),
    };
    els.reportText.value = JSON.stringify(payload, null, 2);
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
      if (typeof localStorage === "undefined") return null;
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (_) {
      return null;
    }
  }

  function persistState() {
    try {
      if (typeof localStorage === "undefined") return;
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
    base.phase_table = (base.phase_table || []).map(sanitizePhaseRow);
    base.party_dpr_table = (base.party_dpr_table || []).map(sanitizeDprRow);
    base.party_nova_table = (base.party_nova_table || []).map(sanitizeNovaRow);

    base.mode_select = ["normal", "adv", "dis"].includes(String(base.mode_select)) ? String(base.mode_select) : "normal";
    base.spread_targets = Math.max(1, safeInt(base.spread_targets, 1));
    base.thp_expr = String(base.thp_expr || "0");

    base.lair_enabled = Boolean(base.lair_enabled);
    base.lair_avg = Math.max(0, safeFloat(base.lair_avg, 6.0));
    base.lair_formula = String(base.lair_formula || damageFormulaForAverage(base.lair_avg, 6)).trim();
    base.lair_targets = Math.max(1, safeInt(base.lair_targets, 2));
    base.lair_every_n = Math.max(1, safeInt(base.lair_every_n, 2));

    base.rech_enabled = Boolean(base.rech_enabled);
    base.recharge_text = String(base.recharge_text || "5-6");
    base.rech_avg = Math.max(0, safeFloat(base.rech_avg, 22.0));
    base.rech_formula = String(base.rech_formula || damageFormulaForAverage(base.rech_avg, 6)).trim();
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
    base.boss_dpr_mult = clamp(safeFloat(base.boss_dpr_mult, 1.0), 0, 20);

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

    base.tune_target_median = clamp(safeFloat(base.tune_target_median, 5.0), 1.0, 20.0);
    base.tune_tpk_cap  = clamp(safeFloat(base.tune_tpk_cap, 0.05), 0.0, 1.0);
    base.tune_kill_rate = clamp(safeFloat(base.tune_kill_rate, 0.75), 0.50, 0.95);
    base.tune_band_max  = Math.max(0.5, safeFloat(base.tune_band_max, 3.0));

    base.pacing_rounds = clamp(safeInt(base.pacing_rounds, 5), 5, 10);
    base.tight_pacing_enabled = Object.prototype.hasOwnProperty.call(src, "tight_pacing_enabled")
      ? Boolean(base.tight_pacing_enabled)
      : true;
    base.tight_cap_pct = clamp(safeFloat(base.tight_cap_pct, 0.24), 0.10, 0.40);
    base.tight_cap_resources = clamp(safeInt(base.tight_cap_resources, 3), 0, 6);
    base.tight_anti_slog_mult = clamp(safeFloat(base.tight_anti_slog_mult, 1.35), 1.0, 2.5);

    base.minion_count     = Math.max(0, safeInt(base.minion_count, 0));
    base.minion_ac        = Math.max(1, safeInt(base.minion_ac, 14));
    base.minion_hp        = Math.max(1, safeFloat(base.minion_hp, 15));
    base.minion_replenish = Boolean(base.minion_replenish);

    base.minion_atk_enabled = Boolean(base.minion_atk_enabled);
    base.minion_atk_bonus   = clamp(safeInt(base.minion_atk_bonus, 4), -10, 20);
    base.minion_atk_avg     = Math.max(0, safeFloat(base.minion_atk_avg, 5));

    base.minion_save_enabled = Boolean(base.minion_save_enabled);
    base.minion_save_stat    = SAVE_KEYS.includes(String(base.minion_save_stat || "DEX").toUpperCase())
      ? String(base.minion_save_stat).toUpperCase() : "DEX";
    base.minion_save_dc      = clamp(safeInt(base.minion_save_dc, 13), 1, 30);
    base.minion_save_avg     = Math.max(0, safeFloat(base.minion_save_avg, 7));

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
      let dmg = String(row.Damage || "").trim();
      if (!dmg && typeof row.DPR === "number" && row.DPR > 0) {
        dmg = damageFormulaForAverage(row.DPR, 6);
      }
      return {
        Member: name,
        Damage: dmg || "1d6",
      };
    });

    const bossAc = Math.max(1, safeInt(st.boss_ac, 16));
    st.party_nova_table = names.map((name) => {
      const nRow = novaMap.get(name) || {};
      return {
        Member: name,
        Method: normalizeNovaMethod(nRow.Method),
        "Attacks": Math.max(1, safeInt(nRow["Attacks"], 1)),
        "Atk Bonus": safeInt(nRow["Atk Bonus"], 7),
        "Roll Mode": normalizeRollMode(nRow["Roll Mode"]),
        "Target AC": bossAc,
        "Crit": Math.max(2, Math.min(20, safeInt(nRow["Crit"], 20))),
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
    // Migrate old saved state: if Damage absent but DPR number present, convert to formula.
    let dmg = String(row.Damage || "").trim();
    if (!dmg && typeof row.DPR === "number" && row.DPR > 0) {
      dmg = damageFormulaForAverage(row.DPR, 6);
    }
    return {
      Member: String(row.Member || "").trim(),
      Damage: dmg || "1d6",
    };
  }

  function sanitizeNovaRow(row) {
    return {
      Member: String(row.Member || "").trim(),
      Method: normalizeNovaMethod(row.Method),
      "Attacks": Math.max(1, safeInt(row["Attacks"], 1)),
      "Atk Bonus": safeInt(row["Atk Bonus"], 7),
      "Roll Mode": normalizeRollMode(row["Roll Mode"]),
      "Target AC": Math.max(1, safeInt(row["Target AC"], 16)),
      "Crit": Math.max(2, Math.min(20, safeInt(row["Crit"], 20))),
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
      "Uses/encounter": Math.max(0, safeInt(row["Uses/encounter"], 0)),
      "Melee?": Boolean(row["Melee?"]),
      "Enabled?": Boolean(row["Enabled?"]),
      "Crit": Math.max(2, Math.min(20, safeInt(row["Crit"], 20))),
    };
  }

  function sanitizePhaseRow(row) {
    return {
      Round: Math.max(1, safeInt(row.Round, 1)),
      Name: String(row.Name || "Mechanic").trim() || "Mechanic",
      Type: normalizeMechanicType(row.Type),
      "Attack bonus": safeInt(row["Attack bonus"], 0),
      Crit: Math.max(2, Math.min(20, safeInt(row.Crit, 20))),
      DC: safeInt(row.DC, 0),
      Save: SAVE_KEYS.includes(String(row.Save || "DEX").toUpperCase()) ? String(row.Save).toUpperCase() : "DEX",
      Damage: String(row.Damage || "1d6").trim() || "1d6",
      Targets: Math.max(1, safeInt(row.Targets, 1)),
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

  function rollD20Mode(mode = "normal") {
    const r1 = randomInt(1, 20);
    if (mode === "normal") return r1;
    const r2 = randomInt(1, 20);
    return mode === "adv" ? Math.max(r1, r2) : Math.min(r1, r2);
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

  function setBtnLoading(el, isLoading) {
    if (!el) return;
    el.disabled = isLoading;
    el.classList.toggle("loading", isLoading);
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

  if (IS_WORKER) {
    self.onmessage = (event) => {
      const data = event.data || {};
      if (data.type !== "runEncounterMc") return;
      try {
        state = normalizeState(data.state || {});
        const result = runEncounterMc(data.overrides || {});
        self.postMessage({ id: data.id, result });
      } catch (error) {
        self.postMessage({ id: data.id, error: error && error.message ? error.message : String(error) });
      }
    };
  }
})();
