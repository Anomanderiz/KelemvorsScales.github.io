## Web App (GitHub Pages)

This repo now includes a browser app in `docs/` built for GitHub Pages.

- Entry point: `docs/index.html`
- Main logic: `docs/app.js`
- Styling: `docs/styles.css`
- Icon/favicons: `docs/assets/icon.ico`
- GitHub Pages deploy workflow: `.github/workflows/pages.yml`

### Local preview

```bash
python -m http.server --directory docs 8000
```

Then open `http://localhost:8000`.

### Deploy on GitHub Pages

1. Push this repo to GitHub on branch `main`.
2. In repo settings, open **Pages** and set **Build and deployment** to **GitHub Actions**.
3. The workflow **Deploy GitHub Pages** will publish the `docs/` app.

Desktop Python app files are still present; the Pages app is the web-hosted version.

---
<img width="500" height="500" alt="Untitled design (1)" src="https://github.com/user-attachments/assets/58770abf-07a2-4e51-ac7b-3ad82d6d8aaa" />

# Kelemvorâ€™s Scales
**Encounter Monte Carlo + Auto-Tune (Crunchy Crits)**

Balance D&D 5e boss fights by simulating entire combats, not just average DPR. Get pacing guardrails (no steamrolls, no slogs), lethality estimates, and an auto-tuner that hits your target time-to-kill.

> *Design goal:* â€œPerilous but fair.â€ Median drama without coin-flip swing or endless grinds.

---
<img width="1594" height="991" alt="image" src="https://github.com/user-attachments/assets/38983781-240d-4a7a-9050-8c6a3201e0af" />
<img width="2207" height="1327" alt="image" src="https://github.com/user-attachments/assets/4ea078f8-897d-49b2-b692-55e85d428cc4" />

# Quick UI Tour

## ðŸ‘¥ Party & DPR

Enter PCsâ€™ AC, HP, and save bonuses (STR/DEX/CON/INT/WIS/CHA).

Choose damage input mode:

Manual Effective DPR (enter your own per-round average)

Novaâ†’Effective (enter â€œNova DPRâ€, attack bonus, target AC, crit ratio, uptime)

## ðŸ§° Boss Kit

Add each attack: Type (attack/save), to-hit or DC, damage dice, uses/round, melee flag.

Optional: Lair (avg dmg, #targets, cadence), Recharge (e.g., 5-6, avg dmg, #targets).

Rider: on-hit effect that grants advantage next round or reduces target AC; duration in rounds.

Expected Party Healing / Round is a party-wide sustain average. The calculator divides it evenly across active PCs; it raises the boss-pressure budget without changing boss HP or the target encounter duration.

## ðŸ“ˆ Deterministic DPR

Fast, average-based snapshot of boss damage versus each PC.

## â³ Boss TTD

Computes time to boss defeat from total incoming effective DPR, resistance, and regen. **Boss Blitz** is shown first as the baseline (all party damage goes to the boss while the full pack keeps attacking); **Minion Clear** is the secondary DM-facing possibility where minion HP absorbs party actions.

## ðŸŽ¯ Bossâ†’PC MC

Monte Carlo of boss damage into a single PC over N rounds (useful for burst safety checks).

## âš”ï¸ Encounter MC

Full fight simulation with Trials, Max Rounds, Initiative, Party DPR CV, and Use Novaâ†’Effective DPR toggle. Every analysis runs the Boss-Blitz baseline and Minion-Clear comparison side by side, including separate survival curves, TTK distributions, casualty timing, and TPK risk.

The Balancer display is split into two layers. **Deterministic Budget & Diagnostics** shows repeatable average-damage HP and pressure calculations used to propose the encounter. **Monte Carlo Validation** is the final authority for the configured probability goals and shows dedicated Boss-Blitz cards for win-by-target, unassisted first-down risk, TPK risk, and late TPK risk among fights still unresolved after the target round.

Tune Everything treats durability and lethality as separate budgets. Boss HP is based on party DPR against the boss alone for a 6–9 round target (default 8), so Boss Blitz never makes the boss itself cheaper. With Boss HP Pacing enabled, the boss cannot fall before round 6; four Legendary Resistances also smooth exceptional pre-target burst.

Encounter damage is calibrated on the Boss-Blitz baseline against one goal: by round 6, the typical (median) PC is down or at 25% HP or less without healing (recurring THP still counts). Using the median PC means a single low-AC or low-HP outlier drops earlier rather than capping boss pressure for the whole party. Full Analysis flags fights that run past round 11 in more than 10% of trials. The configured TPK ceiling applies to both Boss Blitz and Minion Clear and cannot be set above 25%.

Replenishing minions support one- or two-round wave cadence. Auto-tuning never reduces them below 20 HP. If an every-round Minion-Clear comparison is still over the TPK ceiling at 20 HP, their wave cadence shifts to every other round before boss HP is touched.

Outcome summary (victory, wipe, or still fighting), typical resolution rounds, boss HP remaining on failures, cumulative outcome curves, and final outcome mix.

How to read the numbers (plain language)

Victory timing: the median and p10â€“p90 rounds among party victories. Always read it beside the unconditional win/wipe/unresolved rates; rare victories are not presented as the typical encounter outcome.

p10â€“p90: middle 80% of outcomes.

p10 too low â†’ too many steamrolls.

p90 too high â†’ slog risk.

TPK Probability: fights where all PCs drop before the boss dies.

PCs down at victory: how bruising a win is (mean and p90).

Guardrails: compare the outputs to your thresholds (fast/slow caps, K-down target, TPK cap).

The graphs (what they mean)
Survival curve 
ð‘†(ð‘¡)=ð‘ƒ(TTK>ð‘¡)
S(t)=P(TTK>t)

A step that drops at each round boundary.

Read tails directly:

ð‘ƒ(TTKâ‰¤2)=1âˆ’ð‘†(2)
P(TTKâ‰¤2)=1âˆ’S(2) (steamroll frequency)

ð‘ƒ(TTKâ‰¥7)=ð‘†(6)
P(TTKâ‰¥7)=S(6) (slog frequency)

The round where it crosses 0.5 is the median.

Long, shallow tail? Slogs. Cliff-like drop? Consistent pacing.

TTK distribution

A bar-like histogram (TTK is discrete: 3, 4, 5, â€¦).

Spike heights â‰ˆ how often fights end on that round.

Bimodal shape â†’ the fight swings between two regimes (e.g., â€œalpha-strikeâ€ vs â€œattritionâ€).
Adjust kit/initiative/riders to smooth it.

## The math (lightly)

Monte Carlo: simulate thousands of fights with dice randomness. Summarize outcomes (medians, percentiles, probabilities).

Damage variance (Party DPR CV): we sample each PCâ€™s per-round damage from a Gamma distribution with:

mean = your effective DPR,

CV (coefficient of variation) you set; shape = 1/CVÂ², scale = mean/shape.

Hit chance: attack roll vs AC with nat-1 auto miss, nat-20 auto hit (crit).

Crunchy crits: on crit, damage = max dice + one normal roll + modifier once.

Save-for-half: expected damage = 0.5+0.5â‹…ð‘ƒ(fail) times average damage.

Initiative: if the boss acts first, PCs might lose actions that round if they drop before swinging.

Resistance is applied to incoming boss damage; boss regeneration occurs on the boss turn. THP is per target, while expected healing is entered as a party total and divided evenly across active PCs.

Targeting: every individual attack use independently selects a random living PC. Each save ability has its own target count and independently selects up to that many distinct living PCs whenever it fires. Minion attacks and saves follow the same rules. Minion Clear reduces the number of acting minions as the pack takes damage; Boss Blitz leaves the full configured count active until the boss dies.

Balancing playbook

Pick your vibe: set Target median (e.g., 4.0).

Run Encounter MC â†’ check p10â€“p90:

Want fewer steamrolls? Raise early boss sturdiness (resist/HP), reduce party CV, or tweak initiative.

Want fewer slogs? Lower boss HP/resist slightly, add lair/recharge that closes fights.

Check lethality: the unassisted first-down chance should become meaningful by round 6, while TPK stays at or below the configured ceiling (25% maximum).

Tune Everything to set the boss-only HP budget, calibrate encounter damage, then tune ignored-minion pressure against the round-6 first-casualty target.

Re-run until guardrails show âœ“.

---

## TL;DR

- **Download & run (Windows):** grab the latest release, unzip, double-click `KelemvorsScales.exe`.
- **Basic workflow (5 minutes):**
  1. **Party & DPR** tab: enter each PCâ€™s **AC**, **HP**, and **save bonuses**; add either **Manual DPR** or **Novaâ†’Effective** rows.
  2. **Boss Kit** tab: add the bossâ€™s attacks (to-hit/DC, damage, save targets, uses/round, melee?), plus **lair**, **recharge**, and any **on-hit rider**.
  3. **âš”ï¸ Encounter MC** tab: set **Trials** (10â€“30k), **Max Rounds** (~13), **Initiative** (random), and the **Party DPR CV** (variance, default 0.60).
  4. Click **Run Encounter Simulation**. Read each strategy's **party win**, **party wipe**, and **still fighting** rates, then use typical resolution and remaining boss HP to see why failures occurred.
  5. If pacing is off, hit **Tune Everything** to set boss HP for a 6–9 round target and pressure so the typical PC is down or near down by round 6.

---

## Features

- **Encounter Monte Carlo:** full fights under the Boss-Blitz baseline and Minion-Clear comparison, round-by-round: initiative, attacks, saves, crits (house-rule: *max dice + one rolled*), lair cadence, recharge chance, riders, minions, per-target THP/round, expected party healing, regen/resistance, and PCs dropping.
- **Tune Everything:** sets boss HP from boss-only party DPR, independently scales encounter pressure so the typical PC is down or near down by round 6, and tunes ignored-minion pressure without allowing either strategy above the TPK ceiling. Replenishing minions retain at least 20 HP and fall back to every-other-round waves when needed.
- **Guardrail read-outs:** configurable thresholds with âœ“/âœ—:
  - `P(TTK â‰¤ fast) â‰¤ cap` (avoid steamrolls)
  - `P(TTK â‰¥ slow) â‰¤ cap` (avoid slogs)
  - `P(â‰¥ K PCs down at victory) â‰ˆ target` (lethality band)
  - `P(TPK) â‰¤ cap`
- **Charts:** cumulative party victories and wipes by round, plus the unconditional final outcome mix for both strategies.
- **Per-use targeting:** every attack independently samples from **all living PCs**; each save ability samples its configured number of distinct targets. Minions use the same rules.
- **Crunchy crits:** crit = *max dice + one rolled* (modifier once) for attack rolls; save-based effects donâ€™t crit (RAW).

---

## Install

### Option A â€” Download a release (Windows)
1. Head to **Releases** on this repo.
2. Download `KelemvorsScales-win64.zip`.
3. Unzip anywhere you like; run `KelemvorsScales.exe`.

> **Note:** Windows Defender sometimes grumbles about unsigned PyInstaller apps. If it blocks you:
> - Prefer `More info â†’ Run anyway`, or
> - Add an **exclusion** for the app folder (safer than turning Defender off).
> - We do **not** recommend disabling AV system-wide.

### Option B â€” Run from source (Windows/macOS/Linux)
```bash
python -m venv .venv
# Activate the venv (Windows)
.venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate

pip install -r requirements.txt
python boss_balance_desktop_encounter_v5.py

