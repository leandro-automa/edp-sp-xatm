# EDP SP — Substation Automation State Machine

## What this project is

State machine logic for power substation automation (EDP/Bandeirante, São Paulo). The code drives **TM** (Transferência Manual), **NM** (Normalização Manual), and **TA** (Transferência Automática) switching sequences for 13.8 kV GIS substations.

Primary target substation: **ETD Guarulhos** (`se_gul`) — 4-transformer ring busbar layout.

## Key conventions

### Abstract device notation

Logic is written in abstract IDs, resolved to real equipment per layout:

| Abstract | Meaning |
|----------|---------|
| `CB[n]` | Secondary breaker of transformer n |
| `TIE[a,b]` | Bus-tie breaker between busbar a and busbar b |

Real mappings are in `docs/layouts/`. For ETD Guarulhos: `docs/layouts/4T_ring_guarulhos.md`.

### Step encoding

Each maneuver step is a string `"ACTION:DEVICE"`:
- `"CLOSE:TIE[1A,4A]"` — close the tie between B1A and B4A
- `"OPEN:CB[1]"` — open transformer 1's secondary breaker

### Portuguese ↔ English

| PT | EN |
|----|----|
| LIG / FECH / FECHADO | CLOSED |
| DESL / ABER / ABERTO | OPEN |
| Fechar (F) | Close |
| Abrir (A) | Open |
| Impedimento | Out of service (planned) |

## Automation principles

- **TM**: transformer is still live. **Close backup path first**, then open primary (make-before-break).
- **TA**: transformer has already tripped. **Open CB first** to confirm isolation, then open adjacent tie to isolate dead busbar, then restore via alternate transformers.
- **NM**: inverse of TM — restore normal state after impediment is cleared.

## Entry points for logic

- `GetTMSteps(trigger, imp())` — returns ordered step array for manual transfer
- `GetTASteps(trigger, imp())` — returns ordered step array for automatic transfer
- `ExecuteStep(step)` — resolves abstract ID to real equipment ID and sends command

## Key files

| File | Purpose |
|------|---------|
| `docs/overview.md` | Topology, normal state, busbar-to-circuit mapping, CR protection table |
| `docs/state_machine_model.md` | Abstract model, TM vs TA principles, VBA dispatcher pattern |
| `docs/layouts/4T_ring_guarulhos.md` | Full device mapping + TM/TA step sequences for se_gul |
| `docs/layouts/2T_1TIE.md` | Simplest layout — 2 transformers + 1 TIE |
| `docs/TM_TR1.md` | Detailed flowchart and step tables for TM TR1 (Guarulhos) |

## Implementation language

VBA-style pseudocode (target: SCADA/automation controller). Uses `Dim`, `Select Case`, `Split()`, 1-based arrays.
