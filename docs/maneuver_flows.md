# Maneuver Step Flows — TM / NM / TA / RASEAT

Step sequences for each automation type and layout.
Each step executes only after the previous breaker position is confirmed.

---

## 2T + 1 TIE

```
  TR1       TR2
   │         │
  CB[1]    CB[2]
   │         │
  B1 ──TIE── B2
```

### TM — Manual Transfer  *(make-before-break)*

```mermaid
flowchart LR
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE TIE[1,2]"] end
    subgraph G2["Step 2"] P2["OPEN CB[trigger]"] end
    G0 --> P1 --> P2 --> DONE(["Done"])
```

### NM — Manual Restore  *(break-before-make)*

```mermaid
flowchart LR
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE CB[trigger]"] end
    subgraph G2["Step 2"] P2["OPEN TIE[1,2]"] end
    G0 --> P1 --> P2 --> DONE(["Done"])
```

### TA — Automatic Transfer  *(CB already tripped)*

```mermaid
flowchart LR
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["OPEN CB[trigger]"] end
    subgraph G2["Step 2"] P2["CLOSE TIE[1,2]"] end
    G0 --> P1 --> P2 --> DONE(["Done"])
```

> **Note:** If both transformers are simultaneously out of service this layout has no valid
> transfer path — TA/TM is blocked by pre-condition check.

---

## 4T Ring — ETD Guarulhos (se_gul)

```
              TR1          TR2                TR3          TR4
               │            │                  │            │
             CB[1]        CB[2]              CB[3]        CB[4]
               │            │                  │            │
TIE[1A,4A]─B1A─TIE[1B,1A]─B1B─TIE[2A,1B]─B2A─TIE[2B,2A]─B2B─TIE[3A,2B]─B3A─TIE[4A,3A]─B4A─(back to B1A)
```

Normal open point: **TIE[1A,4A]** (ring closure, normally open).

---

### TM — TR1 / TR2  *(3 paths depending on contingency)*

Path selection and step structure are the same for TR1 and TR2 — only the final `OPEN CB[n]` differs.
Diagram below uses TR1 notation; replace `CB[1]` with `CB[2]` for TR2.

```mermaid
flowchart TD
    subgraph G0["Step 0"]
        P0["Init — read impedimentos"]
    end

    subgraph G1["Step 1"]
        P1A["CLOSE TIE[1A,4A]"]
        P1B["CLOSE TIE[1B,1A]"]
    end

    subgraph G2["Step 2"]
        P2A["OPEN TIE[1B,1A]"]
        P2B["OPEN CB[1]"]
    end

    subgraph G3["Step 3"]
        P3A["CLOSE TIE[3A,2B]"]
        P3B["CLOSE TIE[2B,2A]"]
        P3C["CLOSE TIE[2A,1B]"]
    end

    subgraph G4["Step 4"]
        P4A["OPEN TIE[2B,2A]"]
        P4B["OPEN TIE[2A,1B]"]
        P4C["OPEN CB[1]"]
    end

    subgraph G5["Step 5"]
        P5A["CLOSE TIE[2A,1B]"]
    end

    subgraph G6["Step 6"]
        P6A["OPEN CB[1]"]
    end

    G0 -->|"imp(2) — TR2 already out"| P1B
    G0 -->|"no contingency\nTR3 or TR4 already out"| P1A

    P1B --> P2B
    P1A --> P2A

    P2B --> P3B
    P2A -->|"TR3 or TR4 already out"| P3C
    P2A -->|"no contingency"| P3A

    P3B --> P4B
    P3C --> P4C
    P3A --> P4A

    P4A --> P5A
    P5A --> P6A

    P4B --> F["Done"]
    P4C --> F
    P6A --> F
```

### TM — TR3 / TR4  *(single path)*

```mermaid
flowchart LR
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE TIE[4A,3A]"] end
    subgraph G2["Step 2"] P2["OPEN CB[3]  or  CB[4]"] end
    G0 --> P1 --> P2 --> DONE(["Done"])
```

---

NM (Normalização Manual) is the inverse of TM — it restores the substation to the
normal state. The spec tables are read **bottom-to-top** (the reverse of TM). Step
actions still alternate close/open (odd close, even open), and **Step 1 always
re-closes the trigger transformer's own secondary breaker** regardless of contingency.

> Verified against spec §1.4.1.2/.4/.6/.8 (normal) and §1.4.2.2…§1.4.2.24 (contingencies).

### NM — TR1  *(3 paths)*

```mermaid
flowchart TD
    subgraph G0["Step 0"]
        P0["Init — read impedimentos"]
    end

    subgraph G1["Step 1"]
        P1["CLOSE CB[1]"]
    end

    subgraph G2["Step 2"]
        P2A["OPEN TIE[2A,1B]"]
        P2B["OPEN TIE[1B,1A]"]
    end

    subgraph G3["Step 3"]
        P3A["CLOSE TIE[2B,2A]"]
        P3B["CLOSE TIE[2A,1B]"]
        P3C["CLOSE TIE[1B,1A]"]
    end

    subgraph G4["Step 4"]
        P4A["OPEN TIE[3A,2B]"]
        P4B["OPEN TIE[2B,2A]"]
        P4C["OPEN TIE[1A,4A]"]
    end

    subgraph G5["Step 5"]
        P5A["CLOSE TIE[1B,1A]"]
    end

    subgraph G6["Step 6"]
        P6A["OPEN TIE[1A,4A]"]
    end

    G0 --> P1

    P1 -->|"imp(2) — TR2 out"| P2B
    P1 -->|"no contingency\nTR3 or TR4 out"| P2A

    P2B --> P3B
    P2A -->|"TR3 or TR4 out"| P3C
    P2A -->|"no contingency"| P3A

    P3A --> P4A
    P3B --> P4B
    P3C --> P4C

    P4A --> P5A --> P6A

    P4B --> F["Done"]
    P4C --> F
    P6A --> F
```

### NM — TR2  *(3 paths)*

```mermaid
flowchart TD
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE CB[2]"] end

    subgraph G2["Step 2"]
        P2A["OPEN TIE[2A,1B]"]
        P2B["OPEN TIE[2B,2A]"]
    end
    subgraph G3["Step 3"]
        P3A["CLOSE TIE[2B,2A]"]
        P3B["CLOSE TIE[2A,1B]"]
    end
    subgraph G4["Step 4"]
        P4A["OPEN TIE[3A,2B]"]
        P4B["OPEN TIE[1B,1A]"]
    end
    subgraph G5["Step 5"] P5A["CLOSE TIE[1B,1A]"] end
    subgraph G6["Step 6"] P6A["OPEN TIE[1A,4A]"] end

    G0 --> P1
    P1 -->|"imp(1) — TR1 out"| P2B
    P1 -->|"no contingency\nTR3 or TR4 out"| P2A
    P2B --> P3B
    P2A --> P3A
    P3B --> P4B
    P3A --> P4A
    P4A -->|"no contingency"| P5A
    P4A -->|"TR3 or TR4 out"| F["Done"]
    P5A --> P6A
    P4B --> F
    P6A --> F
```

### NM — TR3  *(4 paths — contingencies all distinct)*

```mermaid
flowchart TD
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE CB[3]"] end

    subgraph G2["Step 2"]
        P2A["OPEN TIE[4A,3A]"]
        P2B["OPEN TIE[3A,2B]"]
    end
    subgraph G3["Step 3"]
        P3A["CLOSE TIE[3A,2B]"]
        P3B["CLOSE TIE[1A,4A]"]
        P3C["CLOSE TIE[4A,3A]"]
    end
    subgraph G4["Step 4"]
        P4A["OPEN TIE[2B,2A]"]
        P4B["OPEN TIE[1B,1A]"]
        P4C["OPEN TIE[1A,4A]"]
    end

    G0 --> P1
    P1 -->|"imp(4) — TR4 out"| P2B
    P1 -->|"no contingency\nTR1 or TR2 out"| P2A
    P2A -->|"no contingency"| F["Done"]
    P2A -->|"imp(1) — TR1 out"| P3A
    P2A -->|"imp(2) — TR2 out"| P3B
    P2B --> P3C
    P3A --> P4A
    P3B --> P4B
    P3C --> P4C
    P4A --> F
    P4B --> F
    P4C --> F
```

### NM — TR4  *(4 paths — contingencies all distinct)*

```mermaid
flowchart TD
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE CB[4]"] end

    subgraph G2["Step 2"]
        P2A["OPEN TIE[4A,3A]"]
        P2B["OPEN TIE[1A,4A]"]
    end
    subgraph G3["Step 3"]
        P3A["CLOSE TIE[3A,2B]"]
        P3B["CLOSE TIE[1A,4A]"]
        P3C["CLOSE TIE[4A,3A]"]
    end
    subgraph G4["Step 4"]
        P4A["OPEN TIE[2B,2A]"]
        P4B["OPEN TIE[1B,1A]"]
        P4C["OPEN TIE[3A,2B]"]
    end

    G0 --> P1
    P1 -->|"imp(3) — TR3 out"| P2B
    P1 -->|"no contingency\nTR1 or TR2 out"| P2A
    P2A -->|"no contingency"| F["Done"]
    P2A -->|"imp(1) — TR1 out"| P3A
    P2A -->|"imp(2) — TR2 out"| P3B
    P2B --> P3C
    P3A --> P4A
    P3B --> P4B
    P3C --> P4C
    P4A --> F
    P4B --> F
    P4C --> F
```

---

### TA — TR1  *(3 paths — CB already tripped by protection)*

All paths share Step 1 `OPEN CB[1]` (position confirmation of the tripped breaker) and
Steps 2–3 for Paths A and C. Paths diverge at Step 4.

```mermaid
flowchart TD
    subgraph G0["Step 0"]
        P0["Init — read impedimentos"]
    end

    subgraph G1["Step 1"]
        P1["OPEN CB[1]"]
    end

    subgraph G2["Step 2"]
        P2A["OPEN TIE[1B,1A]"]
        P2B["OPEN TIE[2A,1B]"]
    end

    subgraph G3["Step 3"]
        P3A["CLOSE TIE[1A,4A]"]
        P3B["CLOSE TIE[1B,1A]"]
    end

    subgraph G4["Step 4"]
        P4A["OPEN TIE[2B,2A]"]
        P4B["CLOSE TIE[2B,2A]"]
        P4C["CLOSE TIE[2A,1B]"]
    end

    subgraph G5["Step 5"]
        P5A["CLOSE TIE[3A,2B]"]
    end

    subgraph G6["Step 6"]
        P6A["CLOSE TIE[2A,1B]"]
    end

    G0 --> P1

    P1 -->|"imp(2) — TR2 already out"| P2B
    P1 -->|"no contingency\nTR3 or TR4 already out"| P2A

    P2A --> P3A
    P2B --> P3B

    P3A -->|"no contingency"| P4A
    P3A -->|"TR3 or TR4 already out"| P4C
    P3B --> P4B

    P4A --> P5A
    P5A --> P6A

    P4B --> F["Done"]
    P4C --> F
    P6A --> F
```

> **TA TR2 / TR3 / TR4:** same structural pattern, different TIE set. To be specified separately.

---

## Incomer 88/138 kV — RASEAT  *(automatic reclosing)*

```
   BA-2 ─────────────┬──────────────────────┬─────────────
   BA-1 ────────┬────┼─────────────────┬────┼─────────────
                │    │                 │    │
             BA1A0  BA1B0           BA2A0  BA2B0
                └─┬──┘                 └─┬──┘
              RM01DISJ (10)          RM02DISJ (20)
```

The station runs with **one incomer closed and the other open**. A busbar
`CR` on any transformer trips the incomer that was carrying the station *and*
that transformer's secondary breaker, so the whole entry goes dark — RASEAT
is what brings it back.

`DJ[p]` is whichever incomer was carrying the station, `DJ[b]` the other.
Both are worked out at Step 0 from `MemorizedPosition`; neither is
configured, and the sequence never re-reads them, so no step can be handed a
different answer than the step before it.

### RASEAT — reclosing the entry

```mermaid
flowchart TD
    subgraph G0["Step 0"]
        P0["Init — snapshot the CR set,\nand DJ[p] from MemorizedPosition"]
    end

    subgraph G1["Step 1"]
        P1["WAIT both incomers OPEN"]
    end

    subgraph G2["Step 2"]
        P2["WAIT every TR with CR → Isolated"]
    end

    subgraph G3["Step 3"]
        P3["DWELL LatchDelay"]
    end

    subgraph G4["Step 4"]
        P4["CLOSE DJ[p]"]
    end

    subgraph G5["Step 5"]
        P5["WAIT DJ[p] HasLoadCurrent"]
    end

    subgraph G6["Step 6"]
        P6["CLOSE DJ[b]"]
    end

    G0 --> P1

    P1 -->|"both open"| P2
    P1 -->|"BreakerTimeout"| BAD

    P2 -->|"all isolated"| P3
    P2 -->|"IsolationTimeout"| BAD

    P3 --> P4

    P4 -->|"position closed"| OK
    P4 -->|"BreakerTimeout"| P5

    P5 -->|"current flowing"| OK
    P5 -->|"CurrentTimeout"| P6

    P6 -->|"position closed"| OK
    P6 -->|"BreakerTimeout\nor DJ[b] not permitted"| BAD

    OK(["Successful"])
    BAD(["Unsuccessful\nGeneralBlock"])
```

| Step | Waits for | Timeout | On timeout |
|------|-----------|---------|------------|
| 1 | both incomers open | `BreakerTimeout` | fail |
| 2 | every TR carrying `CR` reports `Isolated` | `IsolationTimeout` | fail |
| 3 | dwell for the bistable mechanism | `LatchDelay` | — |
| 4 | `DJ[p]` position closed | `BreakerTimeout` | → Step 5 |
| 5 | `DJ[p]` load current | `CurrentTimeout` | → Step 6 |
| 6 | `DJ[b]` position closed | `BreakerTimeout` | fail |

**Why this one carries timeouts and the diagrams above do not.** A TM/NM/TA
path is a fixed run of commands once the trigger and contingency are known,
so the only branch worth drawing is the path choice. Three of RASEAT's steps
are *waits* instead, and what happens when the wait expires is part of the
maneuver rather than an error case — Step 4 falling through to Step 5, and
Step 5 to Step 6, are the scheme working as designed.

**Step 2 is a fan-out drawn as one node.** It waits on *every* transformer
carrying `CR`, not one. Several transformers can trip together, and that is
not an ambiguity to resolve but simply a longer step — which is also why
`IsolationTimeout` has to cover the slowest of them, including two motorised
disconnectors travelling before a transformer can call itself isolated.

**Steps 4 and 5 are one close, confirmed two ways.** A breaker that reached
position and one only load current can vouch for are different outcomes, and
the log says which — so each gets a step of its own and its own
`StepExecutionFailed` latch.

---

## 4T Ring — busbar transfers  *(spec §1.4.1.9–§1.4.1.12)*

A different kind of maneuver from everything above. The transformer sequences move
a **transformer's** load off itself and leave the busbars where they are; these move
a **busbar's** supply to the far side of the ring and leave every transformer where
it is. Two ties, no contingency, one fixed path.

Only two busbars have somewhere else to go, so there are only two of them — and each
is a pair, out and back:

| Maneuver | Spec | Busbar | Fed by, before | Fed by, after | Step 1 — close | Step 2 — open |
|----------|------|--------|----------------|---------------|----------------|---------------|
| TM B1A-B4A | §1.4.1.9  | B1A | TR1, through `TIE[1B,1A]` | **TR4**, through the ring closer | `TIE[1A,4A]` DJ10 · 900 | `TIE[1B,1A]` DJ20 · 700 |
| NM B1A-B4A | §1.4.1.10 | B1A | TR4 | **TR1** again | `TIE[1B,1A]` DJ20 · 700 | `TIE[1A,4A]` DJ10 · 900 |
| TM B2B-B3A | §1.4.1.11 | B2B | TR2, through `TIE[2B,2A]` | **TR3**, through `TIE[3A,2B]` | `TIE[3A,2B]` DJ50 · 730 | `TIE[2B,2A]` DJ40 · 720 |
| NM B2B-B3A | §1.4.1.12 | B2B | TR3 | **TR2** again | `TIE[2B,2A]` DJ40 · 720 | `TIE[3A,2B]` DJ50 · 730 |

The state numbers the spec's Tabela 1 gives these are **12** (B1A transferida) and
**13** (B2B transferida).

```mermaid
flowchart TD
    CMD(["CommandStartTM &nbsp; or &nbsp; CommandStartNM<br/>on an instance with a BusbarPair set"])
    G1{"Layout is<br/>4TR4LV_6BB6TIERING?"}
    G2{"BusbarPair names a pair<br/>this layout can transfer?"}
    G3{"Enabled · not running · no block<br/>PreconditionsTM/NM · AutomaticBlockTM/NM<br/>no other automation running"}
    REJ(["Rejected — logged with the reason"])

    S1["Step 1 — CLOSE the tie that will feed the busbar"]
    S2["Step 2 — OPEN the tie feeding it now"]
    OK(["Done"])

    CMD --> G1
    G1 -- no --> REJ
    G1 -- yes --> G2
    G2 -- no --> REJ
    G2 -- yes --> G3
    G3 -- no --> REJ
    G3 -- yes --> S1
    S1 --> S2 --> OK
```

Each step waits for the breaker's own `Position` to reach the commanded state before
moving on, and a step that will not execute goes to global lockout — the same as
every other maneuver. There is no step 3: step 2 sets the FSM straight to 99.

> **Make before break, both ways round.** The busbar is live throughout and stays
> live; all that separates a transfer from its normalisation is which of the two ties
> closes first.
>
> This is worth stating because the spec's normalisation tables look like the
> opposite at a glance. §1.4.1.10 and §1.4.1.12 carry an **upward arrow** in the
> left-hand column and are read from the bottom row up, so what reads top-down as
> "DESLIGA DJ10, LIGA DJ20" is really *close DJ20, then open DJ10*. The transfer
> tables carry a downward arrow and read the ordinary way.

### Which automation runs them

Not a new set of commands — a new *kind of instance*. `xatm_TMTNM` carries a
`BusbarPair` property; where it names a pair, that instance moves those busbars
instead of a transformer's load, and the two commands it already had mean the
busbar sequences:

| On a normal instance | On one with a `BusbarPair` |
|---|---|
| `CommandStartTM` — take the transformer's load off | hand the busbar to the other side of the ring |
| `CommandStartNM` — bring it back | bring the busbar back |
| `CommandStartTM100`…`400` — with one transformer out | *rejected* — a busbar transfer has no contingency |
| `TA`, from the transformer's own trip | *rejected* — and the trigger never routes to it |

Two more instances, created and kept by `SyncAutomation` whenever the busbar layout
is `6BB6TIERING`, alongside `TMTNM1`…`TMTNM4`:

| Instance | `BusbarPair` | `Transformer` | Runs |
|----------|--------------|---------------|------|
| `TMTNM_B1A` | `B1A-B4A` | *unbound* | TM/NM B1A-B4A |
| `TMTNM_B2B` | `B2B-B3A` | *unbound* | TM/NM B2B-B3A |

**`Transformer` is left unbound.** A busbar is not an XObject there is anything to
link to, so which pair the instance moves is carried as text rather than as a link —
and it is not any transformer's automation, so there is nothing it should be bound to.
`RequestTA` inside `xatm_Transformer` looks only for `xatm_TA`, so a trip can never
reach one of these, and `StartMode` asks for no binding before it writes.

**One property, not two.** A flag saying "this is a busbar automation" beside a name
saying which busbar would be two ways of writing one fact, and two ways of writing one
fact can be made to disagree. A pair name that is either there or is not answers both
questions.

The instances are named rather than numbered: there is nothing to count — a layout
either defines the pair or defines neither — and a trailing number would put them in
the same series the prune walks by number. **A two-busbar layout defines neither**, and
`SyncAutomation` creates no such instance there.

The gates are the ordinary ones — `PreconditionsTM` / `AutomaticBlockTM` /
`PreconditionsNM` / `AutomaticBlockNM` — read off the same string as always, and
alarmed on the same pairs. Nothing about these four maneuvers is declared on the class
beyond the one property.

---

## Step Count Summary

| Layout | Mode | Trigger | Path | Steps |
|--------|------|---------|------|-------|
| 2T+1TIE | TM / NM / TA | any | — | 2 |
| 4T Ring | TM | TR1 / TR2 | A (no contingency) | 6 |
| 4T Ring | TM | TR1 / TR2 | B or C (1 contingency) | 4 |
| 4T Ring | TM | TR3 / TR4 | — | 2 |
| 4T Ring | NM | TR1 / TR2 | A | 6 |
| 4T Ring | NM | TR3 / TR4 | — | 2 |
| 4T Ring | TA | TR1 | A (no contingency) | 6 |
| 4T Ring | TA | TR1 | B or C (1 contingency) | 4 |
| 4T Ring | TM / NM | busbar B1A ↔ B4A | — | 2 |
| 4T Ring | TM / NM | busbar B2B ↔ B3A | — | 2 |
| 2BR2BB | RASEAT | any TR `CR` | primary confirms on position | 4 |
| 2BR2BB | RASEAT | any TR `CR` | primary confirms on current | 5 |
| 2BR2BB | RASEAT | any TR `CR` | falls back to the other incomer | 6 |
