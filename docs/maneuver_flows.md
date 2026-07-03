# Maneuver Step Flows — TM / NM / TA

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

### NM — TR1 / TR2  *(restore from standard impedido state → normal)*

Inverse of TM Path A. Each CLOSE step re-energises a busbar via its normal source;
each OPEN step removes the backup connection established during TM.

```mermaid
flowchart TD
    subgraph G0["Step 0"]
        P0["Init — TR1 impedido → Normal"]
    end

    subgraph G1["Step 1"]
        P1["CLOSE CB[1]"]
    end

    subgraph G2["Step 2"]
        P2["OPEN TIE[2A,1B]"]
    end

    subgraph G3["Step 3"]
        P3["CLOSE TIE[2B,2A]"]
    end

    subgraph G4["Step 4"]
        P4["OPEN TIE[3A,2B]"]
    end

    subgraph G5["Step 5"]
        P5["CLOSE TIE[1B,1A]"]
    end

    subgraph G6["Step 6"]
        P6["OPEN TIE[1A,4A]"]
    end

    G0 --> P1 --> P2 --> P3 --> P4 --> P5 --> P6 --> F["Done — Normal state"]
```

> **Contingency NM paths (B and C):** correspond to the inverse of TM Paths B and C.
> They start from the dual-impedimento state left by that TM path. To be specified in `NM_TR1.md`.

### NM — TR3 / TR4  *(single path)*

```mermaid
flowchart LR
    subgraph G0["Step 0"] P0["Init"] end
    subgraph G1["Step 1"] P1["CLOSE CB[3]  or  CB[4]"] end
    subgraph G2["Step 2"] P2["OPEN TIE[4A,3A]"] end
    G0 --> P1 --> P2 --> DONE(["Done"])
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
