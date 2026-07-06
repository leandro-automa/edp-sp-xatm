# Layout: 4T-Ring — ETD Guarulhos (Anel com 6 Barramentos)

## Topologia

```
                   TR1          TR2                TR3          TR4
                    │            │                  │            │
                  CB[1]        CB[2]              CB[3]        CB[4]
                    │            │                  │            │
TIE[1A,4A]──B1A──TIE[1B,1A]──B1B──TIE[2A,1B]──B2A──TIE[2B,2A]──B2B──TIE[3A,2B]──B3A──TIE[4A,3A]──B4A──TIE[1A,4A]
    └─────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                              (anel fecha em TIE[1A,4A])
```

**Notas:**
- CB[1]/DJ03 injeta em **B1B**, não em B1A. B1A é alimentada através de TIE[1B,1A] (DJ20) no estado normal.
- B2B não tem transformador direto — alimentada via TR2→B2A→TIE[2B,2A] no estado normal.
- O diagrama unifilar (`se_gul_diagram.png`) rotula B1A incorretamente como `B1B2` — o nome correto é B1A.

## Mapeamento de Dispositivos

O **ID Lógico** é o valor da propriedade `.Id` que a lógica (`xatm_BTC`) usa — é
o contrato, independente do nome de exibição do dispositivo em `xatm_data`.

### Transformadores

| Nome GUL | ID Lógico |
|----------|-----------|
| TR1      | 100       |
| TR2      | 200       |
| TR3      | 300       |
| TR4      | 400       |

### Disjuntores

| Conceito Abstrato  | ID Real (ETD Guarulhos) | ID Lógico | Posição no Anel |
|--------------------|-------------------------|-----------|-----------------|
| `CB[1]`            | DJ03                    | 120       | Secundário TR1 → B1B |
| `CB[2]`            | DJ04                    | 220       | Secundário TR2 → B2A |
| `CB[3]`            | DJ05                    | 320       | Secundário TR3 → B3A |
| `CB[4]`            | DJ06                    | 420       | Secundário TR4 → B4A |
| `TIE[1B,1A]`       | DJ20                    | 700       | B1B ↔ B1A (= B1B2) |
| `TIE[2A,1B]`       | DJ30                    | 710       | B2A ↔ B1B |
| `TIE[2B,2A]`       | DJ40                    | 720       | B2B ↔ B2A |
| `TIE[3A,2B]`       | DJ50                    | 730       | B3A ↔ B2B |
| `TIE[4A,3A]`       | DJ60                    | 740       | B4A ↔ B3A |
| `TIE[1A,4A]`       | DJ10                    | 900       | B1A ↔ B4A (fecha o anel, NA) |

### Esquema de numeração dos IDs Lógicos

- **100–400** — vãos de transformador (`TRn = n×100`); **+20** = disjuntor
  secundário do transformador (`120 / 220 / 320 / 420`)
- **700–740** — os cinco disjuntores de interligação de barra, percorrendo o
  anel na ordem dos barramentos B1A→B4A (`700, 710, 720, 730, 740`)
- **900** — disjuntor de fechamento do anel (retorno B4A→B1A, normalmente
  aberto — sobre o qual todas as sequências TM/TA se apoiam)

## Estado Normal

| CB[1] | TIE[1B,1A] | TIE[1A,4A] | TIE[2A,1B] | CB[2] | TIE[2B,2A] | TIE[3A,2B] | CB[3] | TIE[4A,3A] | CB[4] |
|-------|-----------|-----------|-----------|-------|-----------|-----------|-------|-----------|-------|
| FECH  | FECH      | ABER      | ABER      | FECH  | FECH      | ABER      | FECH  | ABER      | FECH  |

*(LIG = Fechado/FECH, DESL = Aberto/ABER)*

Alimentação: B1A←TR1, B1B←TR1, B2A←TR2, B2B←TR2, B3A←TR3, B4A←TR4

---

## TM — Seleção de Caminho por triggerTransformer e Impedimentos

### Select Case triggerTransformer = 1 (TR1)

```
If impedimentos(2):          → Caminho B  (TR2 já impedido)
ElseIf impedimentos(3)
    Or impedimentos(4):      → Caminho C  (TR3 ou TR4 já impedido)
Else:                        → Caminho A  (sem contingência)
```

| Caminho | Condição | Passos em ordem |
|---------|----------|-----------------|
| A | Nenhum impedimento | CLOSE TIE[1A,4A] → OPEN TIE[1B,1A] → CLOSE TIE[3A,2B] → OPEN TIE[2B,2A] → CLOSE TIE[2A,1B] → OPEN CB[1] |
| B | impedimentos(2) | CLOSE TIE[1B,1A] → OPEN CB[1] → CLOSE TIE[2B,2A] → OPEN TIE[2A,1B] |
| C | impedimentos(3) ou impedimentos(4) | CLOSE TIE[1A,4A] → OPEN TIE[1B,1A] → CLOSE TIE[2A,1B] → OPEN CB[1] |

### Select Case triggerTransformer = 2 (TR2)

```
If impedimentos(1):          → Caminho B  (TR1 já impedido)
ElseIf impedimentos(3)
    Or impedimentos(4):      → Caminho C  (TR3 ou TR4 já impedido)
Else:                        → Caminho A  (sem contingência)
```

| Caminho | Condição | Passos em ordem |
|---------|----------|-----------------|
| A | Nenhum impedimento | CLOSE TIE[1A,4A] → OPEN TIE[1B,1A] → CLOSE TIE[3A,2B] → OPEN TIE[2B,2A] → CLOSE TIE[2A,1B] → OPEN CB[2] |
| B | impedimentos(1) | CLOSE TIE[1B,1A] → OPEN TIE[2A,1B] → CLOSE TIE[2B,2A] → OPEN CB[2] |
| C | impedimentos(3) ou impedimentos(4) | CLOSE TIE[1A,4A] → OPEN TIE[1B,1A] → CLOSE TIE[2A,1B] → OPEN CB[2] |

### Select Case triggerTransformer = 3 (TR3)

```
Else:                        → Caminho A  (único caminho)
```

| Caminho | Condição | Passos em ordem |
|---------|----------|-----------------|
| A | (qualquer) | CLOSE TIE[4A,3A] → OPEN CB[3] |

### Select Case triggerTransformer = 4 (TR4)

```
Else:                        → Caminho A  (único caminho)
```

| Caminho | Condição | Passos em ordem |
|---------|----------|-----------------|
| A | (qualquer) | CLOSE TIE[4A,3A] → OPEN CB[4] |

---

## Pseudocódigo Completo — GetTMSteps para este Layout

```vba
Function GetTMSteps(trigger As Integer, imp() As Boolean) As String()
    Select Case trigger

        Case 1  ' TR1
            If imp(2) Then
                GetTMSteps = Split("CLOSE:TIE[1B,1A],OPEN:CB[1],CLOSE:TIE[2B,2A],OPEN:TIE[2A,1B]", ",")
            ElseIf imp(3) Or imp(4) Then
                GetTMSteps = Split("CLOSE:TIE[1A,4A],OPEN:TIE[1B,1A],CLOSE:TIE[2A,1B],OPEN:CB[1]", ",")
            Else
                GetTMSteps = Split("CLOSE:TIE[1A,4A],OPEN:TIE[1B,1A],CLOSE:TIE[3A,2B],OPEN:TIE[2B,2A],CLOSE:TIE[2A,1B],OPEN:CB[1]", ",")
            End If

        Case 2  ' TR2
            If imp(1) Then
                GetTMSteps = Split("CLOSE:TIE[1B,1A],OPEN:TIE[2A,1B],CLOSE:TIE[2B,2A],OPEN:CB[2]", ",")
            ElseIf imp(3) Or imp(4) Then
                GetTMSteps = Split("CLOSE:TIE[1A,4A],OPEN:TIE[1B,1A],CLOSE:TIE[2A,1B],OPEN:CB[2]", ",")
            Else
                GetTMSteps = Split("CLOSE:TIE[1A,4A],OPEN:TIE[1B,1A],CLOSE:TIE[3A,2B],OPEN:TIE[2B,2A],CLOSE:TIE[2A,1B],OPEN:CB[2]", ",")
            End If

        Case 3  ' TR3
            GetTMSteps = Split("CLOSE:TIE[4A,3A],OPEN:CB[3]", ",")

        Case 4  ' TR4
            GetTMSteps = Split("CLOSE:TIE[4A,3A],OPEN:CB[4]", ",")

    End Select
End Function
```

---

## TA — Seleção de Caminho por triggerTransformer e Impedimentos

### Princípio TA (diferente de TM)

O TR já está faltoso — a barra(s) que ele alimentava já está(ão) morta(s). Sequência:
1. **OPEN CB[trigger]** — confirma isolamento
2. **OPEN TIE adjacente** — isola barramento morto
3. **CLOSE TIEs** — restaura busbars via outros transformadores

### Select Case triggerTransformer = 1 (TR1)

```
If impedimentos(2):          → Caminho B  (TR2 já impedido)
ElseIf impedimentos(3)
    Or impedimentos(4):      → Caminho C  (TR3 ou TR4 já impedido)
Else:                        → Caminho A  (sem contingência)
```

| Caminho | Condição | Passos em ordem |
|---------|----------|-----------------|
| A | Nenhum impedimento | OPEN CB[1] → OPEN TIE[1B,1A] → CLOSE TIE[1A,4A] → OPEN TIE[2B,2A] → CLOSE TIE[3A,2B] → CLOSE TIE[2A,1B] |
| B | impedimentos(2) | OPEN CB[1] → OPEN TIE[2A,1B] → CLOSE TIE[1B,1A] → CLOSE TIE[2B,2A] |
| C | impedimentos(3) ou impedimentos(4) | OPEN CB[1] → OPEN TIE[1B,1A] → CLOSE TIE[1A,4A] → CLOSE TIE[2A,1B] |

**Estado final caminho A:** B1A←TR4, B1B←TR2, B2A←TR2, B2B←TR3, B3A←TR3, B4A←TR4
**Estado final caminho B:** B1A←TR4, B1B←TR4, B2A←TR3, B2B←TR3, B3A←TR3, B4A←TR4
**Estado final caminho C:** B1A←TR4, B1B←TR2, B2A←TR2, B2B←TR2, B3A←(TR4 ou TR2), B4A←TR4

---

## Pseudocódigo Completo — GetTASteps para este Layout

```vba
Function GetTASteps(trigger As Integer, imp() As Boolean) As String()
    Select Case trigger

        Case 1  ' TR1
            If imp(2) Then
                GetTASteps = Split("OPEN:CB[1],OPEN:TIE[2A,1B],CLOSE:TIE[1B,1A],CLOSE:TIE[2B,2A]", ",")
            ElseIf imp(3) Or imp(4) Then
                GetTASteps = Split("OPEN:CB[1],OPEN:TIE[1B,1A],CLOSE:TIE[1A,4A],CLOSE:TIE[2A,1B]", ",")
            Else
                GetTASteps = Split("OPEN:CB[1],OPEN:TIE[1B,1A],CLOSE:TIE[1A,4A],OPEN:TIE[2B,2A],CLOSE:TIE[3A,2B],CLOSE:TIE[2A,1B]", ",")
            End If

        ' Case 2, 3, 4 — a definir com spec §1.4.3.x

    End Select
End Function
```

---

## Referências na Especificação

| trigger | Caminho | Seção |
|---------|---------|-------|
| TR1 | A (normal) | §1.4.1.1 |
| TR1 | B (TR2 imp.) | §1.4.2.1 |
| TR1 | C (TR3/TR4 imp.) | §1.4.2.3 / §1.4.2.5 |
| TR2 | A (normal) | §1.4.1.3 |
| TR2 | B (TR1 imp.) | §1.4.2.7 |
| TR2 | C (TR3/TR4 imp.) | §1.4.2.9 / §1.4.2.11 |
| TR3 | A | §1.4.1.5 |
| TR4 | A | §1.4.1.7 |
