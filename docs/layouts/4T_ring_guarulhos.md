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

### Entrada 88 kV — RAE GUL-1 e RAE GUL-2

Duas linhas da LT Norte-Nordeste alimentam as barras de entrada BA-1 e BA-2.
Os vãos de entrada ficam na pasta **Income**, e ocupam a faixa **0–99**, que
estava inteiramente livre.

**A dezena é o vão, a unidade é o papel do equipamento:**

| Vão | Equipamento | ID Real | ID Lógico |
|-----|-------------|---------|-----------|
| RAE GUL-1 | Seccionadora de linha | 3393 | **12** |
| RAE GUL-1 | **Disjuntor**         | 1    | **10** |
| RAE GUL-1 | Seccionadora → BA-1   | 3997 | **14** |
| RAE GUL-1 | Seccionadora → BA-2   | 3998 | **15** |
| RAE GUL-2 | Seccionadora de linha | 3394 | **22** |
| RAE GUL-2 | **Disjuntor**         | 2    | **20** |
| RAE GUL-2 | Seccionadora → BA-1   | 3995 | **24** |
| RAE GUL-2 | Seccionadora → BA-2   | 3996 | **25** |

> **0–9 fica reservado, e nenhum equipamento pode ter Id 0.** Em VBScript
> `Empty = 0` é verdadeiro, e um `.Id` não configurado lê como `Empty`. Um
> `Id` igual a 0 colide com todo equipamento ainda não configurado: a busca
> `If obj.Id = id` em `FindInFolder` devolveria o equipamento errado, e o
> `dict.Add item.Id` de `BuildDeviceDictionary` trata `Empty` e `0` como a
> **mesma chave** — o segundo é descartado em silêncio. Por isso o primeiro
> vão começa em 10 e não em 0.

**Atenção à numeração real:** as seccionadoras de barra do vão GUL-2
(3995/3996) têm número **menor** que as do vão GUL-1 (3997/3998) — a ordem
real corre ao contrário da ordem dos vãos. Não é engano do mapeamento.

**A confirmar no unifilar:** `3393`, `3394` e `3347` foram lidos como `3993`,
`3994` e `3447` na leitura do PDF (REV24). Os quatro números já acordados
(3995–3998) formam uma sequência contínua com 3993/3994, o que sugere que a
leitura do PDF esteja certa — confirmar antes de criar os objetos.

**Cuidado com os nomes:** `BA-1`/`BA-2` (barras de entrada 88 kV) e
`B1A`/`B1B` (barras do anel 13,8 kV) diferem por uma transposição de
caracteres. Na notação abstrata as barras de entrada são `IN[1]`/`IN[2]`,
justamente para que um passo nunca se pareça com o outro.

### Seccionadoras de alta tensão dos transformadores

Ficam na pasta **Transformer**, junto do transformador a que pertencem, e
seguem a regra `TRn = n×100`: **+30** e **+80** para as duas barras de entrada.

| Transformador | ID Real | ID Lógico | ID Real | ID Lógico |
|---------------|---------|-----------|---------|-----------|
| TR1 (100)     | 3999    | **130**   | 4000    | **180**   |
| TR2 (200)     | 4001    | **230**   | 4002    | **280**   |
| TR3 (300)     | 4003    | **330**   | 4004    | **380**   |
| TR4 (400)     | 3347    | **430**   | 3236    | **480**   |

> **A confirmar:** qual das duas — `+30` ou `+80` — liga em BA-1. No unifilar
> os transformadores ficam **abaixo** das barras, ao contrário dos vãos de
> entrada, então a geometria não resolve a dúvida sozinha.

Estas seccionadoras **não são declaradas pelo layout de transformador**.
Quantas existem depende de quantas barras de entrada há, que é assunto do
layout de entrada: são `transformadores × barras de entrada` — 4 × 2 = 8 aqui,
e nenhuma quando a entrada é `NONE`. São os equipamentos que só existem no
cruzamento dos dois layouts, e é por isso que `4TR4LV` não as menciona no nome.

### Esquema de numeração dos IDs Lógicos

- **0–9** — reservado, nunca atribuído (ver a nota sobre `Empty = 0`)
- **10–99** — vãos de entrada 88 kV, dezena = vão (`RAE GUL-1 = 10`,
  `RAE GUL-2 = 20`), unidade = papel: **+0** disjuntor, **+2** seccionadora de
  linha, **+4** BA-1, **+5** BA-2
- **100–400** — vãos de transformador (`TRn = n×100`); **+20** = disjuntor
  secundário (`120 / 220 / 320 / 420`); **+30** e **+80** = seccionadoras de
  alta tensão para as barras de entrada
- **700–740** — os cinco disjuntores de interligação de barra, percorrendo o
  anel na ordem dos barramentos B1A→B4A (`700, 710, 720, 730, 740`)
- **900** — disjuntor de fechamento do anel (retorno B4A→B1A, normalmente
  aberto — sobre o qual todas as sequências TM/TA se apoiam)

---

## Layouts

A configuração tem **três eixos independentes**, cada um com seu seletor na
tela Config e sua tag em `XATM_Data.Automation.Layout`:

| Eixo | Tag | Valores | Pasta que sincroniza |
|------|-----|---------|----------------------|
| Transformer | `Layout.Transformer` | `4TR4LV`, `2TR2LV` | `Substation.Transformer` |
| Busbar | `Layout.Busbar` | `6BB6TIERING`, `2BB1TIE` | `Substation.Busbar` |
| Income | `Layout.Income` | `2BR2BB`, `NONE` | `Substation.Income` |

**ETD Guarulhos = `4TR4LV` + `6BB6TIERING` + `2BR2BB`.**

### Como se lê o nome de um layout

Uma contagem por tipo de equipamento, mais uma palavra de topologia quando ela
distingue alguma coisa — `6BB6TIERING` são seis barras e seis disjuntores em
**anel**.

`2BR2BB` são **2 disjuntores de entrada** e **2 barras de entrada**. As
seccionadoras não entram no nome porque não são informação nova: são
`vãos × (1 linha + 1 por barra)` = 2 × 3 = **6**. Nomear uma quantidade
derivada só cria dois fatos que podem passar a discordar.

Pela mesma razão `4TR4LV` continua sem mencionar as seccionadoras de alta
tensão: elas são `transformadores × barras de entrada`, e essa conta só existe
quando os dois eixos se encontram.

### `NONE`

Uma entrada `NONE` não cria equipamento nenhum: nem os vãos de entrada, nem as
seccionadoras de alta tensão dos transformadores. É o valor para uma
subestação cuja entrada não é supervisionada por este automatismo.

### Combinações com sequência de manobra

`IsSupportedCombination` continua olhando apenas **Transformer + Busbar** —
`4TR4LV + 6BB6TIERING` e `2TR2LV + 2BB1TIE`. O eixo de entrada não entra nessa
conta porque nenhum passo de TM/TA opera equipamento de entrada; ele é validado
à parte, apenas quanto a ser um layout conhecido.

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
