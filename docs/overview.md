# ETD Guarulhos — Visão Geral da Automação

## Topologia

A ETD Guarulhos possui três conjuntos blindados (CBBS-1, CBBS-2 e CBBS-3) com **seis barramentos de 13,8 kV interligáveis em anel** e **quatro transformadores** de alimentação.

```
                    TR-01         TR-02               TR-03         TR-04
                      │             │                   │             │
                    DJ03          DJ04                DJ05          DJ06
                      │             │                   │             │
DJ10──B1A──DJ20──B1B──┘  DJ30──B2A─┘  DJ40──B2B──DJ50──┘  DJ60──B3A──┘──B4A──DJ10
  └────────────────────────────────────────────────────────────────────────────────┘
                                   (anel fecha em DJ10)
```

**Nota:** B1A e B2B não possuem transformador direto — são alimentadas pelo anel em estado normal (B1A via TR1→B1B→DJ20; B2B via TR2→B2A→DJ40).

**Nota sobre o diagrama unifilar:** O barramento `B1A` aparece rotulado como `B1B2` no arquivo `se_gul_diagram.png` — esse rótulo é um erro no desenho. O barramento correto é `B1A`.

| Barramento | Circuitos (alimentadores) | DJ transformador |
|------------|--------------------------|-----------------|
| B1A        | DJ11, DJ12, DJ13         | — (via anel)    |
| B1B        | DJ14, DJ15, DJ16         | DJ03 (TR1)      |
| B2A        | DJ21, DJ22, DJ23         | DJ04 (TR2)      |
| B2B        | DJ24, DJ25, DJ26         | — (via anel)    |
| B3A        | DJ31, DJ32, DJ33         | DJ05 (TR3)      |
| B4A        | DJ34, DJ35, DJ36         | DJ06 (TR4)      |

## Estado Normal (Situação 1)

| DJ10 | DJ20 | DJ03 | DJ30 | DJ04 | DJ40 | DJ50 | DJ05 | DJ60 | DJ06 |
|------|------|------|------|------|------|------|------|------|------|
| DESL | LIG  | LIG  | DESL | LIG  | LIG  | DESL | LIG  | DESL | LIG  |

Alimentação no estado normal:
- **B1A** ← TR1 (via DJ03 → DJ20)
- **B1B** ← TR1 (via DJ03)
- **B2A** ← TR2 (via DJ04)
- **B2B** ← TR2 (via DJ04 → DJ40)
- **B3A** ← TR3 (via DJ05)
- **B4A** ← TR4 (via DJ06)

## Tipos de Automação

| Sigla | Nome | Descrição |
|-------|------|-----------|
| **TM** | Transferência Manual | Sequência de manobras para colocar um transformador em impedimento (planejado). Trata casos com ou sem contingência de outro transformador. |
| **NM** | Normalização Manual | Retorno ao estado normal após TM. |
| **TA** | Transferência Automática | Transferência automática disparada pela atuação da CR (Chave Relé de proteção de barramento). |

## Proteção de Barra (CR)

Cada barramento possui uma chave relé biestável:

| CR         | Desliga via fiação                        |
|------------|-------------------------------------------|
| CRBA1A     | DJ10, DJ20, DJ11, DJ12, DJ13             |
| CRBA1B     | DJ03, DJ20, DJ30, DJ14, DJ15, DJ16       |
| CRBA2A     | DJ04, DJ30, DJ40, DJ21, DJ22, DJ23       |
| CRBA2B     | DJ40, DJ50, DJ24, DJ25, DJ26             |
| CRBA3A     | DJ05, DJ50, DJ60, DJ31, DJ32, DJ33       |
| CRBA4A     | DJ06, DJ60, DJ10, DJ34, DJ35, DJ36       |

## Arquivos de Documentação

| Arquivo | Conteúdo |
|---------|----------|
| `overview.md` | Este arquivo — topologia e estados |
| `TM_TR1.md` | Lógica de Transferência Manual do TR1 |
| `se_gul_diagram.png` | Diagrama unifilar da ETD Guarulhos |
| `Descritivo Automatismos ETD Guarulhos (Anexo Especificacao).pdf` | Especificação funcional completa (fonte) |
