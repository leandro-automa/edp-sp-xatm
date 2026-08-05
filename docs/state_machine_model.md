# Modelo Genérico — State Machine de Transferência Manual (TM)

## Conceito Central

A automação TM é **orientada por dados**: o mesmo código executa qualquer layout de subestação. O `triggerTransformer` seleciona o bloco de lógica, e os `impedimentos` selecionam o caminho dentro desse bloco. Os IDs reais de equipamentos vêm de uma **tabela de layout** externa ao código.

```vba
' Estrutura principal — igual para qualquer layout
Sub ExecuteTM(triggerTransformer As Integer)
    Dim impedimentos(1 To N_TR) As Boolean
    Call ReadImpedimentos(impedimentos)   ' lê estado atual dos TRs

    Dim steps() As String                ' array de passos: "CLOSE:DEV_ID" ou "OPEN:DEV_ID"
    steps = GetTMSteps(triggerTransformer, impedimentos)

    Dim i As Integer
    For i = 0 To UBound(steps)
        Call ExecuteStep(steps(i))       ' envia comando ao equipamento
    Next i
End Sub
```

---

## Variáveis de Entrada

| Variável | Tipo | Descrição |
|----------|------|-----------|
| `triggerTransformer` | Integer | Índice do transformador a ser posto em impedimento (1-based) |
| `impedimentos(i)` | Boolean | `True` se o TR `i` está atualmente em impedimento |

---

## Conceitos Abstratos de Equipamentos

Cada layout define um mapeamento de nomes abstratos para IDs reais:

| Conceito Abstrato | Descrição | Exemplo 2T-1TIE | Exemplo 4T-Ring (Guarulhos) |
|-------------------|-----------|-----------------|------------------------------|
| `CB[n]`           | Disjuntor secundário do transformador n | CB1, CB2 | DJ03, DJ04, DJ05, DJ06 |
| `TIE[a,b]`        | Disjuntor de interligação entre barras a e b | TIE[1,2] | DJ10, DJ20, DJ30, DJ40, DJ50, DJ60 |

---

## Estrutura de um Passo de Manobra

Cada passo é uma tupla `(ação, dispositivoAbstrato)`:

```
CLOSE CB[1]       → fecha o disjuntor secundário do TR1
OPEN  TIE[1A,4A]  → abre o disjuntor de interligação entre B1A e B4A
```

O código resolve o dispositivo abstrato para o ID real consultando o dicionário do layout ativo antes de enviar o comando.

---

## Algoritmo de Seleção de Caminho

```
GetTMSteps(trigger, impedimentos):
    Select Case trigger
        Case 1:  return SelectPathTR1(impedimentos)
        Case 2:  return SelectPathTR2(impedimentos)
        ...
        Case N:  return SelectPathTRN(impedimentos)
    End Select

SelectPathTR[n](impedimentos):
    ' Cada layout define quais condições de impedimento
    ' levam a qual sequência de passos.
    ' A condição é testada em ordem de prioridade.
    If impedimentos(k):  return PATH_K_STEPS
    ElseIf impedimentos(j): return PATH_J_STEPS
    Else:                return PATH_DEFAULT_STEPS
```

---

## TA — Transferência Automática

### Diferença fundamental em relação à TM

| Aspecto | TM (Transferência Manual) | TA (Transferência Automática) |
|---------|--------------------------|-------------------------------|
| Gatilho | Operador inicia deliberadamente | Proteção dispara (relay de TR ou CR de barra) |
| Estado do TR no início | Ainda energizado e em carga | Já desligado por falta ou trip de proteção |
| Princípio de sequência | **Fechar antes de abrir** (make-before-break) | **Abrir primeiro, depois restaurar** |
| Passo 1 | FECHAR caminho de backup | ABRIR CB[trigger] (confirmar isolamento do TR) |
| Passo 2 | ABRIR ligação principal | ABRIR TIE adjacente (isolar barramento morto) |
| Passos seguintes | (redistribuição completa com sobreposição) | FECHAR TIEs para restaurar cada barramento |

### Estrutura de GetTASteps

```vba
Function GetTASteps(trigger As Integer, imp() As Boolean) As String()
    ' Mesma assinatura que GetTMSteps — mesma máquina de estados, sequência inversa
    Select Case trigger
        Case 1:  return SelectTA_PathTR1(imp)
        Case 2:  return SelectTA_PathTR2(imp)
        ...
    End Select
End Function
```

---

## Layouts Disponíveis

O layout ativo é a combinação de duas seleções independentes — `Layout.Transformer`
e `Layout.Busbar` — e `GetLayoutType()` as concatena com `_`. **Só duas combinações
têm sequência de passos implementada:**

| Chave (`GetLayoutType`) | Arquivo | Transformadores | Disjuntores de interligação |
|-------------------------|---------|-----------------|----------------------------|
| `2TR2LV_2BB1TIE`        | [layouts/2T_1TIE.md](layouts/2T_1TIE.md) | 2 | 1 TIE |
| `4TR4LV_6BB6TIERING`    | [layouts/4T_ring_guarulhos.md](layouts/4T_ring_guarulhos.md) | 4 | 6 TIEs em anel |

Qualquer outro par (`4TR4LV_2BB1TIE`, `2TR2LV_6BB6TIERING`, …) é **não suportado**.
A árvore de dispositivos pode ser montada para ele, mas nenhuma automação roda:

- **Configuração** — a tela Config avisa e pede confirmação ao aplicar
  (`btnApply`) e ao salvar (`btnSave`) uma combinação fora da lista.
- **Execução** — `Main_Step00` consulta `IsSupportedLayout()` e encerra antes de
  qualquer comando, com log `"Not executed - layout '<chave>' has no step sequence."`

---

## Princípio de Segurança: "Fechar antes de Abrir"

Toda sequência TM obedece à regra de transferência com sobreposição (make-before-break em nível de carga):

1. **Fecha** o caminho de backup (TIE ou CB de outro TR) — estabelece alimentação alternativa
2. **Abre** a ligação principal (CB do trigger TR ou TIE anterior) — remove o TR em impedimento

Nunca abrir antes de confirmar o fechamento do passo anterior.
