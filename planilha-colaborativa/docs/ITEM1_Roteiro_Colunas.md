# ITEM 1 — Alteração de colunas + migração de dados
### Aplicar na **cópia de teste** do Excel Online

> ⚠️ Siga a ordem. As exclusões vão da **direita para a esquerda** de propósito: assim as letras das colunas seguintes não mudam durante o processo.
> O Excel Online reajusta sozinho as fórmulas da GERAL e do DASHBOARD ao apagar colunas. As que apontarem para colunas apagadas viram `#REF!` — isso é esperado e será tratado na Etapa 4.

---

## ETAPA 1 — Apagar 10 colunas da aba PROGRAMAÇÃO

Apague **uma a uma, nesta ordem** (clique com o botão direito no cabeçalho da coluna → *Excluir coluna*):

| Ordem | Coluna | Cabeçalho (linha 4) |
|---|---|---|
| 1 | **BK** | DATA ÚLTIMA ALTERAÇÃO |
| 2 | **BH** | STATUS GERAL |
| 3 | **BD** | GATE ANTECIPADO? |
| 4 | **BC** | PULMÃO / DEPOT SAÍDA |
| 5 | **BB** | SEMANA REAL |
| 6 | **BA** | OBSERVAÇÕES FRIGO |
| 7 | **AX** | NO REDEX? *(a AY DATA ENTRADA REDEX permanece)* |
| 8 | **AQ** | CARRETA |
| 9 | **AN** | UNIDADE (TRANSPORTE) *(a N UNIDADE DETALHADA permanece)* |
| 10 | **AM** | TIPO DE CONTRATO |

✅ **Confira antes de seguir:** a planilha deve terminar na coluna **BB — JUSTIFICATIVA SAVING** (54 colunas).

---

## ETAPA 2 — Renomear 2 cabeçalhos (linha 4)

| Célula | De | Para |
|---|---|---|
| **AN4** | CAVALO | `CAVALO/CARRETA` |
| **AP4** | POSIÇÃO ATUAL | `POSIÇÃO` |

---

## ETAPA 3 — Inserir as 4 colunas novas

1. Clique com o botão direito no cabeçalho da coluna **AW** (LOCAL DE DEPÓSITO) → *Inserir 4 colunas à esquerda*.
2. Preencha os cabeçalhos na **linha 4**:

| Célula | Cabeçalho |
|---|---|
| **AW4** | `AUT. EMBARQUE` |
| **AX4** | `Nº FORMULÁRIO` |
| **AY4** | `Nº CTE VAZIO` |
| **AZ4** | `VALOR CTE VAZIO` |

3. Copie a formatação: selecione **AV4** (DOCUMENTAÇÃO) → *Copiar* → selecione **AW4:AZ4** → *Colar especial → Somente formatação*.
4. **Faixa de responsabilidade (linha 3):** desfaça e refaça as mesclagens para que fiquem:
   - `AH3:AZ3` → **🟢 FRIGO preenche — Container / Transporte / Gate / Docs**
   - `BA3:BF3` → **🔵 OPERADOR preenche — Controle / Status / Saving**

✅ **Layout final: 58 colunas (A → BF).**

---

## ETAPA 4 — Corrigir a aba GERAL

Ao apagar as colunas da PROGRAMAÇÃO, a GERAL vai exibir `#REF!` exatamente nas colunas que perderam a origem. **Apague essas 8 colunas** (todas ficarão com erro, é o próprio marcador):

`TIPO DE CONTRATO` · `SEMANA REAL` · `UNIDADE` · `CARRETA` · `GATE ANTECIPADO?` · `NO REDEX?` · `OBSERVAÇÕES FRIGO` · `STATUS GERAL`

Depois:

1. **Renomeie** o cabeçalho `CAVALO` → `CAVALO/CARRETA` e `POSIÇÃO ATUAL` → `POSIÇÃO`.
2. **Corrija o bug pré-existente** da coluna `COLETADO` (já estava `#REF!` antes das nossas mudanças). Cole na 1ª linha de dados e arraste para baixo:

```
=IFERROR(IF(XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!O:O;"")="";"";XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!O:O;""));"")
```

3. **Adicione as 4 colunas novas no fim da GERAL** (depois de `JUST. SAVING`), com estes cabeçalhos e fórmulas (1ª linha de dados = linha 4; arraste até a linha 503):

| Cabeçalho | Fórmula |
|---|---|
| `AUT. EMBARQUE` | `=IFERROR(IF(XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AW:AW;"")="";"";XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AW:AW;""));"")` |
| `Nº FORMULÁRIO` | `=IFERROR(IF(XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AX:AX;"")="";"";XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AX:AX;""));"")` |
| `Nº CTE VAZIO` | `=IFERROR(IF(XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AY:AY;"")="";"";XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AY:AY;""));"")` |
| `VALOR CTE VAZIO` | `=IFERROR(IF(XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AZ:AZ;"")="";"";XLOOKUP($C4;PROGRAMAÇÃO!B:B;PROGRAMAÇÃO!AZ:AZ;""));"")` |

> Se sua planilha usa vírgula como separador de argumentos, troque `;` por `,`.

---

## ETAPA 5 — Colar os dados migrados (carga única)

Abra o arquivo **`MIGRACAO_POSICOES_para_PROGRAMACAO.xlsx`**. Ele tem 199 linhas, **na mesma ordem** das linhas 5 a 203 da PROGRAMAÇÃO. As colunas laranja são as que serão coladas.

Cole sempre com **Ctrl + Shift + V** (colar somente valores):

| Copie do arquivo | Cole na PROGRAMAÇÃO |
|---|---|
| Coluna `CAVALO/CARRETA` (linhas 2–200) | **AN5** |
| Coluna `POSIÇÃO` (linhas 2–200) | **AP5** |
| Bloco `AUT. EMBARQUE` → `VALOR CTE VAZIO` (4 colunas, linhas 2–200) | **AW5** |

As colunas `LINHA`, `PROFORMA`, `REF. CLIENTE`, `CONTAINER` e `ORIGEM DO MATCH` servem **só para conferência** — não devem ser coladas.

---

## ETAPA 6 — Validação (fazer antes de seguir para o item 2)

- [ ] PROGRAMAÇÃO termina em **BF — JUSTIFICATIVA SAVING**
- [ ] Nenhum `#REF!` na GERAL
- [ ] O DASHBOARD continua exibindo os mesmos números de antes
- [ ] As cores por STATUS continuam pintando a linha inteira até a coluna BF
- [ ] Linha 5 (proforma 1051/2026): `CAVALO/CARRETA` = **QTH8B47/FGV4B87**, `POSIÇÃO` = **CARREGADO**, `AUT. EMBARQUE` = **OK**
- [ ] As listas suspensas de STATUS, TIPO DE CARGA e TIPO DE RETIRADA continuam funcionando

---

## Resultado do cruzamento (transparência)

| Situação | Linhas |
|---|---|
| Match por **CONTAINER** | 127 |
| Match por **CONTRATO** (Ref. Cliente) | 48 |
| **Sem match** → veio vazio | 24 |

Campos preenchidos: CAVALO/CARRETA 128 · POSIÇÃO 129 · AUT. EMBARQUE 105 · Nº FORMULÁRIO 14 · Nº CTE VAZIO 11 · VALOR CTE VAZIO 21.
Os números baixos dos 3 últimos **não são erro do cruzamento**: esses campos já estão majoritariamente vazios na própria POSIÇÕES 2026.

---

## O que **não** foi tocado
DASHBOARD (nenhuma alteração) · fórmulas de STATUS, SEMANA, COLETADO, DATA IDEAL RETIRADA, DATA LIMITE DEPÓSITO, QTDE PODE CARREGAR · formatação condicional · validações de dados · coluna ALTERADO (virá no item 5).
