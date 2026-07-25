# Controle Colaborativo de Contêineres — Operador Logístico × Cliente Exportador

Ferramenta colaborativa de gestão de programação de contêineres para exportação de
cargas refrigeradas/congeladas, construída sobre **Excel Online (Microsoft 365 /
SharePoint)** e automatizada com **Office Scripts (TypeScript)**.

> **Nota sobre dados:** este repositório usa nomes genéricos ("operador logístico",
> "cliente exportador") e figuras operacionais generalizadas. Nenhum dado real de
> cliente é incluído. É documentação de arquitetura e engenharia, não a base de dados.

---

## 1. O problema

Antes desta ferramenta, cada departamento das duas empresas mantinha seu próprio
controle isolado. O fluxo funcionava assim:

1. O cliente consolidava manualmente os controles das áreas num arquivo próprio de
   acompanhamento.
2. O operador logístico extraía e reformatava esses dados e enviava por e-mail.
3. O cliente redigitava tudo de volta no seu arquivo.

O resultado era **informação perpetuamente desatualizada** e **baixa visibilidade da
programação de contêineres para a liderança** — a dor central, escalada pelos diretores
do cliente.

## 2. A solução

Uma **única planilha viva e compartilhada** substituiu o ciclo fragmentado por uma fonte
de dados única para as duas empresas. Saiu de piloto (fim de junho/2026) para produção
oficial (meados de julho/2026) e hoje é usada por **10+ pessoas em 7 áreas funcionais**
das duas empresas:

- **Operador logístico:** Equipamentos, Operacional, Vendas
- **Cliente exportador:** Transporte, Comercial, PCP, Gestão

Cada área alimenta seu bloco de colunas conforme a informação fica disponível; os
cálculos, alertas e dashboards se atualizam ao vivo.

Meu papel evoluiu ao longo do projeto: comecei como um dos executores da migração e hoje
sou o **administrador da ferramenta**, responsável por correção de bugs, validação de
dados, automações e ajustes estruturais solicitados pelo cliente.

## 3. Arquitetura da planilha

**Abas principais (operacionais)**
- `PROGRAMAÇÃO` — uma linha por contêiner/operação; cabeçalho na linha 4, dados a partir
  da linha 5. Colunas organizadas por responsabilidade de preenchimento, com convenção de
  cores no topo indicando qual empresa preenche cada bloco.
- `ESTUFAGEM` — replica o layout do e-mail de estufagem que antes era montado à mão,
  puxando os dados por `XLOOKUP` a partir da referência do cliente.

**Dashboards**
- `DASHBOARD` — indicadores-resumo, série temporal por semana (janela móvel) e quebras
  por planta e por país, com filtro que atualiza os KPIs mantendo o gráfico semanal como
  contexto fixo.
- `DASHBOARD COLETA` — visão de coleta.

**Histórico**
- `ARQUIVO` — aba de arquivamento de operações concluídas.

**Colunas-chave** (por finalidade, não por letra fixa)

| Campo | Papel |
|---|---|
| STATUS | Situação da operação — fórmula + override manual |
| SEMANA PROGRAMADA | **Referência universal de filtro dos dashboards** |
| REF. CLIENTE | Chave de negócio (1 contrato = 1 referência distinta) |
| PAÍS DESTINO / PLANTA | Quebras dos dashboards |
| TIPO DE SOLICITAÇÃO | Programação regular × Extra |
| BOOKING | Confirmado (número) / Aguardando ("AG BOOKING") / Sem booking (vazio) |
| ALERTA DETENTION | Dias de margem até o prazo de depósito no porto |

Detalhamento completo em [`docs/ARQUITETURA.md`](./docs/ARQUITETURA.md), incluindo abas,
convenções de preenchimento e pendências estruturais conhecidas. Uma amostra fictícia da
estrutura de dados está disponível em
[`docs/amostra-dados.md`](../dashboard-diretoria/docs/amostra-dados.md).

## 4. Lógicas principais

- **STATUS calculado com override manual** — a coluna combina uma fórmula automática
  (dispara pela data de coleta) com a possibilidade de sobrescrita manual via dropdown
  (ex.: "Cancelado", "Depositado"), preservando flexibilidade sem perder o cálculo padrão.
- **Alerta de detention** — contagem regressiva = data-limite de depósito − data atual;
  congela quando o item é depositado e fica em branco quando cancelado. Formatação
  condicional em três faixas (vermelho / amarelo / verde) sinaliza o risco.
- **Geração de e-mail de estufagem** — aba dedicada monta o e-mail automaticamente por
  `XLOOKUP` na referência do cliente, eliminando a montagem manual campo a campo.
- **Dashboard com janela móvel** — série de semanas recentes reconstruída com `FILTER` /
  `SORT` / `TAKE`; filtro por `SEMANA PROGRAMADA` como referência única (nunca a data real
  ou `ISOWEEKNUM`, causa-raiz de erros de contagem em versões anteriores).

## 5. Automação — Office Scripts

O ambiente é Excel Online, então a automação usa **Office Scripts (TypeScript)** via a
aba **Automate**, não Apps Script. Todas as mudanças estruturais, migrações de dados e
reescritas de fórmula em massa passam por scripts com um padrão rígido de segurança (ver
[`docs/PROMPTS.md`](./docs/PROMPTS.md)).

Scripts disponíveis neste repositório (código verbatim, ver pasta [`scripts/`](./scripts/)):

| Finalidade | Script |
|---|---|
| Anatomia-padrão (template de referência) | `_TEMPLATE_office-script-verificacao.ts` |
| Correção de formato de data na ESTUFAGEM (API neutra, independente de idioma) | `corrigir_formato_data_estufagem.ts` |
| Migração de dados + reestruturação de colunas (item 1) | `ITEM1_OfficeScript.ts` |
| Reposicionamento de colunas + fórmulas + proteção (fase 2) | `Fase2_Reposicionar_Colunas.ts` |
| Reconstrução de formatação condicional (fase 3) | `Fase3_Formatacao_Correcoes.ts` |
| Proteger/desproteger colunas de fórmula | `Proteger_Formulas_PROGRAMACAO.ts` / `Desproteger_PROGRAMACAO.ts` |
| Alerta de detention (3 faixas de risco) | `ITEM2_AlertaDetention.ts` |
| Geração automática de e-mail de estufagem | `ITEM3_AbaEstufagem.ts` |
| Dashboard de coleta com filtro por semana | `ITEM4_v2_DashboardColeta_Filtro.ts` |

## 6. Aprendizados técnicos

**Armadilhas confirmadas da API do Office Scripts**
- `setNumberFormatLocal(...)` não funciona → use `setNumberFormat("dd/mm/yyyy")` (string
  em inglês).
- `format.setLocked()` não existe → use `format.getProtection().setLocked()`.
- `preset.getRule().setCriterion()` não existe → use `preset.setRule({ criterion: ... })`.
- `setColumnHidden` não existe → use `setColumnWidth(0)` para ocultar.

**Corrupção de código ao colar** — emojis e strings com acento corrompem o script no
editor. Solução: montar via `String.fromCharCode()` / `String.fromCodePoint()` no fonte.
Evitar `console.log` dentro de laços (avisos de performance).

**`TEXT()` é dependente do idioma da sessão** — `"yyyy"` funciona em sessão EN e sai
literal em PT (e vice-versa com `"aaaa"`). Numa planilha compartilhada entre idiomas,
usar `YEAR()` (independente de idioma) no lugar do código de ano dentro de `TEXT()`.

**Localizar coluna por nome de cabeçalho** (não por letra fixa) é o padrão que sobrevive a
mudanças estruturais.

**`SEMANA PROGRAMADA` é a referência universal de filtro** dos dashboards — não a data de
coleta nem `ISOWEEKNUM` de datas reais. Desalinhamento aqui foi causa-raiz de erros
antigos.

**Verificação-primeiro é inegociável** — rodar com `MODO_VERIFICACAO = true` antes de
executar já pegou overrides manuais perdidos e erros de conversão de data/fórmula.

## 7. Roadmap

- **Portal interno de operações** — dashboard multi-página (retirada, risco de
  detention, risco de overstay, documentação). Avaliação de plataforma em andamento.
- **Refinamentos do case study** — ajustes de tom, anonimização, export `.docx`.

## 8. Estrutura deste módulo

```
planilha-colaborativa/
├── README.md                                  # este arquivo
├── docs/
│   ├── ARQUITETURA.md                         # detalhamento de abas, colunas e regras de negócio
│   ├── PROMPTS.md                              # engenharia de prompt — método e templates
│   └── ITEM1_Roteiro_Colunas.md                # roteiro de migração de dados e reestruturação de colunas
└── scripts/                                    # Office Scripts (TypeScript)
    ├── README.md                               # manifesto dos scripts
    ├── _TEMPLATE_office-script-verificacao.ts
    ├── corrigir_formato_data_estufagem.ts
    ├── ITEM1_OfficeScript.ts                    # migração de dados + reestruturação de colunas
    ├── Fase1_1_Corrigir_K_L.ts                  # correção pós-migração (data/semana)
    ├── Fase2_Reposicionar_Colunas.ts            # reposicionamento de colunas + fórmulas + proteção
    ├── Fase3_Formatacao_Correcoes.ts             # reconstrução da formatação condicional
    ├── Proteger_Formulas_PROGRAMACAO.ts
    ├── Desproteger_PROGRAMACAO.ts
    ├── ITEM2_AlertaDetention.ts                  # alerta de detention (3 faixas de risco)
    ├── ITEM3_AbaEstufagem.ts                     # geração automática de e-mail de estufagem
    └── ITEM4_v2_DashboardColeta_Filtro.ts        # dashboard de coleta com filtro por semana
```

## 9. Stack

Excel Online / SharePoint · Office Scripts (TypeScript) · fórmulas dinâmicas
(`XLOOKUP`, `FILTER`, `SORT`, `TAKE`, `LET`, `ISOWEEKNUM`) · SheetJS + Chart.js
(protótipo HTML, ver módulo [`dashboard-diretoria/`](../dashboard-diretoria/))
