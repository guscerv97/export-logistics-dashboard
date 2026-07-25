# Scripts — Office Scripts (TypeScript)

Todos os scripts do projeto seguem o padrão **verificação-primeiro**: começam com
`MODO_VERIFICACAO = true`, listam no log o que fariam sem alterar nada, e só executam
depois de o log ser conferido e o modo trocado para `false`. Veja
`_TEMPLATE_office-script-verificacao.ts` para a anatomia-padrão.

## Ordem cronológica real do projeto

Os scripts abaixo refletem a evolução real da ferramenta, da migração inicial até a
manutenção corrente como administrador.

| Ordem | Arquivo | O que fez |
|---|---|---|
| 1 | `ITEM1_OfficeScript.ts` (+ roteiro em `../docs/ITEM1_Roteiro_Colunas.md`) | Migração de dados de um arquivo legado + reestruturação de colunas da aba `PROGRAMAÇÃO`/`GERAL`. |
| 2 | `Fase1_1_Corrigir_K_L.ts` | Correção pós-migração: conversão de texto para data (coluna de estufagem) e restauração de fórmula de semana. |
| 3 | `Fase2_Reposicionar_Colunas.ts` | Reposicionamento de colunas, reescrita de fórmulas dependentes e proteção. |
| 4 | `Fase3_Formatacao_Correcoes.ts` | Reconstrução unificada da formatação condicional e do painel de "próximas estufagens". |
| 5 | `Proteger_Formulas_PROGRAMACAO.ts` / `Desproteger_PROGRAMACAO.ts` | Par de scripts para travar/destravar colunas de fórmula automática, preservando o restante editável. |
| 6 | `ITEM2_AlertaDetention.ts` | Reescreve a fórmula e a formatação condicional do alerta de detention (3 faixas de risco). |
| 7 | `ITEM3_AbaEstufagem.ts` | Gera a aba de comunicação de estufagem (e-mail automático) a partir de um contrato digitado, via `XLOOKUP`. |
| 8 | `ITEM4_v2_DashboardColeta_Filtro.ts` | Cria a aba `DASHBOARD COLETA` com cartões, gráficos e filtro interativo por semana. |
| — | `_TEMPLATE_office-script-verificacao.ts` | Esqueleto/anatomia-padrão para novos scripts (referência, não faz parte da linha do tempo). |
| — | `corrigir_formato_data_estufagem.ts` | Correção pontual de formato de data (ver módulo raiz). |

> Uma versão anterior do item 4 (sem filtro por semana) foi descartada em favor da v2,
> incluída aqui.

## Como rodar

1. Excel Online → aba **Automate** → **New Script** → cole o código → **Save**.
2. Rode com `MODO_VERIFICACAO = true`. Confira o log contra o gabarito esperado.
3. Se bater, troque para `false` e rode **uma vez**.
4. Rollback: as linhas `ANTES:` do log de verificação são a receita de reversão.
