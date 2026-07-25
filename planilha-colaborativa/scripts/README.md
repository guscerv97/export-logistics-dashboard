# Scripts — Office Scripts (TypeScript)

Todos os scripts do projeto seguem o padrão **verificação-primeiro**: começam com
`MODO_VERIFICACAO = true`, listam no log o que fariam sem alterar nada, e só executam
depois de o log ser conferido e o modo trocado para `false`. Veja
`_TEMPLATE_office-script-verificacao.ts` para a anatomia-padrão.

## Presentes neste repositório (código verbatim)

| Arquivo | Finalidade |
|---|---|
| `_TEMPLATE_office-script-verificacao.ts` | Esqueleto/anatomia-padrão para novos scripts. |
| `corrigir_formato_data_estufagem.ts` | Ajusta formato de data (C14/C22) na ESTUFAGEM via API neutra. |

## Como rodar

1. Excel Online → aba **Automate** → **New Script** → cole o código → **Save**.
2. Rode com `MODO_VERIFICACAO = true`. Confira o log contra o gabarito esperado.
3. Se bater, troque para `false` e rode **uma vez**.
4. Rollback: as linhas `ANTES:` do log de verificação são a receita de reversão.
