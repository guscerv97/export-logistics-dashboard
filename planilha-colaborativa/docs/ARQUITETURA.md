# Arquitetura & Regras de Negócio

## Abas

| Aba | Papel |
|---|---|
| `PROGRAMAÇÃO` | Base operacional. 1 linha = 1 contêiner/operação. Cabeçalho na linha 4, dados da linha 5 em diante. |
| `ESTUFAGEM` | Monta o e-mail de estufagem por `XLOOKUP` na REF. CLIENTE. |
| `DASHBOARD` | KPIs + série semanal (janela móvel) + quebras por planta/país. |
| `DASHBOARD COLETA` | Visão de coleta. |
| `ARQUIVO` | Histórico de operações concluídas. |

> Evolução: a aba `GERAL` (espelho posicional) existia nas primeiras versões e foi
> removida na refatoração estrutural; o histórico passou para `ARQUIVO`.

## Convenção de preenchimento

O topo da `PROGRAMAÇÃO` tem faixas coloridas indicando quem preenche cada bloco de
colunas (operador logístico × cliente). É uma convenção **visual** — a lógica dos
scripts não depende dela, mas os usuários sim.

## Colunas de fórmula (automáticas)

Algumas colunas são calculadas e não devem ser editadas à mão. A coluna **STATUS** é
híbrida: fórmula + override manual por dropdown ("Cancelado" / "Depositado"). Por isso
ela fica **fora** do escopo de proteção por padrão — travá-la mataria o override.

## Regras de negócio confirmadas

- **1 contrato = 1 REF. CLIENTE distinta.**
- **STATUS** dispara pela **data de coleta** (não pela presença de contêiner).
- **Booking:** número = "Confirmado"; texto "AG BOOKING" = "Aguardando"; vazio = "Sem booking".
- **Risco de detention:** margem baixa na coluna de alerta, sem depósito.
- **Filtro de dashboard:** sempre por **SEMANA PROGRAMADA**, nunca por data real /
  `ISOWEEKNUM`.
- **Programação × Extra:** vem da coluna TIPO DE SOLICITAÇÃO.

## Qualidade de dados — pontos de atenção

- `MÉXICO` × `MEXICO` (grafias diferentes no campo país) — dashboards unificam na
  contagem; recomendável padronizar na origem via lista suspensa. Normalização estrutural
  na `DASHBOARD COLETA` ainda pendente.
- "AG BOOKING" digitado como texto livre — mesma recomendação de lista suspensa.

## Pendências estruturais conhecidas

- Texto do guia rápido do `DASHBOARD` ainda referencia a aba `GERAL` (removida).
- Normalização `MÉXICO/MEXICO` na `DASHBOARD COLETA`.
- Alinhar com o cliente o preenchimento das colunas de saída da planta e
  documentação/CSI.
- Layout de colunas da aba `ARQUIVO` vs. `PROGRAMAÇÃO` (reordenar antes de arquivar em
  massa).
