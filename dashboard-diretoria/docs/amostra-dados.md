# Amostra de dados (fictícia)

Esta tabela ilustra o formato real de uma linha da aba `PROGRAMAÇÃO`, com valores
**totalmente fictícios** — nenhum dado de cliente, motorista, container ou booking real
foi usado. Serve apenas para dar contexto visual ao mapeamento de colunas descrito no
README deste módulo.

| STATUS | DESTINO | PAÍS DE DESTINO | REF CLIENTE | SEMANA PROGRAMADA | DATA DE ESTUFAGEM | TIPO DE SOLICITAÇÃO | BOOKING | NAVIO | ARMADOR | DATA LIMITE DEPÓSITO | CONTAINER |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Container retirado | ROTTERDAM | HOLANDA | PED-00123 | 30 | 2026-07-20 | Programacao | BKG0001234 | NAVIO EXEMPLO / 001A | ARMADOR FICTÍCIO | 2026-08-02 | TESU1234567 |
| Nova programação | SINGAPORE | SINGAPURA | PED-00124 | 31 | 2026-07-27 | Extra | AG BOOKING | (aguardando) | (aguardando) | 2026-08-09 | — |
| Container retirado | HAMBURGO | ALEMANHA | PED-00125 | 30 | 2026-07-21 | Programacao | BKG0005678 | NAVIO EXEMPLO 2 / 002B | ARMADOR FICTÍCIO 2 | 2026-08-01 | DEMO7654321 |
| Cancelado | HAMBURGO | ALEMANHA | PED-00126 | 30 | — | Programacao | — | — | — | — | — |

**Convenções ilustradas:**
- `BOOKING` vazio ou como texto ("AG BOOKING") representa os estados *Sem booking* /
  *Aguardando*; um valor alfanumérico representa *Confirmado*.
- `CONTAINER` só é preenchido depois da retirada física.
- Linha `Cancelado` mostra como a maior parte dos campos operacionais fica em branco,
  sem entrar nas contagens de ativos do dashboard.
