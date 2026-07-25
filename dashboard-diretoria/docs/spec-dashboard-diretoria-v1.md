# Especificação Funcional — Dashboard Diretoria (Cliente Exportador × Operador Logístico)
Versão 1 — 17/07/2026 — base: mockup `PAINEL_Dashboard__Operador Logístico_x_Vale_Grande.html`

## Fonte de dados
- Upload manual de arquivo `.xlsx` (botão "Atualizar dados" no site).
- Fonte: aba `PROGRAMAÇÃO` da planilha `Controle Cliente Exportador e Operador Logístico`.
- Leitura direta dos **valores calculados** já salvos no xlsx (não recalcula fórmulas).
- Ao carregar um novo arquivo, a base do site é **substituída** (não acumula histórico de uploads).
- Exibir no topo da página: **data/hora da última atualização** (do momento do upload).

## Mapeamento de colunas confirmado (aba PROGRAMAÇÃO, cabeçalho na linha 4)
| Campo               | Coluna | Observação |
|---|---|---|
| Status              | A  | "retirado" / "nova programação" (pendente) / "Cancelado" |
| Destino             | C  | campo distinto de "País de destino" (D) — não confundir |
| País de destino     | D  | possui grafias MÉXICO/MEXICO — **normalizar na leitura**, não só na exibição |
| Ref. Cliente        | H  | chave de "contrato" (contagem distinta) |
| Semana Programada   | K  | filtro de semana |
| Data de Estufagem   | L  | |
| Unidade Detalhada   | N  | "planta" |
| Tipo de Solicitação | S  | Programação / Extra |
| Booking             | W  | número válido = Confirmado; texto "AG BOOKING" = Aguardando; vazio = Sem booking |
| Alerta Detention    | AF | dias de margem até data-limite de depósito |

## Indicadores (KPIs no topo)
1. **Contratos** — contagem de REF. CLIENTE (H) distintas
2. **Contêineres ativos** — total de linhas ativas (excluindo Cancelado), com detalhe Programação × Extra
3. **Retirados** — contagem por Status (A)
4. **Pendentes** — contagem por Status (A), "nova programação"
5. **Aguardando booking** — W = "AG BOOKING"
6. **Risco detention** — AF ≤ 2 dias

## Cards / gráficos
7. **Retirados × Pendentes** — donut (não mais barras por semana), somando os dados da(s) semana(s) selecionada(s); se nenhuma semana selecionada ("Todas"), soma tudo.
8. **Programação × Extra** — donut (coluna S)
9. **Status de Booking** — donut (Confirmado / Aguardando / Sem booking)
10. **Por país de destino** — barras horizontais (coluna D, MÉXICO/MEXICO normalizados)
11. **Por planta** — barras horizontais (coluna N)
12. **Faixas de risco de detention** — barras (Estourado <0 / Crítico 0–2 / Atenção 3–5 / Ok >5), base coluna AF

## Filtro
13. Seletor de semana (coluna K) — single ou multi-seleção (Ctrl/⌘+clique), afeta todos os cards.

## Interatividade — clique para detalhar
- Cards clicáveis: **Retirados × Pendentes, País, Planta, Booking, Risco Detention**.
- Ao clicar em um segmento/barra, exibe abaixo uma **lista de detalhe** com as linhas correspondentes.
- Colunas fixas da lista de detalhe (mesmas para todos os cards, inclusive Risco Detention):
  **Referência (H) · Data de Estufagem (L) · Destino (C) · Tipo de Solicitação (S) · Booking (W)**
- Comportamento: cada novo clique **substitui** a lista anterior (não acumula).

## Regras de negócio confirmadas
- MÉXICO / MEXICO: tratados como o mesmo país já na leitura/parsing dos dados (não só na exibição).
- Cancelado (Status = A) não entra em nenhuma contagem de ativos.
- Booking: regra de 3 estados confirmada (Confirmado / Aguardando / Sem booking).

## Em aberto / próximos passos
- [ ] Indicadores adicionais que o usuário ainda vai definir.
- [ ] Layout/design final via Nanobanana (esta spec não trata de estilo visual, só de dados e comportamento).
- [ ] Prompt final para base44 — a ser montado após o design estar definido.
