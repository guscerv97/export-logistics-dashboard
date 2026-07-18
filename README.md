# Programação de Exportação — De Planilha Fragmentada a Dashboard Executivo

Projeto real de logística internacional: eliminação de um ciclo manual de
controles paralelos entre duas empresas, substituído por uma planilha única
compartilhada, e posterior evolução para um dashboard executivo construído com
uma plataforma no-code guiada por IA.

**Status:** em andamento — planilha em produção desde julho/2026, dashboard em
evolução contínua. Veja o [CHANGELOG](./CHANGELOG.md) para atualizações.

## O problema

Duas empresas parceiras (operador logístico e cliente exportador) controlavam
a programação de containers de exportação em arquivos separados. A informação
era replicada manualmente por e-mail entre as partes, gerando dados sempre
defasados, retrabalho e falta de visibilidade da liderança sobre o andamento
real da operação.

## O que foi construído

1. **Planilha única compartilhada** (Excel Online) — fonte única de verdade
   entre as duas empresas, com automações via Office Scripts, alerta visual de
   prazo de depósito no porto (detention) e geração automática de comunicação
   de estufagem.
2. **Dashboard executivo** (base44, no-code) — construído a partir da mesma
   planilha, traduzindo dados operacionais densos em indicadores de gestão,
   com um processo de desenvolvimento em três camadas: protótipo funcional
   validado → direção visual via IA generativa → implementação no-code.

## Decisões técnicas que valem destacar

- Leitura de colunas por **nome, não posição** — a planilha de origem muda de
  estrutura ao longo do tempo; ler por índice fixo quebra silenciosamente.
- **Chave composta** validada estatisticamente contra a base real para
  comparar diferentes versões dos dados com segurança.
- Validações de dado tratadas como **alertas, não sobrescritas automáticas** —
  a fonte de verdade manual é preservada por padrão.
- Todo indicador do dashboard tinha um **número esperado documentado antes da
  implementação**, tornando o QA uma checagem objetiva.

## Documentação completa

- 📄 [Case: a planilha de controle compartilhado](./docs/01-planilha-controle.md)
- 📄 [Case: o dashboard construído com IA (base44)](./docs/02-dashboard-ia.md)

## Stack

Excel Online · Office Scripts (TypeScript) · base44 (no-code) · IA generativa
(análise de dados, prompt engineering, geração de imagem)

---

*Nomes de empresas e valores absolutos de operação foram omitidos por
confidencialidade. A arquitetura e o processo descritos refletem fielmente o
projeto real.*
