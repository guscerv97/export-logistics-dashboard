# Programação de Exportação — De Planilha Fragmentada a Dashboard Executivo

Projeto real de logística internacional: eliminação de um ciclo manual de controles
paralelos entre duas empresas parceiras, substituído por uma planilha única
compartilhada, e evolução posterior para um dashboard executivo construído com uma
plataforma no-code guiada por IA.

**Status:** em andamento. Planilha em produção desde julho/2026, dashboard em evolução
contínua. Veja o [CHANGELOG](./CHANGELOG.md) para atualizações.

---

## O problema

Duas empresas parceiras (um operador logístico e um cliente exportador) controlavam a
programação de containers de exportação em arquivos separados. A informação era
replicada manualmente por e-mail entre as partes, gerando dados sempre defasados,
retrabalho e falta de visibilidade da liderança sobre o andamento real da operação.

## Os dois módulos deste projeto

### 1. [Planilha Colaborativa](./planilha-colaborativa/)
Fonte única de verdade entre as duas empresas, em Excel Online, com automações via
Office Scripts (TypeScript), alerta visual de prazo de depósito no porto (detention) e
geração automática de comunicação de estufagem. Hoje em uso por mais de 10 pessoas em 7
áreas funcionais das duas empresas. Atuo como administrador dessa ferramenta: correção
de bugs, validação de dados, automações e ajustes estruturais sob demanda.

### 2. [Dashboard de Diretoria](./dashboard-diretoria/)
Painel executivo read-only, construído em uma plataforma no-code (base44) a partir da
mesma planilha operacional, traduzindo dados densos em indicadores de gestão. Conduzido
de ponta a ponta: levantamento de requisitos, protótipo funcional validado contra dados
reais, direção visual via IA generativa, engenharia de prompt e QA.

## Decisões técnicas que valem destacar

- Leitura de colunas por **nome, não posição** — a planilha de origem muda de estrutura
  ao longo do tempo; ler por índice fixo quebra silenciosamente.
- **Chave composta** validada estatisticamente contra a base real para comparar
  diferentes versões dos dados com segurança.
- Validações de dado tratadas como **alertas, não sobrescritas automáticas** — a fonte
  de verdade manual é preservada por padrão.
- Todo indicador do dashboard tinha um **número esperado documentado antes da
  implementação**, tornando o QA uma checagem objetiva.
- Protocolo de **verificação-primeiro** em toda automação: nenhum script altera dados em
  produção sem antes rodar em modo de simulação e ter o resultado conferido.

## Stack

Excel Online / SharePoint · Office Scripts (TypeScript) · fórmulas dinâmicas (`XLOOKUP`,
`FILTER`, `SORT`) · base44 (no-code) · IA generativa (análise de dados, prompt
engineering, geração de imagem)

---

*Nomes de empresas, clientes finais e valores absolutos de operação foram generalizados
por confidencialidade. A arquitetura e o processo descritos refletem fielmente o projeto
real.*
