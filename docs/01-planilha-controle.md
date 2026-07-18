# Case: Da Planilha Fragmentada ao Controle Compartilhado — Logística de Exportação

## 1. Contexto e o problema

Numa operação de exportação de carga congelada/resfriada, a programação logística
(contêineres, plantas de origem, bookings, prazos de depósito no porto) dependia de
informações que nasciam espalhadas em controles separados de diferentes áreas —
tanto do lado do operador logístico quanto do lado do cliente exportador.

O fluxo, antes da solução:

1. Cada departamento mantinha seu próprio controle isolado.
2. O cliente coletava manualmente esses dados dispersos para montar um arquivo
   consolidado interno.
3. O operador logístico, por sua vez, extraía dados desse arquivo e montava seu
   próprio controle paralelo em Excel.
4. Esse controle era enviado por e-mail para o cliente.
5. O cliente então realimentava manualmente seu arquivo interno e redistribuía a
   versão atualizada para as demais áreas.

Esse ciclo gerava informação sempre defasada (nunca "ao vivo"), retrabalho manual em
múltiplas pontas, e — o gatilho real do projeto — **falta de visibilidade da
liderança do cliente sobre o andamento da programação**, o que gerou pressão direta
para que existisse um controle único e atualizado em tempo real.

## 2. A solução

Criação de uma **planilha única compartilhada** entre as duas empresas, substituindo
os controles paralelos por uma fonte única de verdade, com dados populados
diretamente pelas áreas responsáveis à medida que ficam disponíveis (sem
intermediação por e-mail).

- Piloto/testes: última semana de junho de 2026.
- Entrada em produção oficial: segunda semana de julho de 2026, já usando dados
  reais da programação corrente.
- Adoção: mais de dez pessoas operando ao vivo, divididas em times técnicos e
  comerciais dos dois lados (equipamentos, operação e comercial no operador
  logístico; transporte, comercial, PCP e gestão no cliente).

Essa mudança de "arquivo replicado e reenviado por e-mail" para "fonte única
acessada ao vivo por todos" é o núcleo do ganho — antes de qualquer melhoria
posterior de automação ou dashboard.

## 3. Estrutura de dados

- Aba principal de programação, com uma linha por contêiner/operação e colunas
  organizadas por responsabilidade de preenchimento (uma convenção visual indica
  quais colunas cada empresa preenche).
- Campos centrais: status da operação, semana (calculada automaticamente a partir
  da data), referência do cliente, planta/unidade de origem, país de destino, tipo
  de solicitação (programação regular vs. extra), status de booking, e um alerta de
  prazo de depósito no porto (detention).
- Volume: centenas de linhas ativas por período, cobrindo múltiplas semanas
  simultaneamente — volume que cresce continuamente conforme novas operações
  entram.
- Origem dos dados: preenchimento manual colaborativo (cada área alimenta seu
  bloco de colunas conforme a informação fica disponível), complementado por uma
  migração pontual de um arquivo legado do cliente que estava sendo descontinuado
  (a maior parte dos registros históricos foi migrada automaticamente, casando por
  identificador do contêiner e, como alternativa, pelo número do contrato).

## 4. Principais lógicas e construção

- **Status calculado com override manual**: a coluna de status combina uma fórmula
  automática com a possibilidade de sobrescrita manual (ex.: marcação de
  "cancelado"), preservando flexibilidade operacional sem perder o cálculo padrão.
- **Alerta de prazo (detention)**: contagem regressiva calculada como a diferença
  entre a data-limite de depósito e a data atual, que se congela quando o item já
  foi depositado e fica em branco quando cancelado. Formatação condicional em três
  cores (vermelho/amarelo/verde) sinaliza o nível de risco visualmente.
- **Geração de comunicação (e-mail de estufagem)**: uma aba dedicada replica o
  layout de e-mail que antes era montado manualmente, buscando os dados
  automaticamente na aba de programação através de fórmulas de busca (lookup) por
  referência do cliente — eliminando a montagem manual campo a campo.
- **Dashboard de acompanhamento**: painel com indicadores-resumo, série temporal
  de semanas recentes (janela móvel) e quebras por planta e por país, com filtro
  interativo que atualiza os indicadores mantendo o gráfico semanal como contexto
  fixo não filtrado.
- **Automação via scripts**: como o ambiente é Excel Online (não Google Sheets),
  a automação foi construída com Office Scripts (TypeScript) em vez de Apps
  Script — o que exigiu adaptar abordagens comuns em Google Sheets para as APIs e
  limitações específicas do Excel Online (ex.: um seletor/slicer nativo só se
  conecta a uma tabela dinâmica por vez, contornado com uma tabela auxiliar oculta
  e fórmulas reativas).
- **Disciplina de entrega**: cada script de automação é entregue com um modo de
  verificação (dry-run) ativado por padrão, permitindo validar o resultado antes
  de aplicar qualquer alteração na planilha em produção.

## 5. Aprendizados

- O maior ganho de um projeto assim nem sempre é a automação mais sofisticada —
  foi a eliminação do ciclo de replicação manual entre planilhas separadas que
  resolveu a dor original de visibilidade.
- Trabalhar sobre uma planilha em produção, usada por múltiplas áreas de duas
  empresas simultaneamente, exige processo: testar em cópia, validar com o
  responsável antes de aplicar, e preservar a experiência já conhecida pelos
  usuários em vez de redesenhar tudo de uma vez.
- Identificar corretamente a plataforma real (Excel Online vs. Google Sheets)
  muda a abordagem técnica inteira — soluções que parecem óbvias em um ambiente
  não existem ou funcionam diferente no outro.
