# Engenharia de Prompt

Este projeto foi construído em pareamento com um assistente de IA (Claude). O que fez a
diferença não foram prompts "mágicos", mas um **método** consistente: um prompt de papel
fixo, um padrão de execução verificação-primeiro, e templates de tarefa reutilizáveis.
Este documento registra esses padrões para reprodução.

---

## 1. Prompt de papel (o "contrato" fixo do projeto)

Colado no início de cada sessão/projeto para fixar comportamento. É a peça mais
importante: mais valiosa que qualquer prompt de tarefa individual, porque governa
**como** a IA age.

```
Você atuará como Analista de Negócios, Analista de Processos e Desenvolvedor
especialista em Excel Online / Office Scripts.

Objetivo: evoluir uma planilha operacional JÁ EM PRODUÇÃO. O objetivo não é
reconstruir a solução, mas implementar melhorias preservando a experiência atual.

Forma de trabalho — antes de implementar qualquer alteração:
- analise o impacto da mudança;
- identifique dependências;
- explique o que pretende alterar;
- aguarde minha aprovação.
Nunca implemente alterações sem confirmação.

Princípios (sempre priorize): menor impacto possível; preservação do layout;
preservação da experiência atual; simplicidade; facilidade de manutenção;
estabilidade. Havendo duas soluções, escolha a menos invasiva.

Melhorias fora do escopo: não implemente — apresente como sugestão, explique
benefícios e impactos, aguarde aprovação.

Comunicação: em dúvida sobre regra de negócio, pergunte. Nunca suponha. Se uma
implementação puder afetar outra funcionalidade, informe antes.

Implementação: uma funcionalidade por vez. Após cada uma, valide o que foi
alterado antes da próxima.

Arquivos do projeto são a fonte oficial de referência — consulte antes de
responder; não recrie estruturas que já existam.

Mantenha uma memória do projeto: requisitos aprovados, alterações, pendências,
decisões — e consulte antes de responder.
```

**Por que funciona:** transforma um assistente "faz-tudo-agora" num parceiro cauteloso.
Os ganhos concretos foram: overrides manuais preservados, zero reescrita destrutiva não
autorizada, e uma trilha de decisões auditável.

---

## 2. Template de tarefa — gerar um Office Script

Usado toda vez que uma mudança precisava de script. Os itens em **negrito** são o que
evita retrabalho (aprendidos na marra — ver seção 4).

```
Gere um Office Script (TypeScript, Excel Online) que faça: <objetivo>.

Requisitos obrigatórios:
- Comece com `const MODO_VERIFICACAO = true;` — nesse modo NÃO altere nada,
  apenas liste no log o que SERIA alterado (linha, valor antes/depois).
- **Localize colunas por nome de cabeçalho**, nunca por letra fixa.
- **Sem acentos e sem emoji no código** — monte strings acentuadas com
  `String.fromCharCode()` se precisar.
- **Nada de `console.log` dentro de laço** — acumule num array e imprima uma vez.
- Aborte com segurança se encontrar mais de N mudanças que o esperado.
- Não toque em layout, formatação ou outras fórmulas fora do escopo.

Antes do código: explique o que ele vai alterar e as dependências. Depois: me diga
exatamente o que confere no log de verificação antes de eu trocar para `false`.
```

---

## 3. O loop de execução (padrão verificação-primeiro)

Não é um prompt único, é um protocolo de duas rodadas — o coração da segurança do
projeto:

1. **Rodada 1 — conferência.** Script roda com `MODO_VERIFICACAO = true`. Não altera
   nada; imprime o que faria. O operador lê o log e confere contra um gabarito esperado.
2. **Rodada 2 — execução.** Só depois de o log bater, troca para `false` e roda **uma
   vez**.
3. **Rollback trivial.** As linhas `ANTES:` do log de verificação são a receita de
   reversão manual, caso preciso.

Prompt típico de conferência entre as rodadas:

```
Rodei em MODO_VERIFICACAO. Segue o log: <colar log>. Confere se bate com o esperado
e me diz se pode executar. Se algo estiver fora, NÃO execute — ajuste o script.
```

---

## 4. Anti-padrões (o que instruir a IA a evitar)

Aprendizados que viraram regra de prompt porque custaram tempo:

- **`setNumberFormatLocal` / códigos localizados** → sempre a API neutra
  (`setNumberFormat("dd/mm/yyyy")`). Peça isso explicitamente.
- **`format.setLocked()` e `preset.getRule().setCriterion()`** → não existem; peça
  `getProtection().setLocked()` e `setRule({ criterion: ... })`.
- **Colar código com acento/emoji** corrompe o script no editor → montar por
  `String.fromCharCode()`.
- **Hardcode de letra de coluna** quebra na próxima mudança estrutural → coluna por
  cabeçalho, sempre.
- **`TEXT(...,"yyyy")`** sai literal em sessões PT → usar `YEAR()` em planilha
  multi-idioma.
- **Fórmula reescrita apaga override manual** → o script deve preservar/relatar
  overrides antes de sobrescrever.

---

## 5. Prompts de documentação/portfólio

Para gerar o case study e textos de portfólio sem vazar dados de cliente:

```
Escreva um resumo do projeto para portfólio externo. Use linguagem neutra
("operador logístico", "cliente exportador"). Generalize as figuras operacionais —
nada de números exatos de cliente. Foque no problema resolvido e no impacto, não no
detalhe técnico. Formato: <case study markdown | trecho de LinkedIn | post>.
```
