# aulas-logistica-senai

Skill para o Claude que gera os slides das aulas (PPTX) no estilo SENAI a partir do plano de aula. Uma apresentação por Unidade Curricular; cada aula de 4 horas tem abertura, dois blocos de conteúdo (10 a 12 slides) e a última hora para atividade de fixação, questões e síntese.

## O que você precisa

- Claude com skills habilitadas (claude.ai, app ou Claude Code).
- O plano de aula da UC. Funciona melhor com o plano gerado pelo **GeraMaterial** (Word ou PDF), que já separa os encontros em "Dia 01", "Dia 02"... com Conhecimentos, Objetivo, Estratégia MSEP, Atividade e Avaliação. Também aceita outros planos em Word, PDF, Excel ou texto.

## Como instalar

**No claude.ai / app:** compacte a pasta `aulas-logistica-senai` em um .zip e envie em Configurações > Capacidades > Skills.

**No Claude Code:** copie a pasta para `.claude/skills/` do seu projeto (ou `~/.claude/skills/` para usar em qualquer projeto).

## Como usar

1. Em uma conversa nova, peça: "Use a skill aulas-logistica-senai para gerar os slides deste plano de aula".
2. Anexe o plano de aula.
3. O Claude lê os encontros, planeja os slides de cada aula e devolve um .pptx por UC. Se o plano não tiver o nome do instrutor, ele pergunta.

## O que sai em cada aula

- Capa da aula com objetivos e roteiro (sem horários por item).
- Bloco 1 e bloco 2 de conteúdo: conceitos, exemplos, comparações, processos e perguntas para a turma.
- Última hora: atividade de fixação (em grupo ou individual, com tabela quando houver cálculo), questões e síntese.

## Arquivos

- `SKILL.md`: instruções que o Claude segue.
- `gen_aulas.js`: gerador dos slides (Node.js + pptxgenjs).
- `exemplo-aula.json`: exemplo de uma aula com todos os tipos de slide.
