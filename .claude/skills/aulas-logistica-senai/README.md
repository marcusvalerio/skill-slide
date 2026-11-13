# aulas-logistica-senai

Skill para o Claude que gera os slides das aulas (PPTX) no estilo SENAI a partir do plano de aula.

## Como usar

1. Carregue esta skill (pasta `aulas-logistica-senai`) no Claude, por exemplo em `.claude/skills/` do projeto.
2. Envie o plano de aula (Word, PDF, Excel ou texto).
3. O Claude segue as aulas do plano e gera um .pptx por UC.

## Arquivos

- `SKILL.md`: instruções da skill.
- `gen_aulas.js`: gerador dos slides (Node.js + pptxgenjs).
- `exemplo-aula.json`: exemplo de uma aula com todos os tipos de slide.
