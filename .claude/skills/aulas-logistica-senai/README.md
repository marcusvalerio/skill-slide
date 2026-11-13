# aulas-logistica-senai

Skill para o Claude que gera os slides das aulas (PPTX) no estilo SENAI a partir do PER do curso.

## Como usar

1. Carregue esta skill (pasta `aulas-logistica-senai`) no Claude, por exemplo em `.claude/skills/` do projeto.
2. Envie o PER (PDF, DOCX ou texto).
3. O Claude lista as UCs encontradas, confirma o nome do professor e gera um .pptx por UC, com aulas de 4 horas.

## Arquivos

- `SKILL.md`: instruções da skill.
- `gen_aulas.js`: gerador dos slides (Node.js + pptxgenjs).
- `exemplo-aula.json`: exemplo de uma aula com todos os tipos de slide.
