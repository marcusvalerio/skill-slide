---
name: aulas-logistica-senai
description: Gera os slides de aula (PPTX) no estilo SENAI a partir do plano de aula enviado pelo usuário: uma apresentação por UC, seguindo as aulas do plano.
---

# Aulas em PPTX no estilo SENAI a partir do plano de aula

Fluxo de uso: a pessoa carrega esta skill e depois envia o **plano de aula** da UC (Word, PDF, Excel ou texto). A skill lê o plano e gera um .pptx por UC, com uma sequência de slides para cada aula do plano, no estilo visual SENAI (capa e aberturas em azul-marinho, conteúdo em fundo off-white, títulos em Georgia).

## 1. Ler o plano de aula

Extraia o texto do arquivo (python-docx, pdftotext, openpyxl ou leitura direta). Anote:

- curso, UC (número e nome), carga horária e instrutor
- capacidades e conhecimentos da UC
- para cada aula do plano: número, data, conhecimentos trabalhados, estratégias/atividades, recursos e avaliação

O plano de aula manda: cada aula do plano vira uma aula nos slides, na mesma ordem e com os mesmos conhecimentos. Não invente aulas nem conteúdos fora do plano; use as estratégias e atividades do plano como base da atividade de fixação. Se o plano não trouxer o instrutor, pergunte só o nome que vai na capa. Se o plano tiver mais de uma UC, gere um arquivo por UC.

## 2. Planejar as aulas

- Siga a divisão de aulas do plano. Escreva os objetivos de cada aula a partir das capacidades e conhecimentos que o plano associa a ela.
- Toda aula tem 4 horas, nesta estrutura (o roteiro da abertura lista as etapas, sem horários):
  1. Abertura e revisão
  2. Bloco 1 de conteúdo (5 ou 6 slides)
  3. Intervalo
  4. Bloco 2 de conteúdo (5 ou 6 slides)
  5. Última hora: atividade de fixação (em grupo ou individual) + questões de fixação + síntese
- Cada aula tem de 10 a 12 slides de conteúdo, variando os tipos. Não repita o mesmo tipo em sequência.
- Use "pergunta" (pergunta para a turma) em pelo menos 3 slides por aula e "atencao" quando houver alerta de segurança, norma ou erro comum.
- A atividade de fixação de cada aula é de fixação do conteúdo do dia, coerente com a estratégia prevista no plano: em grupo ("modo": "grupo") ou individual ("modo": "individual"). Quando houver cálculo, classificação ou preenchimento, inclua uma "tabela" para os alunos completarem.
- A última aula de cada UC fecha com uma atividade que revisa a UC inteira.
- Se uma UC for de integração/projeto, use um caso único que avança uma etapa por aula.

## 3. Regras de texto

- Português do Brasil, linguagem simples e didática, para turma de qualificação/técnico.
- Nunca use travessão longo (—). Use dois-pontos, vírgula ou ponto médio (·).
- Sem citações de autores: o material é para dar aula.
- Todo slide tem "notas" para o professor: como conduzir, exemplo para usar, resposta esperada.
- Respeite os limites abaixo; se o texto não couber, divida em dois slides em vez de diminuir a fonte.

| Tipo | Campos | Limites |
| --- | --- | --- |
| conceito | definicao, pontos [[título, texto]] | definição até 230 caracteres; 3 ou 4 pontos; texto até 90 |
| cards | itens [[título, texto]] | 3 ou 4 itens; texto até 110 |
| grade | itens [[título, texto]] | 3 a 8 itens; texto até 60 |
| processo | passos [[nome, descrição]] | 4 a 6 passos; nome até 22; descrição até 45 |
| tabela | cab [...], linhas [[...]], colW opcional (soma 12,23) | até 7 linhas; células curtas |
| comparar | a [título, [itens]], b [título, [itens]] | até 5 itens por lado, 60 caracteres cada |
| destaque | numero, rotulo, texto | número até 10 caracteres; texto até 260 |
| pratica | caso, pergunta | caso até 260; pergunta até 120 |
| exemplo | cab, linhas, passos [...], resultado | até 8 linhas e 4 colunas; até 5 passos; resultado até 120 |
| visual | blocos [[valor, rótulo]], texto | 2 a 6 blocos com valor de até 4 caracteres; texto até 200 |

Campos opcionais em qualquer slide de conteúdo: "bloco" (1 ou 2), "pergunta" (até 120 caracteres), "atencao" (até 110 caracteres).

## 4. Formato dos dados (um JSON por UC)

```json
{
  "uc": 3, "nome": "Logística de Armazenagem", "horas": 60,
  "curso": "Assistente de Operações em Logística", "professor": "Nome do Professor",
  "objetivo": "Uma frase sobre o que a UC desenvolve.",
  "aulas": [
    {
      "titulo": "Acuracidade, curva ABC e auditoria",
      "blocos": ["Acuracidade e inventário", "Curva ABC e auditoria"],
      "objetivos": ["Calcular ...", "Classificar ...", "Explicar ..."],
      "slides": [
        {"tipo": "destaque", "bloco": 1, "titulo": "...", "numero": "95%", "rotulo": "...", "texto": "...", "notas": "..."},
        {"tipo": "tabela", "bloco": 1, "titulo": "...", "cab": ["Classe", "Itens"], "linhas": [["A", "20%"]], "atencao": "...", "notas": "..."},
        {"tipo": "visual", "bloco": 2, "titulo": "...", "blocos": [["03", "Rua"], ["12", "Prédio"]], "texto": "...", "notas": "..."},
        {"tipo": "exemplo", "bloco": 2, "titulo": "...", "cab": ["Item", "Valor"], "linhas": [["P01", "30.000"]], "passos": ["..."], "resultado": "...", "pergunta": "...", "notas": "..."}
      ],
      "atividade": {"titulo": "Classifique o estoque", "modo": "grupo", "formato": "Em grupos de 4",
        "passos": ["...", "..."], "tabela": {"cab": ["Item", "Preço", "Classe"], "linhas": [["P01", "50"], ["", ""]]},
        "entrega": "...", "notas": "..."},
      "quiz": ["Pergunta 1?", "Pergunta 2?", "Pergunta 3?"],
      "respostas": "Respostas esperadas das 3 perguntas."
    }
  ]
}
```

Veja `exemplo-aula.json` nesta pasta para uma aula completa com todos os tipos.

## 5. Gerar e conferir

1. Escreva o JSON de cada UC com Python (`json.dump(..., ensure_ascii=False)`).
2. Localize o tema da skill de PPTX: `export APPLY_THEME=$(find / -name apply_theme.js -path "*pptx*" 2>/dev/null | head -1)`.
3. Rode o gerador desta pasta: `node gen_aulas.js ucN.json "UCN - Nome da UC.pptx"` (usa `pptxgenjs`).
4. Valide com o `validate.py` da skill de PPTX, converta para PDF e olhe os slides mais carregados (tabelas longas, processos de 6 passos, grades de 7 ou 8 itens, exemplos). Corrija o que estourar encurtando ou dividindo o conteúdo.
5. Entregue os arquivos e diga em uma linha quantas aulas e slides cada UC tem.

## O que o gerador monta

Capa (marinho) → Mapa da UC → para cada aula: abertura (marinho: objetivos e roteiro da aula) → abertura do Bloco 1 → slides do Bloco 1 → abertura do Bloco 2 → slides do Bloco 2 → atividade de fixação → questões de fixação → síntese da aula (com a próxima aula). Rodapé com curso, UC e professor em todos os slides.
