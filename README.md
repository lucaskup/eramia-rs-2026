# Programa ERAMIA-RS 2026

Página estática com o programa das sessões técnicas e de pôsteres do ERAMIA-RS 2026, com busca por autor, ID e título, filtros por dia, tipo de sessão e trilha, e tema claro/escuro.

## Arquivos

- `index.html`: a página. Não depende de servidor nem de rede; abre direto do disco.
- `programa.js`: dados do programa, gerados a partir da planilha. Não edite à mão.
- `data/programa.xlsx`: cópia da planilha oficial usada para gerar os dados.
- `scripts/build_data.py`: converte a planilha em `programa.js`.

## Atualizar o programa

1. Baixe a planilha atualizada como `.xlsx` e salve em `data/programa.xlsx`.
2. Rode:

```sh
uv run scripts/build_data.py
```

O script lê as abas `Sessões Técnicas` e `Sessões de Pôsteres`, remove o asterisco dos títulos e marca esses trabalhos como candidatos a best paper, limpa restos de LaTeX e normaliza nomes escritos todos em maiúsculas. Correções pontuais de digitação ficam em `TITLE_FIXES` e `AUTHOR_FIXES` no início do script. O local de cada sessão vem de `ROOM_POSTER` e `ROOMS_ORAL`, também no início do script; se a planilha ganhar uma coluna `Local`, ela tem precedência.

## Publicar

Cada push em `main` dispara o workflow `.github/workflows/pages.yml`, que regenera `programa.js` a partir de `data/programa.xlsx` e publica `index.html` no GitHub Pages. Para atualizar o programa basta trocar a planilha em `data/` e fazer push. A página fica em https://lucaskup.github.io/eramia-rs-2026/.

Fora do GitHub, qualquer servidor estático serve: basta copiar `index.html` e `programa.js`.

## Links compartilháveis

A busca e os filtros ficam na URL, por exemplo `index.html#q=kupssinsku&dia=2026-11-09&tipo=oral&trilha=Graduação`.
