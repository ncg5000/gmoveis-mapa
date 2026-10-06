# gmoveis-mapa — Mapa dos Produtos G. Móveis

Site estático (GitHub Pages) com um card por modelo: foto, ref, medidas, classe e filtros de tela.
Dois modos na mesma página: **trabalho** (status de foto, quadro de combinações, site) e **catálogo**.

## Pastas
- `index.html` — o mapa (gerado; não editar à mão).
- `fotos/` — fotos reduzidas para a web (geradas a partir dos originais no SharePoint).
- `gerador/` — `gerar_mapa_produtos.py` + `config.json`: lê a tabela de produtos e a pasta
  `público Catálogo\Fundo_infinito` no SharePoint e escreve `index.html` + `fotos/`.
- `notas/` — plano e decisões.

## Fontes (SharePoint "Dados - Documentos", sincronizado pelo OneDrive)
- Lista de produtos: `gm\TABELA - PRODUTOS\+Tabela Líquida completa X gmoveis até <data>.xlsx` (aba `Representante`).
- Fotos: `gm\Imagens\público Catálogo\Fundo_infinito\<Categoria>\<ref>-<DESCRIÇÃO>.jpg`.
- Quadros de combinações: `...\Fundo_infinito\Cadeira\Variações_Acabamentos\<ref>\mapa\`.

## Regras
- O site é público: **nunca** publicar preço nem os originais em alta resolução.
- Os originais ficam no SharePoint; aqui só entra o que o gerador escreve.
- Publicar = commit + push na branch `main`; o GitHub Pages atualiza em ~1 min.

## Duas máquinas (micro-01 e luis-casa), mesma configuração
- Clone nas duas em `C:\Projetos\gmoveis-mapa`. Sempre `git pull` antes de mexer e `git push` ao terminar.
- O OneDrive fica em `%USERPROFILE%\OneDrive - G MOVEIS LTDA\Dados - Documentos\` (usuário `User` no micro-01,
  `Usuario` em casa) — o gerador resolve pelo `%USERPROFILE%`, nunca por caminho fixo.
- luis-casa tem Python (Anaconda) → gerador roda local. micro-01 não tem → gerar na nuvem e gravar de volta.
- `gh` precisa estar logado (`gh auth login --web`) em cada máquina para criar/ajustar coisas no GitHub.
