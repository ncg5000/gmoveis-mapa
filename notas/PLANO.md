# PLANO — Mapa dos Produtos (06/10/2026)

## Decisões
- Finalidade: os dois em camadas — modo *trabalho* (status, pendências) e modo *catálogo* (sem status).
- Unidade: um card por modelo. Mesas: um card por linha × tampo (vid/mad), medidas listadas dentro.
- Ocultar/exibir = filtro de tela (categoria, linha, campos, modo). Não persiste.
- Preço: botão, desligado por padrão — SÓ na saída interna (SharePoint). O site público nunca leva preço.
- Local: repositório público `gmoveis-mapa` no GitHub (conta ncg5000) + GitHub Pages. Clone em `C:\Projetos\gmoveis-mapa`.
- Fonte da lista: `+Tabela_gmoveis.xlsx` (mestre; antes `+Tabela Líquida completa X gmoveis até 01-10-26.xlsx`, cópia em `arquivo\Tabela_gmoveis_2026-10-01.xlsx`), aba `Representante` (OOXML strict → converter com LibreOffice).
- WhatsApp: link com prévia (og:tags). Depois, feed CSV do catálogo do WhatsApp Business gerado pelo mesmo gerador.

## Números (tabela 01-10-26 × Fundo_infinito)
- 479 refs; 278 com preço de venda; 195 MontProd (componentes) fora.
- Acabados 233 refs: MESA 166 · CADEIRA 26 · BANQUETA 22 · POLTRONA 13 · APARADOR 6. Peças avulsas 45 (assento/encosto) — grupo desligado por padrão.
- ≈ 82 cards: 26 cadeiras + 22 banquetas + 13 poltronas + ~20 mesas + 1 aparador.
- Fotos: cadeira 26/26 · poltrona 13/13 · banqueta 21/22 · mesa 150/166 refs · aparador 0/6.
- Nomes de foto a corrigir: `194811962` (Madrid 1962), `1920` repetido (Niterói 1930), `2798-…ESTOF+MAD` (parece 2797).
- Quadro de combinações: só Capri 2438 (210/210).

## Fotos (em conversa)
- Principal do card: hoje a que existe (acabamentos variados); padronizar via esteira ao longo do tempo (a decidir).
- Extras: regra de nome `<ref>-<descrição>.jpg` = principal, `<ref>_<ângulo>.jpg` = extras; o que não casa vira pendência.
- Web: miniatura ~600 px no card + ~1600 px no clique; originais ficam no SharePoint.
- Faltam: Aparador Ipanema (6), Mesa Caribe vid 2317/2318/2651/2652, Banqueta Ibiza 2797.

## Próximos passos
1. Criar o repositório no GitHub e ligar o Pages (feito nesta sessão, se o login der certo).
2. Gerador v0: lê tabela + Fundo_infinito → index.html + fotos/ (cards, filtros, 2 modos).
3. Publicar v0 e revisar com o Luis.
