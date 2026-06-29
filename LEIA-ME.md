# D&A · Custo Industrial + BI

Sistema web (HTML + CSS + JS puro) com motor de custo portado 1:1 da planilha
Estrutura e dashboard comercial. Roda em GitHub Pages, sem backend.

## Arquitetura de dados (2 arquivos = 2 ciclos)

| Arquivo | Conteúdo | Quem mantém | Frequência |
|---|---|---|---|
| `dados.json` | Tabelas de referência (máquinas, substratos, ribbons, impostos) + catálogo de SKUs | Você | Baixa |
| `vendas.json` | Base de pedidos com financeiro já calculado por linha | Importação do ERP | Diária |

O app só faz `fetch` desses dois arquivos. Nenhum dado fica embutido no código.

## Deploy

1. Suba `index.html`, `styles.css`, `app.js`, `dados.json`, `vendas.json` no repositório do GitHub Pages.
2. Para servir os JSON de um repo público separado (padrão do seu PCP), edite no topo do `app.js`:
   ```js
   const CONFIG = {
     base: 'https://raw.githubusercontent.com/lhtorres1995/pcp-sync/main/',
     dadosUrl: 'dados.json',
     vendasUrl: 'vendas.json',
     editavel: true
   };
   ```

## Modo visualização (membros da empresa)

Publique uma cópia com `editavel: false`. Isso oculta cadastro de SKU, importação e
botões de publicar. Os membros abrem a URL e veem só o dashboard e as tabelas, em leitura.
Como tudo é `fetch` de raw URL, não há token nem escrita exposta.

## Fluxo diário de importação (extrato do ERP)

1. Exporte os pedidos do dia em Excel/CSV (1ª linha = cabeçalhos).
2. No app: **Importar → selecione o arquivo**. As colunas são mapeadas sozinhas
   (reconhece: Emissão, Pedido, Cliente, Família de cliente, Segmento de atividade,
   Vendedor 1, Produto, Descrição, Qtde. Pedido, R$ Un., Grupo do Produto, e, se houver,
   R$ Pedido / Custo / Lucro / MC). Ajuste o que faltar.
3. Escolha **Adicionar** (dedup por pedido+produto) ou **Substituir base inteira**.
4. Pré-visualize e clique **Aplicar**. O `vendas.json` atualizado é baixado.
5. Suba esse `vendas.json` no repositório. O dashboard dos membros reflete na próxima carga.

### De onde vem a margem na importação
- Se o extrato **traz** R$ Pedido / Custo / Lucro: o app usa esses valores direto.
- Se **não traz**: o motor calcula a margem a partir do SKU (precisa estar em `dados.json`).
- Se não há SKU nem custo no extrato: a linha entra só com receita (sinalizada como "sem custo").

> Recomendação: para o dashboard ficar 100% à prova de produto novo, inclua no extrato
> a coluna de valor do pedido (R$ Pedido). Assim a receita nunca depende do cadastro de SKU.

## Publicar mudança de estrutura/custo
Editou um SKU ou precisa atualizar preço de substrato/alíquota? Faça a alteração,
clique **Baixar dados.json** (tela Produtos) e suba o arquivo. Uma mudança de custo
não exige tocar no código.
