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


## Mapeamento confirmado com o seu extrato do ERP (iQuattro)

Testado com o relatório "Pedido de Venda". Reconhecimento automático das colunas:

| Coluna do ERP | Vira |
|---|---|
| Emissão | Data |
| Pedido | Pedido |
| Cliente | Cliente |
| Família de cliente | Família |
| Segmento de Atuação | Segmento |
| Vendedor 1 | Vendedor |
| Produto | Produto (código) |
| Descrição | Descrição |
| Qtde. Pedido | Quantidade |
| R$ Un. | Preço unit. |
| R$ Pedido | Receita |

Tratamentos automáticos na importação:
- **Linhas secundárias** (continuação sem Produto) são descartadas. Cada pedido entra uma vez.
- **Códigos na frente dos nomes** ("001 - HENKEL", "0005 - QUIMICO") são removidos para agrupar certo.
- Como o extrato **não traz custo**, a margem é calculada pelo motor a partir do SKU. Produtos que ainda não estão no `dados.json` entram só com receita (sinalizados como "sem custo"). Mantenha o catálogo de SKUs atualizado para margem completa.

### Recomendação sobre a base inicial
O `vendas.json` que veio pronto foi montado do seu Power BI e usa nomes completos (ex.: "ALEXANDRE FRANCO"). O ERP abrevia alguns ("ALEXANDRE FRANC"). Para evitar dois rótulos do mesmo vendedor, o caminho mais limpo é: na primeira importação do ERP, escolha **Substituir a base inteira** com o período que você quer. A partir daí, use **Adicionar** todo dia. Assim toda a base fala a mesma língua do ERP.
