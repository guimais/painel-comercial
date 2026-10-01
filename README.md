# 📊 Painel Comercial de Leads

Um painel simples e direto para acompanhar os leads do time comercial. A ideia é que gestores e vendedores consigam ver, num só lugar, em que etapa cada negociação está, quem está cuidando dela e qual foi o esforço de cada pessoa ao longo do semestre.

O painel não olha só para quem fechou negócio. Ele também mostra quem manteve constância, fez follow-up e recebeu respostas positivas, para que o feedback de fim de semestre seja mais justo.

Você pode ver a versão publicada aqui: [guimais.github.io/painel-comercial](https://guimais.github.io/painel-comercial/)

## O que dá para fazer

**Acompanhar os leads numa tabela completa.** Cada linha mostra o lead, a empresa, a etapa do funil, a data do último contato, as respostas positivas, as ações realizadas, a constância em dias, o responsável, a próxima ação e o status.

**Ver os números do funil na hora.** No topo ficam os indicadores de total de leads, negociações em andamento, ganhos, perdidos, taxa de conversão, constância média, respostas positivas e média de atividades por lead. Tudo se atualiza conforme os filtros aplicados.

**Filtrar e buscar.** Dá para filtrar por etapa, status e responsável, ou buscar por nome, empresa ou próxima ação. A busca ignora acentos, então "joao" encontra "João".

**Ordenar qualquer coluna.** Basta clicar no cabeçalho. Datas e números são ordenados pelo valor real, e não como texto.

**Receber insights rápidos.** Um painel logo acima da tabela mostra quem tem mais respostas positivas, quem registrou mais atividades e quais leads estão parados há mais tempo.

**Ver os destaques do mês.** O botão Destaques monta um ranking com os três leads de melhor desempenho, levando em conta respostas, atividades e constância. Eles também ganham uma coroa na tabela.

**Anotar observações sobre cada lead.** No painel de observações, cada lead aparece numa lista que abre ao clicar. Ali dá para escrever notas e percepções, que ficam salvas no próprio navegador. Um pontinho azul ao lado do nome indica quem já tem anotação.

**Exportar e copiar.** A tabela filtrada pode ser exportada em CSV ou Excel, já com acentos corretos e com a coluna de motivo de perda. Também dá para copiar tudo e colar direto numa planilha.

**Editar direto na tabela.** Com o modo PLUS ativo, um duplo clique permite alterar a próxima ação, o responsável e o status. Quando um lead é marcado como perdido, o painel pergunta o motivo e mostra essa informação nas observações.

## Atalhos de teclado

| Tecla | O que faz |
| :--- | :--- |
| F | Limpa todos os filtros |
| G | Abre os destaques do mês |
| E | Exporta em CSV |
| Shift + E | Exporta em Excel |
| Ctrl + Shift + P | Liga e desliga o modo PLUS |
| Esc | Fecha a janela de destaques |

## Pensado para iPad e celular

O painel foi feito para funcionar bem em telas de toque. No celular, a coluna com o nome do lead fica fixa enquanto você rola a tabela para o lado, os filtros deslizam horizontalmente e os campos de texto não dão aquele zoom automático do iOS. O layout também respeita as bordas arredondadas e a área segura dos aparelhos mais novos.

Também houve cuidado com acessibilidade: tudo pode ser usado pelo teclado, o foco fica sempre visível e as animações são reduzidas para quem ativou essa opção no sistema.

## Tecnologias

O projeto usa só HTML, CSS e JavaScript puro, sem frameworks e sem dependências para instalar. Os ícones vêm do Remix Icon e as fontes são Plus Jakarta Sans e Manrope, do Google Fonts. A publicação é feita pelo GitHub Pages.

As anotações, os motivos de perda e os filtros escolhidos ficam guardados no navegador de quem está usando. Isso significa que cada pessoa vê as próprias notas e que nada é enviado para um servidor.

## Estrutura do projeto

    painel-comercial/
        index.html
        assets/
            css/
                styles.css
            js/
                main.js
            img/
                preview.png

## Como rodar localmente

Clone o repositório:

    git clone https://github.com/guimais/painel-comercial.git

Depois é só abrir o arquivo `index.html` no navegador. Não precisa instalar nada.

## Autor

Feito por Guilherme Mais.
Você pode me encontrar no GitHub em [github.com/guimais](https://github.com/guimais).
