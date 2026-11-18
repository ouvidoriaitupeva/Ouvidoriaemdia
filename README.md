# Painel Diário da Ouvidoria

Aplicação estática para carregar a planilha diária da Ouvidoria diretamente no navegador. Nenhum arquivo é enviado a servidor: a base é guardada no IndexedDB do navegador e permanece disponível ao reabrir a página no mesmo navegador, mesmo com textos longos.

## Publicação no GitHub Pages

Envie estes arquivos para a raiz do repositório:

- `index.html`
- `styles.css`
- `app.js`
- `manus-routes.json`

Depois, em **Settings → Pages**, selecione **Deploy from a branch**, escolha a branch principal e a pasta `/ (root)`. A página será publicada pelo GitHub Pages.

## Uso

Abra a página e arraste a planilha Excel. O arquivo deve conter a aba `Dados` e as colunas essenciais do padrão da Ouvidoria. Cada importação diária substitui a base anterior pelo retrato do arquivo atual; isso evita somar históricos de vários dias. O campo `Finalizado em` ou `Finalizados em` é usado apenas internamente para validar a conclusão e não é exibido na exportação. São considerados não concluídos os status `Aberto`, `Agendado`, `Em andamento` e `Pendente`; um registro só é contado como concluído quando possui status `Concluído` e data preenchida na coluna de finalização. O bloco de prazos apresenta duas listas diretas: vencimentos agrupados por dia e atrasados agrupados por secretaria. Para gerar um PDF, use o botão **Imprimir relatório** e selecione **Salvar como PDF** no diálogo do navegador.

## Observação de privacidade

A aplicação não possui login, backend ou banco de dados remoto. A base fica no IndexedDB local do navegador; limpar os dados do site ou trocar de navegador/computador remove a base. Evite compartilhar a tela com dados pessoais.
