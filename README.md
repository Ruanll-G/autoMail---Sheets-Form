# Envio Automático de Email por Sheets e Forms

Dentro de "autoMail.js" contém um script para o Google Sheets Planilhas com uma automação de envio de email automático de acordo com as respostas de um formulário, vinculada junto a planilha.

# Como colocar na Planilha

Para adicionar essa automação, abra o arquivo de "autoMail.js", copie o script dentro dele e insera dentro do AppScript da Planilha que você deseja.
Faça as devidas edições no código HTML dentro do código para seu caso de mensagem.
Após colocar o código dentro do AppScript da Planilha crie um Acionador/Trigger para iniciar o processo. Quando colocar a função para iniciar a automação coloque o gatilho para a função: aoEnviarFormulario.

# Importante!!

O código foi feito para se utilizar Alias, então certifique-se que o Alias desejado realmente tenha permissão para mandar email dentro de sua conta.
