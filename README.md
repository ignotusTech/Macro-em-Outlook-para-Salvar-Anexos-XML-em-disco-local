# Macro em Outlook para Salvar Anexos XML em disco local

Esta macro em VBA foi criada para o Microsoft Outlook. Ela verifica uma pasta selecionada no Outlook e salva os anexos .xml dos emails nessa pasta em uma localização específica no disco local.

## Funcionalidade

- Permite selecionar uma pasta de emails no Outlook.
- Percorre os emails da pasta selecionada.
- Salva os anexos .xml dos emails em uma pasta de destino local.

## Instruções de Uso

1. Abra o Microsoft Outlook.
2. Pressione `Alt + F11` para abrir o Editor do Visual Basic for Applications (VBA).
3. Clique com botão direito na area Projeto e clique em Importar Arquivo e selecione o arquivo `OutlookSalvarAnexoXML.bas`.
4. Clique em `Modulos` para expandir a pasta e clique 2x no `OutlookSalvarAnexoXML`.
5. Modifique o valor da variável `saveFolder` para o caminho onde você deseja salvar os anexos .xml.
6. Salva e feche a tela da Macro (Visual Basic) voltando para a tela do Outlook.
7. Entre na pasta na qual deseja salvar os arquivos '.XML'.
8. Execute a macro pressionando `Alt + F8`, escolhendo `SalvarAnexosXML` e clicando em "Executar".
9. A mensagem 'Anexos .xml salvos na pasta `escolhida no item 5`' aparecerá na tela e terminará de executar a macro salvando os arquivos '.XML' no disco local.

## Requisitos

- Microsoft Outlook.
- Conhecimento básico de VBA pode ser útil para personalizar a macro conforme necessário.

Lembre-se de que é importante ter cuidado ao executar macros e scripts em seu computador. Certifique-se de que o código seja seguro e confiável antes de executá-lo. Se tiver dúvidas ou preocupações, consulte um profissional de TI ou desenvolvedor.

## Contribuição

Se você quiser contribuir para este projeto, fique à vontade para enviar pull requests. Certifique-se de seguir as diretrizes de contribuição.