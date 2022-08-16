# linkedin_to_excel
Exporte a sua lista de certificados do Linkedin para uma planilha Excel.

## Objetivos
1. Script feito em Python, serve para você obter a sua lista de certificados e licenças do LinkedIn em formato Excel (.xlsx).
2. Tudo é realizado com apenas um click. Ao executar, o script faz tudo automaticamente (você não precisa estar logado no LinkedIn).  
3. Útil para quando você precisa apresentar a sua lista de certificados em formato físico ou então precisa enviar a lista por e-mail.
4. Recomendado para pessoas que possuam centenas de certificados no LinkedIn ou que desejam classificar os certificados por instituição.
5. Facilita a busca por certificados com código de verificação e sem código de verificação.

## Estrutura
1. A planilha é dividida em 04 (quatro) colunas: Certificado, Instituição, Data de Emissão e Código da Credencial.
2. Caso o certificado não tenha código de verificação, é exibido o valor "Sem código" na célula respectiva.

## Requisitos
1. Ter o [Python](https://www.python.org/downloads/) instalado no computador.
2. Ter as seguintes bibliotecas instaladas: [openpyxl](https://openpyxl.readthedocs.io/en/stable/), [selenium](https://selenium-python.readthedocs.io/), [pyperclip](https://pypi.org/project/pyperclip/), [pyautogui](https://pyautogui.readthedocs.io/en/latest/).

## OBSERVAÇÕES
1. Nenhuma informação pessoal do LinkedIn é repassada para servidores de terceiros, o script é executado no sistema do usuário, de acordo com o e-mail e senha fornecidos.
2. O código é voltado para o site em português do LinkedIn (https://br.linkedin.com/). Não irá funcionar em outros idiomas.
3. No LinkedIn, é necessário que sejam preenchidos todos os campos obrigatórios do certificado (nome, instituição, data de emissão), sendo opcional o código de verificação (pode deixar em branco). Sem essas informações, o código não irá funcionar corretamente e a planilha ficará desorganizada.
4. É necessário alterar os campos de e-mail e senha no código, para poder acessar a sua conta no Linkedin e obter os nomes dos certificados.
5. Tambem é necessário alterar o campo de número de certificado no código, para corresponder ao número total de certificados que você possui.
6. Caso o código não execute alguma função no tempo certo, é possível alterar o tempo de espera entre cada comando para se adequar ao seu computador. Basta modificar o campo "sleep()" e colocar o valor númerico que deseja (em segundos) no espaço entre parênteses ( ). Ex: sleep(1) = 1 segundo. [Sobre sleep](https://realpython.com/python-sleep/).
7. O programa demora um pouco para executar todos os comandos, tenha paciência e aguarde até o script ser finalizado. Se tudo ocorrer bem, irá aparecer ao final a mensagem "Excel gerado com sucesso".
8. Em alguns casos, o LinkedIn costuma exigir um código de verificação ao tentar logar no site. Caso isso aconteça, resolva essa etapa de verificação, feche o navegador e reinicie o script, para poder gerar a planilha corretamente.
9. Após finalizar a execução do script, serão gerados dois arquivos dentro da pasta do script: um arquivo em formato de texto ("testeselenium.txt") - onde os dados da página de certificado serão salvos, e a planilha com a lista de certificados - salva em formato xlsx ("lista_certificados.xlsx").
10. O código ainda está em estágio de desenvolvimento, podem ocorrer bugs e outras falhas inesperadas.
11. O código não está otimizado, fique à vontade para aprimorá-lo ou modificá-lo de acordo com as suas necessidades.
12. Sugestões e críticas são bem vindas :)

## Próximas atualizações
1. Organizar o tamanho das colunas para autoajustar ao conteúdo das células.
2. Obter dados de outras abas do LinkedIn (atividades, experiência, formação acadêmica, etc).
3. Adicionar uma interface para o usuário digitar o seu e-mail e senha do LinkedIn, antes de executar o script.
4. Mostrar em gráfico a quantidade de certificados, por instituição.
5. Adicionar suporte para outros idiomas.

*Em desenvolvimento...*
