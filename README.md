# Finanças

Meu projeto em VBA com conexão ao banco de dados ACCESS.

Funciona com um arquivo excel e um arquivo ACCESS que ficam no computador local, onde utilizo a linguagem VBA que faz
toda a conexão com o banco de dados através da liguagem SQL e salva, altera, exclui e busca as informações do banco.

É um sistema completo onde se insere as entradas e saídas, debitando ou creditando o saldo em contas pré cadastradas 
por um interface. São cadastradas classes de custo para que se tenha o histórico de gasto, por exemplo: supermercado, gás, água,
energia, etc...

As compras em cartões de crédito são lançadas em outro formulário, onde se pode cadastrar os cartões de crédito com dados
como: nome do cartão, limite, melhor data para a compra, dia de vencimento, tudo para que o sistema faça os devidos 
cálculos quando for salvar uma compra no crédito. Aqui também é possível colocar em quantas vezes a compra foi dividida 
e o sistema já salva a divisão para todos os meses para que se tenha a visão de quanto está a fatura no futuro.

São várias funcionalidades neste sistema como: Menu em lista com diversos filtros para se buscar uma ou mais compras
no crédito ou débito além de um formulário só para fazer transferência de saldo entre contas cadastradas no sistema.

Uma das funções mais interessantes e que faz com que o sistema seja de fato útil na gestão das finanças pessoais
é um formulário para se cadastrar a provisão de gastos futuros. Neste formulário o usuário pode inserir os gastos ou
receitas previstos(as) para cada mês e assim ir acompanhando o saldo restante para cada classe conforme vai lançando os gastos 
durante o mês. O sistema identifica o que vai sendo debitado nos lançamentos e já traz a diferença para cada classe e 
um resumo traz quanto falta para se pagar todos os gastos daquele mês selecionado conforme o saldo em conta vai diminuindo 
e o previsto vai sendo alcançado. Assim pode se ter uma visão de como está a prévia de fechamento daquele mês.

Além de todas essas funcionalidades, ainda é possível fazer querys em um formuário criado em formado de prompt de comando
SQL para quem deseja fazer SELECTS nas tabelas do banco de dados e fazer pesquisas personalizadas de acordo com sua necessidade.

Emfim, é um sistema muito completo em que eu apliquei todos meus conhecimentos de lógica de programação e SQL usando a
linguagem Visual Basic (VBA). Foi com esta linguagem que me interessei por programação e comecei a fazer diversos
projetos de automação de rotinas em casa e no trabalho.