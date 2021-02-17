# Linguagem VBA

Criei esse repositório para compartilhar funções que utilizo em minhas automações diárias no trabalho.

## mes_valido

Função que verifica se o número correspondente ao mês informado é válido. (entre 1 e 12)

## mes_solicitado 

Função que sugere ao usuário a opção de selecionar o mês anterior em uma caixa de mensagem, caso não solicite o mês anterior, valida o novo mês.   

## diretorio_default

Função que retorna o diretório default utilizado conforme o mês que foi soliictado.

## ToColletter

Recebe uma referência de célula como parâmetro e devolve a letra correspondente a coluna.

## ToColNum

Recebe uma referência de célula como parâmetro e devolve número correspondente a coluna.

## out_diferenca

Função que recebe strOld e strCurrent como Strings com nomes separado por um mesmo delimitor.
Retorna o que estava na strOld e não está mais na strCurrent. 

## new_diferenca

Função que recebe strOld e strCurrent como Strings com nomes separado por um mesmo delimitor.
Retorna o que está na strCurrent e que não estava na strOld.
