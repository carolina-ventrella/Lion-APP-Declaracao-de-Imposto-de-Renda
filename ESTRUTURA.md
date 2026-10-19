# Mapa técnico do LION APP

Este mapa descreve o arquivo `LION_APP_MACRO.xlsm` sem reproduzir os dados preenchidos.

## TITULAR

Os rótulos de cadastro ficam em `C5:C18`, com os respectivos campos em `D5:D18`: nome, CPF, nascimento, título de eleitor, cônjuge, rua, rua abreviada, CEP, telefone, celular, e-mail e três perguntas cadastrais.

`D16:D18` possui validação por lista com `SIM` e `NÃO`, permitindo células vazias. A existência de campos de CPF e e-mail não comprova validação desses formatos.

## INFORMES

| Bloco | Banco | Valor atual | Indicação de PDF |
| --- | --- | --- | --- |
| Primeiro | D9 | D10 | D11 |
| Segundo | D14 | D15 | D16 |
| Terceiro | D19 | D20 | D21 |

O total em `C6` usa `SUM(D10,D15,D20)`. A estrutura analisada não possui validação de dados nessa aba. Portanto, a lista de bancos em `TABELAS` não deve ser descrita como um menu suspenso bancário já confirmado.

## NOTAS

Os cabeçalhos ficam em `C7:E7`: data, categoria e valor. A lista de categorias cobre `D8:D26`, com as opções `HOLERITE`, `CNPJ` e `FREELANCE`, permitindo células vazias. Não foi identificada fórmula de totalização nessa aba.

## TABELAS

Aba oculta com o cabeçalho `bancos` em `A1` e 50 itens em `A2:A51`. Os nomes representam o conteúdo do arquivo, sem validação da situação atual das instituições.

## Componentes adicionais

O pacote contém desenhos, imagens, vínculos e um projeto VBA. A inspeção confirmou a existência desses componentes, sem executar os procedimentos VBA ou validar a navegação na aplicação.
