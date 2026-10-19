# Verificação e limites da revisão

## Verificações realizadas

| Item | Resultado |
| --- | --- |
| Estrutura | TITULAR, INFORMES, NOTAS e TABELAS oculta |
| Totalização | SUM(D10,D15,D20) em INFORMES!C6; valores demonstrativos somam R$ 1.071.296,00 |
| Validações | SIM/NÃO em TITULAR!D16:D18; HOLERITE/CNPJ/FREELANCE em NOTAS!D8:D26 |
| VBA | Projeto presente no XLSM; leitura dos dois eventos nas capturas |
| Telas | TITULAR, INFORMES e NOTAS capturadas no Excel |
| Navegação | Menu lateral levou às telas INFORMES e NOTAS |
| Dados | Fictícios e demonstrativos, conforme confirmação da autora |

## Observações

A inspeção estrutural anterior encontrou valores de erro #VALUE! em TITULAR!O6 e INFORMES!O7. A causa não foi determinada. O arquivo não foi corrigido durante a documentação: a cópia mantém o conteúdo original.

A leitura das macros identificou pontos de atenção em tratamento de erros, alterações em várias células e substituição do objeto anterior. Consulte [MACROS.md](docs/MACROS.md).

## Testes manuais pendentes

Execute em uma cópia com PDFs demonstrativos. Esta lista é um roteiro, não uma declaração de aprovação.

- [ ] Duplo clique em D11, D16 e D21 abre o seletor.
- [ ] Cancelar a seleção mantém o anexo anterior.
- [ ] Selecionar um PDF incorpora o objeto e mostra o nome formatado.
- [ ] Substituir o PDF remove o objeto anterior da célula.
- [ ] Limpar uma célula remove o objeto e restaura a frase.
- [ ] Salvar e reabrir preserva o objeto incorporado.
- [ ] Falhas na incorporação recebem tratamento e restauram os eventos.
- [ ] Limpar várias células é tratado sem erro.
- [ ] Alterar valores em uma cópia recalcula o total.
- [ ] Conferir validações e expansão da tabela de entradas.

As macros de anexação e limpeza não foram executadas nesta revisão. Não houve auditoria integral dos demais módulos VBA.
