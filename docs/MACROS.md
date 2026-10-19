# Documentação técnica das macros

## Fontes e escopo

Descrição baseada em [vba1.png](../images/vba1.png), [vba2.png](../images/vba2.png) e [vba3.png](../images/vba3.png), do módulo Planilha2 (INFORMES). O arquivo contém um projeto VBA; os demais módulos e formulários não foram inventariados. As imagens documentam o código existente, sem alteração da planilha.

## Eventos e responsabilidades

| Procedimento | Gatilho | Função |
| --- | --- | --- |
| Worksheet_BeforeDoubleClick | Duplo clique | Selecionar e incorporar ou substituir PDF |
| Worksheet_Change | Alteração de conteúdo | Remover o objeto e restaurar a orientação se a célula ficar vazia |

Ambos restringem a atuação com `Intersect(Target, Range("D11, D16, D21"))`. Por serem eventos de planilha, pertencem ao módulo da aba INFORMES, e não a um módulo padrão de macros executadas manualmente.

## Variáveis do duplo clique

| Variável | Tipo | Uso |
| --- | --- | --- |
| Filtro | String | Filtro de arquivos PDF |
| TituloJanela | String | Texto do seletor |
| CaminhoArquivo | Variant | Caminho escolhido ou False se cancelado |
| ObjetoPDF | Object | Objeto OLE incorporado |
| ShapeAntigo | Shape | Item percorrido para remover o anexo anterior |
| NomeArquivo | String | Nome obtido por Dir |

## Incorporação e substituição

`GetOpenFilename` devolve um caminho, mas não incorpora o documento. A rotina só continua se `CaminhoArquivo <> False`, de modo que cancelar a seleção não deve executar a substituição.

O código percorre `Me.Shapes`, verifica os tipos `msoOLEControlObject` ou `msoEmbeddedOLEObject` e compara `TopLeftCell.Address` com `Target.Address`. Objetos correspondentes são removidos.

`Application.EnableEvents = False` evita que `Target.ClearContents` e a escrita do nome acionem Worksheet_Change durante a inclusão. O objeto é inserido por `OLEObjects.Add`, com `Link:=False` e `DisplayAsIcon:=True`.

As propriedades Top, Left, Width e Height alinham o objeto à célula. `Placement = xlMoveAndSize` acompanha sua posição e tamanho. `Visible = False` oculta o objeto, mas não o remove do documento.

O nome recebe Calibri 11, sublinhado e cor `RGB(0, 102, 204)`, com alinhamentos horizontal e vertical centralizados. A rotina reativa os eventos e informa o sucesso.

## Limpeza

Ao encontrar uma célula de anexo vazia, Worksheet_Change desativa os eventos, remove objetos OLE ancorados nela e restaura a orientação de seleção. A fonte retorna à cor automática e perde o sublinhado. No final, os eventos são reativados.

Não há comando nas capturas que apague o PDF original do computador.

## Distinções importantes

- Formatação azul sublinhada não é `Hyperlinks.Add`.
- Não aparece uma rotina para abrir o PDF. O duplo clique inicia seleção ou substituição.
- Um PDF oculto continua incorporado e pode aumentar o tamanho do XLSM.
- Os elementos de navegação possuem vínculos internos. Não são descritos como macros sem evidência de atribuição de código.

## Pontos de evolução identificados por leitura

Estes pontos não foram alterados no arquivo demonstrativo:

1. **Tratamento de erros:** não aparece um `On Error` que garanta restaurar `Application.EnableEvents` após falha. Um erro na incorporação pode deixar os eventos desligados na sessão.
2. **Alteração em várias células:** `Target.Value = ""` pode falhar se Target contiver várias células e retornar uma matriz. O fluxo documentado recomenda limpar uma célula de anexo por vez.
3. **Preservação do anexo anterior:** a remoção ocorre antes da incorporação do novo PDF. Uma falha nessa etapa pode deixar a célula sem o anexo anterior.
4. **Abertura do PDF:** seria uma funcionalidade adicional; não consta dos procedimentos mostrados.
5. **Compatibilidade:** testar a incorporação no Excel e no ambiente OLE do computador usado para avaliação.

## Como revisar no Excel

Use Alt + F11, localize Microsoft Excel Objetos e abra Planilha2 (INFORMES). Compare os procedimentos com as capturas. Execute os testes de inclusão, cancelamento, substituição e limpeza em uma cópia, usando PDFs demonstrativos.
