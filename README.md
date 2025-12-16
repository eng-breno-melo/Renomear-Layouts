# Renomeação Automática de Layouts CAD via Excel (VBA)

Este repositório apresenta uma automação desenvolvida em **VBA** para **listagem e renomeação padronizada de layouts em softwares CAD (AutoCAD e GstarCAD)**, utilizando o Excel como interface de controle.

A solução foi criada para otimizar rotinas de projetos elétricos, reduzindo retrabalho manual em arquivos com grande quantidade de layouts. O uso do VBA permite que os nomes sejam facilmente alterados diretamente na planilha, sem necessidade de edição em AutoLISP.

## Funcionalidades

* Conexão automática com AutoCAD ou GstarCAD aberto
* Listagem dos layouts existentes na planilha
* Renomeação em lote baseada nos nomes definidos no Excel
* Limpeza rápida da planilha para novo uso

## Tecnologias

* Excel
* VBA
* AutoCAD / GstarCAD

## Observação
Recomenda-se salvar o arquivo CAD antes da execução das macros.


## Mapeamento dos Botões (Interface Excel)

A interface em Excel foi estruturada com botões vinculados a macros VBA específicas, responsáveis pela integração com o AutoCAD/GstarCAD e pelo gerenciamento dos layouts.

* **Conectar CAD**
  Executa a macro `ConnectCAD`, responsável por estabelecer a conexão COM com uma instância ativa do **AutoCAD ou GstarCAD** e definir o desenho ativo (`ActiveDocument`) para as demais operações.

* **Listar Layouts**
  Executa a macro `ListarLayouts`, que percorre todos os layouts do desenho (exceto *Model*) e exporta seus nomes para a planilha do Excel, permitindo edição manual e padronização antes da aplicação.

* **Renomear Layouts**
  Executa a macro `RenomearLayouts`, que lê os nomes atuais e os novos nomes definidos no Excel e aplica a renomeação diretamente no CAD, com retorno de status (Renomeado, Mantido ou Erro).

* **Apagar**
  Executa a macro `ap`, responsável por limpar o conteúdo das colunas utilizadas na planilha, preparando o ambiente para uma nova execução do fluxo.
