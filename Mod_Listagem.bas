Attribute VB_Name = "Mod_Listagem"
Sub AtualizaCaixaListagemRecibos()

Dim AbaOrçamento As Object
Dim ultLinha As Long


Set AbaOrçamento = Sheets("RECIBOS1")

ultLinha = AbaOrçamento.Range("A10000").End(xlUp)


BlocodeAbas.ListBoxRecibos.ColumnCount = 6
BlocodeAbas.ListBoxRecibos.ColumnHeads = True
BlocodeAbas.ListBoxRecibos.ColumnWidths = "65;107.74;107;80;107.74;107.74;"
BlocodeAbas.ListBoxRecibos.RowSource = "RECIBOS1!A2:F" & ultLinha

End Sub

Sub AtualizaCaixaListagemCaixa()

Dim AbaCAIXA As Object
Dim ultLinha As Long


Set AbaCAIXA = Sheets("CAIXA")

ultLinha = AbaCAIXA.Range("A10000").End(xlUp)


BlocodeAbas.ListBoxCaixa.ColumnCount = 9
BlocodeAbas.ListBoxCaixa.ColumnHeads = True
BlocodeAbas.ListBoxCaixa.ColumnWidths = "49;29;145;89;89;59;59;69;69;"
BlocodeAbas.ListBoxCaixa.RowSource = "CAIXA!A2:I" & ultLinha




End Sub

Sub AtualizaCaixaListagemCadastro()

Dim AbaCadastro As Object
Dim ultLinha As Long


Set AbaCadastro = Sheets("CADASTRO")

ultLinha = AbaCadastro.Range("A10000").End(xlUp)


BlocodeAbas.ListBoxCadastro.ColumnCount = 9
BlocodeAbas.ListBoxCadastro.ColumnHeads = True
BlocodeAbas.ListBoxCadastro.ColumnWidths = "49;89;89;89;25;110;59;69;69;"
BlocodeAbas.ListBoxCadastro.RowSource = "CADASTRO!A2:I" & ultLinha




End Sub

