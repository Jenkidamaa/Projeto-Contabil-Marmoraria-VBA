Private Sub acabamento_Change()
        
If acabamento.Value = "Meia Esquadria" Then
    vm = " "
    vm = Int(65)
        

ElseIf acabamento.Value = "Meia Esquadria Importados" Then
    vm = " "
    vm = Int(120)
End If
        
        
End Sub

Private Sub LOC_change()

If LOC.Value = "banheiro" Then
        trabalho.Clear
        
        trabalho.AddItem "Lavatório"
        trabalho.AddItem "Espelho"
        trabalho.AddItem "Cuba esculpida"
        trabalho.AddItem "Nicho"
        trabalho.AddItem "Porta Shampoo"
        
        trabalho.AddItem "Lavatório e espelho"
        trabalho.AddItem "Lavatório e cuba esculpida"
        trabalho.AddItem "Lavatório e nicho"
        
        trabalho.AddItem "Espelho e cuba esculpida"
        trabalho.AddItem "Espelho e nicho"
        
        trabalho.AddItem "Cuba esculpida e nicho"

ElseIf LOC.Value = "cozinha" Then
        trabalho.Clear
        
        
        trabalho.AddItem "Pia"
        trabalho.AddItem "Balcão"
        trabalho.AddItem "Espelhos"
        trabalho.AddItem "Pé"
        trabalho.AddItem "Calha Umida"
        
        
        trabalho.AddItem "Pia e balcão"
        trabalho.AddItem "Pia e espelho"
        trabalho.AddItem "Pia e Pé"
        
        trabalho.AddItem "Balcão e espelho"
        trabalho.AddItem "Balcão e pé"
        
        trabalho.AddItem "Pia, balcão e Espelhos"
        
        trabalho.AddItem "Pia, balcão, espelhos e pé"
        
        trabalho.AddItem "Pia, balcão, espelhos, pé e ilha"
        
        
ElseIf LOC.Value = "Área Gourmet" Then
        trabalho.Clear
        
        trabalho.AddItem "Pia"
        trabalho.AddItem "Balcão"
        trabalho.AddItem "Espelhos"
        trabalho.AddItem "Pé"
        trabalho.AddItem "Calha Umida"
        
        
        trabalho.AddItem "Pia e balcão"
        trabalho.AddItem "Pia e espelho"
        trabalho.AddItem "Pia e Pé"
        
        trabalho.AddItem "Balcão e espelho"
        trabalho.AddItem "Balcão e pé"
        
        trabalho.AddItem "Pia, balcão e Espelhos"
        
        trabalho.AddItem "Pia, balcão, espelhos e pé"
        
        trabalho.AddItem "Pia, balcão, espelhos, pé e ilha"
    
    
ElseIf LOC.Value = "casa" Then
        trabalho.Clear
        trabalho.AddItem "Soleiras e peitoris"

End If

End Sub

Private Sub material_change()
    
    'adicionar valor dos materiais
    
If material.Value = "GRANITO BRANCO ITAUNAS" Then
    valor_pedra = ""
    
    valor_pedra = Int(480)
    
    
ElseIf material.Value = "GRANITO BRANCO ITAUNAS ESCOVADO" Then
    valor_pedra = ""
    valor_pedra = "500,00"
    
ElseIf material.Value = "BRANCO ALASKA" Then
    valor_pedra = ""
    
    
ElseIf material.Value = "BRANCO PRIME" Then
    valor_pedra = ""
    valor_pedra = "950,00"
    
ElseIf material.Value = "BRANCO PARANÁ" Then
    valor_pedra = ""
    valor_pedra = "1350,00"

ElseIf material.Value = "GRANITO VERDE PÉROLA" Then
    valor_pedra = ""
    valor_pedra = "290,00"


ElseIf material.Value = "BRANCO ZEUS" Then
    valor_pedra = ""
    valor_pedra = "1850,00"
    
ElseIf material.Value = "CINZA PRIME" Then
    valor_pedra = ""
    

ElseIf material.Value = "CINZA CLEAN" Then
    valor_pedra = "1690,00"
ElseIf material.Value = "GRANITO BRANCO SIENA" Then
    valor_pedra = Int(480#)
    
    
    

ElseIf material.Value = "MARROM ABSOLUTO" Then
    valor_pedra = ""
    
    
ElseIf material.Value = "MARROM CAFÉ" Then
    valor_pedra = ""
    


ElseIf material.Value = "CREMA MOCA" Then
    valor_pedra = ""
    valor_pedra = 1300#
    
    
ElseIf material.Value = "GRANITO PRETO SÂO GABRIEL" Then
    valor_pedra = ""
    valor_pedra = 480#
    
    
ElseIf material.Value = "GRANITO PRETO SÃO GABRIEL ESCOVADO" Then
    valor_pedra = ""
    valor_pedra = "490,00"
    
ElseIf material.Value = "MÁRMORE TRAVERTINO" Then
    valor_pedra = "350,00"
    
    
ElseIf material.Value = "VERDE CANDEIAS" Then
    valor_pedra = ""
    
    
ElseIf material.Value = "VERDE CURAÇOL" Then
    valor_pedra = ""
    valor_pedra = "1800,00"
    
ElseIf material.Value = "GRANITO VERDE UBATUBA" Then
    valor_pedra = ""
    valor_pedra = "300,00"






End If


End Sub

Private Sub UserForm_Activate()
    Sheets("ORÇAMENTO (2)").Select
    
    
    'adicionar local
    
    LOC.AddItem "Área Gourmet"
    LOC.AddItem "cozinha"
    LOC.AddItem "banheiro"
    LOC.AddItem "casa"
    
    'adicionar material
    
    material.AddItem "GRANITO BRANCO ITAUNAS"
    material.AddItem "GRANITO BRANCO ITAUNAS ESCOVADO"
    material.AddItem "BRANCO PRIME"
    material.AddItem "GRANITO BRANCO DALLAS"
    material.AddItem "BRANCO ALASKA"
    material.AddItem "GRANITO BRANCO SIENA"
    material.AddItem "CINZA PRIME"
    material.AddItem "CINZA CLEAN"
    material.AddItem "MÁRMORE CREMA MOCA"
    material.AddItem "MARROM ABSOLUTO"
    material.AddItem "MARROM CAFÉ"
    material.AddItem "GRANITO PRETO SÃO GABRIEL"
    material.AddItem "GRANITO PRETO SÃO GABRIEL ESCOVADO"
    material.AddItem "MÁRMORE TRAVERTINO"
    material.AddItem "VERDE CANDEIAS"
    material.AddItem "VERDE CURAÇOL"
    material.AddItem "GRANITO VERDE PÉROLA"
    material.AddItem "GRANITO VERDE UBATUBA"
        
    'Montagem
    
    montagem.AddItem "Com Montagem"
    montagem.AddItem "Sem Montagem"
    
    'acabamento
    
    acabamento.AddItem "MÃO DE OBRA RETO SIMPLES"
    acabamento.AddItem "RETO SIMPLES(TRAVERTINO)"
    acabamento.AddItem "RETO DUPLO"
    acabamento.AddItem "RETO DUPLO(TRAVERTINO)"
    acabamento.AddItem "MEIA CANA"
    acabamento.AddItem "MEIA CANA(TRAVERTINO)"
    acabamento.AddItem "ABALOADO SIMPLES"
    acabamento.AddItem "ABALOADO DUPLO"
    acabamento.AddItem "ABALOADO SIMPLES(TRAVERTINO)"
    acabamento.AddItem "ABALOADO DUPLO(TRAVERTINO)"
    acabamento.AddItem "Meia Esquadria"
    acabamento.AddItem "Meia Esquadria Importados"
    acabamento.AddItem "MEIA ESQUADRIA CUBA ESCULPIDA (PORCELANATO)"
    acabamento.AddItem "MEIA ESQUADRIA (NANO GLASS)"
    acabamento.AddItem "MARMORES IMPORTADOS, CARRARA CREMAFIL E OUTROS"

    

End Sub



Private Sub adicionar_Click()

Dim i As Integer
Dim acum As Long
Dim erro As Error
Dim ultima_linha_cel As Integer
Dim Total1 As Long, Total2 As Long
Dim Nome_cliente As String 'Nome que vai para a aba cadastro de orçamentos
Dim planilha As Worksheet



Worksheets("ORÇAMENTO (2)").Select
    
    
    'inserir dados da tabela
    Range("b16") = num_ta
    Range("c17") = LOC
    Range("c18") = material
    Range("b20") = trabalho
    Range("d20") = m2
    Range("e20") = valor_pedra
    
    Range("b21") = acabamento
    Range("d21") = m
    Range("e21") = vm
    Range("B22") = montagem
    Range("b24") = anotacao
    Range("f23") = valor_montagem
    
    'limpar caixas de inserção de dados
    
    num_ta = ""
    LOC = ""
    material = ""
    m2 = ""
    valor_pedra = ""
    acabamento = ""
    m = ""
    anotacao = ""
    trabalho = ""
    vm = ""
    montagem = ""
    
    'adicionar nova tabela de dados
    
    ADD
   
 


continuar = MsgBox("Adicionar nova tabela?", vbYesNo)

If continuar = vbYes Then
    

Else
        remove_tab1
        MsgBox ("Aguarde o salvamento em PDF e Excel")
        Total1 = WorksheetFunction.SumIf(Range("B:B"), "VALOR TOTAL DESTE ORÇAMENTO")
        Total2 = WorksheetFunction.SumIf(Range("B:B"), "VALOR FECHADO A VISTA")
        Nome_cliente = Range("C5")
        
        
        
        SalvarAba
        fechar_aba_orçamento
        Sheets("Lista Orçamentos").Activate
        
        
        
        ultima_linha_cel = Cells(Rows.Count, 2).End(xlUp).Row
        
        Worksheets("Lista Orçamentos").Cells(ultima_linha_cel + 1, 2).Value = Date
        
        Worksheets("Lista Orçamentos").Cells(ultima_linha_cel + 1, 3) = Nome_cliente
        Worksheets("Lista Orçamentos").Cells(ultima_linha_cel + 1, 4) = Total1
        Worksheets("Lista Orçamentos").Cells(ultima_linha_cel + 1, 5) = Total2
        
        'Selection.Range("C" & ultima_linha_cel).Value = Nome_cliente
        'Selection.Range("D" & ultima_linha_cel).Value = Total1
        'Selection.Range("E" & ultima_linha_cel).Value = Total2
        

        tabela.Hide
        'Application.Sheets("ORÇAMENTO (2)").Delete
        
        
End If







End Sub


Sub ADD()
'
' ADD Macro

    Range("B16:F24").Select
    Selection.Copy
    Range("B26").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    Rows("26:26").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

End Sub

Sub remove_tab1()
'
' remove_tab1 Macro

    Rows("16:24").Select
    Range("A24").Activate
    Selection.Delete Shift:=xlUp
End Sub


Sub Macro1()
'
' Macro1 Macro
'

'
    Rows("16:24").Select
    Range("A22").Activate
    Selection.Delete Shift:=xlUp
    Rows("16:18").Select
    Range("A18").Activate
    Selection.Delete Shift:=xlUp
    Range("C14").Select
End Sub




