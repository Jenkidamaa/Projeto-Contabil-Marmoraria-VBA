VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tabela 
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   OleObjectBlob   =   "Tabela para cria��o de or�amentos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




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
        
        trabalho.AddItem "Lavat�rio"
        trabalho.AddItem "Espelho"
        trabalho.AddItem "Cuba esculpida"
        trabalho.AddItem "Nicho"
        trabalho.AddItem "Porta Shampoo"
        
        trabalho.AddItem "Lavat�rio e espelho"
        trabalho.AddItem "Lavat�rio e cuba esculpida"
        trabalho.AddItem "Lavat�rio e nicho"
        
        trabalho.AddItem "Espelho e cuba esculpida"
        trabalho.AddItem "Espelho e nicho"
        
        trabalho.AddItem "Cuba esculpida e nicho"

ElseIf LOC.Value = "cozinha" Then
        trabalho.Clear
        
        
        trabalho.AddItem "Pia"
        trabalho.AddItem "Balc�o"
        trabalho.AddItem "Espelhos"
        trabalho.AddItem "P�"
        trabalho.AddItem "Calha Umida"
        
        
        trabalho.AddItem "Pia e balc�o"
        trabalho.AddItem "Pia e espelho"
        trabalho.AddItem "Pia e P�"
        
        trabalho.AddItem "Balc�o e espelho"
        trabalho.AddItem "Balc�o e p�"
        
        trabalho.AddItem "Pia, balc�o e Espelhos"
        
        trabalho.AddItem "Pia, balc�o, espelhos e p�"
        
        trabalho.AddItem "Pia, balc�o, espelhos, p� e ilha"
        
        
ElseIf LOC.Value = "�rea Gourmet" Then
        trabalho.Clear
        
        trabalho.AddItem "Pia"
        trabalho.AddItem "Balc�o"
        trabalho.AddItem "Espelhos"
        trabalho.AddItem "P�"
        trabalho.AddItem "Calha Umida"
        
        
        trabalho.AddItem "Pia e balc�o"
        trabalho.AddItem "Pia e espelho"
        trabalho.AddItem "Pia e P�"
        
        trabalho.AddItem "Balc�o e espelho"
        trabalho.AddItem "Balc�o e p�"
        
        trabalho.AddItem "Pia, balc�o e Espelhos"
        
        trabalho.AddItem "Pia, balc�o, espelhos e p�"
        
        trabalho.AddItem "Pia, balc�o, espelhos, p� e ilha"
    
    
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
    
ElseIf material.Value = "BRANCO PARAN�" Then
    valor_pedra = ""
    valor_pedra = "1350,00"

ElseIf material.Value = "GRANITO VERDE P�ROLA" Then
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
    
    
ElseIf material.Value = "MARROM CAF�" Then
    valor_pedra = ""
    


ElseIf material.Value = "CREMA MOCA" Then
    valor_pedra = ""
    valor_pedra = 1300#
    
    
ElseIf material.Value = "GRANITO PRETO S�O GABRIEL" Then
    valor_pedra = ""
    valor_pedra = 480#
    
    
ElseIf material.Value = "GRANITO PRETO S�O GABRIEL ESCOVADO" Then
    valor_pedra = ""
    valor_pedra = "490,00"
    
ElseIf material.Value = "M�RMORE TRAVERTINO" Then
    valor_pedra = "350,00"
    
    
ElseIf material.Value = "VERDE CANDEIAS" Then
    valor_pedra = ""
    
    
ElseIf material.Value = "VERDE CURA�OL" Then
    valor_pedra = ""
    valor_pedra = "1800,00"
    
ElseIf material.Value = "GRANITO VERDE UBATUBA" Then
    valor_pedra = ""
    valor_pedra = "300,00"






End If


End Sub

Private Sub UserForm_ACTIVATE()
    Sheets("OR�AMENTO (2)").Select
    
    
    'adicionar local
    
    LOC.AddItem "�rea Gourmet"
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
    material.AddItem "M�RMORE CREMA MOCA"
    material.AddItem "MARROM ABSOLUTO"
    material.AddItem "MARROM CAF�"
    material.AddItem "GRANITO PRETO S�O GABRIEL"
    material.AddItem "GRANITO PRETO S�O GABRIEL ESCOVADO"
    material.AddItem "M�RMORE TRAVERTINO"
    material.AddItem "VERDE CANDEIAS"
    material.AddItem "VERDE CURA�OL"
    material.AddItem "GRANITO VERDE P�ROLA"
    material.AddItem "GRANITO VERDE UBATUBA"
        
    'Montagem
    
    montagem.AddItem "Com Montagem"
    montagem.AddItem "Sem Montagem"
    
    'acabamento
    
    acabamento.AddItem "M�O DE OBRA RETO SIMPLES"
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

Worksheets("OR�AMENTO (2)").Select
    
    
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
    
    'limpar caixas de inser��o de dados
    
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
        
        SalvarAba
        fechar_aba_or�amento
        Worksheets("MENU").Activate
        
        
        
        
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




