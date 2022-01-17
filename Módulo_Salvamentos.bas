Sub salvar_()
'
' salvar_ Macro
'

'
    ActiveWorkbook.Save
End Sub

Sub contar()
    X = Cells(Rows.Count, 4).End(xlUp).Row
    MsgBox (X)
    
End Sub

Sub SALVAR_PDF()

        
        Dim arquivo As String

        arquivo = "C:\Users\andreia limoli\Desktop\TRABALAHO ORÇAMENTOS\ORCAMENTOS PDF\" & Range("C5").Value & ".pdf"



        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        arquivo, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
End Sub
Sub SALVAR_EXCEL()
                       
        Dim arquivo1 As String

        arquivo1 = "C:\Users\andreia limoli\Desktop\TRABALAHO ORÇAMENTOS\ORCAMENTOS EXCEL\" & Range("C5").Value & ".xlsx"



    ChDir "C:\Users\andreia limoli\Desktop"
    ActiveWorkbook.SaveAs Filename:= _
        arquivo1 _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.WindowState = xlNormal
    
End Sub


Sub SalvarAba()
'Impede que o Excel atualize a tela
Application.ScreenUpdating = False
'Impede que o Excel exiba alertas
Application.DisplayAlerts = False

'Seta uma variável para se referir a nova pasta de trabalho
Dim NovoWB As Workbook
Dim ultima_linha As Long
Dim Total1 As Long, Total2 As Long


'Cria esta nova aba
Set NovoWB = Workbooks.ADD(xlWBATWorksheet)


With NovoWB
'Copia a aba atual para o novo arquivo, como a segunda aba

ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count)
'Deleta a primeira aba do arquivo criado (Aba em branco)

.Worksheets(1).Delete

'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir

.Worksheets.Select
ultima_linha = ActiveSheet.Range("D105810").End(xlUp)
ActiveSheet.PageSetup.PrintArea = "$A$1:$E" & ultima_linha
.SaveAs ThisWorkbook.Path & " " & Range("C5") & ".xlsx"
'Fecha o novo arquivo
'ActiveWorkbook.SaveAs Filename:="C:\Users\Carlos Isaque\Desktop\LOL\" & Range("C5") & ".xlsx"

.Close True
'Application.Sheets("ORÇAMENTO (2)").Delete



End With

'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub

