VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BlocoDeAbas 
   Caption         =   "Sistema Estilo Art Classe A"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12960
   OleObjectBlob   =   "BlocoDeAbas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BlocoDeAbas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Balanço_Click()
    BlocodeAbas.Value = 6
End Sub
Private Sub Cadasstro_clientes_Click()
    BlocodeAbas.Value = 2
End Sub

Private Sub Cadastrar_cliente_Click()
    UserForm1.Show
End Sub

Private Sub Caixa_Click()
    BlocodeAbas.Value = 4
End Sub

Private Sub Chamar_Orçamento_Click()
UserForm1.Show
End Sub

Private Sub CommandButton1_Click()
    Dim w As Workbook
    w.Close
End Sub

Private Sub CommandButton2_Click()
    Dim w As Worksheet
End Sub

Private Sub Controle_Click()
    BlocodeAbas.Value = 5
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Orçamento_Click()
    BlocodeAbas.Value = 1

End Sub

Private Sub PDF_Click()
Dim Arq As String

With Application.FileDialog(msoFileDialogOpen)
    .InitialFileName = Application.DefaultFilePath & "\"
    .Title = "Título"
    .Filters.Clear
    .Filters.ADD "PDF Files", "*.pdf"
    .Show
    
    If .SelectedItems.Count = 0 Then
    Else
    Arq = .SelectedItems(1)
    WebBrowser1.Navigate Arq
    End If
End With

End Sub

Private Sub Recibos_Click()
    BlocodeAbas.Value = 3
End Sub

Private Sub UserForm_Initialize()
    Call AtualizaCaixaListagemRecibos
    Call AtualizaCaixaListagemCaixa
    Call AtualizaCaixaListagemCadastro
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Viewer1_OnDocumentLoaded()
    Dim Arq As Object
End Sub
