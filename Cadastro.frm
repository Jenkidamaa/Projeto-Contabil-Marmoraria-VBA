Private Sub btnadd_Click()
    Dim resp As VbMsgBoxResult
    Dim resp1 As VbMsgBoxResult
    Dim nome As String  
    
    resp1 = MsgBox("Adicionar Cadastro ao banco de dados", vbYesNo)
    resp = MsgBox("Proseguir para orçamento?", vbYesNo)
    If resp1 = vbYes Then
        Sheets("CADASTRO").Activate
        
        Dim X As Integer
        X = Cells(Rows.Count, 4).End(xlUp).Row
        
    
        'adicionar os dados na tabela
        Worksheets("CADASTRO").Cells(X + 1, 2) = cd_cliente
        Worksheets("CADASTRO").Cells(X + 1, 3) = nome
        Worksheets("CADASTRO").Cells(X + 1, 4) = endereco
        Worksheets("CADASTRO").Cells(X + 1, 5) = bairro
        Worksheets("CADASTRO").Cells(X + 1, 6) = Ne
        Worksheets("CADASTRO").Cells(X + 1, 7) = cidade
        Worksheets("CADASTRO").Cells(X + 1, 8) = estado
        Worksheets("CADASTRO").Cells(X + 1, 9) = email
        Worksheets("CADASTRO").Cells(X + 1, 10) = telefone
        
        If resp = vbYes Then
            
            UserForm1.Hide
            nome = "ORÇAMENTO"
            
            ThisWorkbook.Sheets(nome).Copy Before:=Sheets(1)
            
            Sheets("ORÇAMENTO (2)").Select
            Range("C5") = nome
            Range("C6") = endereco
            'Range("C7") = bairro
            'Range("C8") = cidade
            'Range("C9") = telefone
            'Range("C10") = email
                
            tabela.Show
        
    Else
        
        Range("C5") = nome
        Range("C6") = endereco
        Range("C7") = bairro
        Range("C8") = cidade
        Range("C9") = telefone
        Range("C10") = email   
       
        resp2 = MsgBox("Fechar janela de usuarios", vbYesNo)
            
            If resp2 = vbYes Then
            
            Application.Quit
            
            
            End If          
           
    End If
    End If
End Sub





'Private Sub cd_cliente_Change()
    
      '  Dim y As Integer
        
        
       ' If cd_cliente = Worksheets("CADASTRO").Cells(y, 2) Then
        
        
          '  nome = Worksheets("CADASTRO").Cells(y, 3)
          '  endereco = Worksheets("CADASTRO").Cells(y, 4)
          '  bairro = Worksheets("CADASTRO").Cells(y, 5)
         '   Ne = Worksheets("CADASTRO").Cells(y, 6)
          '  cidade = Worksheets("CADASTRO").Cells(y, 7)
           ' estado = Worksheets("CADASTRO").Cells(y, 8)
            'email = Worksheets("CADASTRO").Cells(y, 9)
            'telefone = Worksheets("CADASTRO").Cells(y, 10)
            
            
       ' End If
        

Private Sub estado_Change()
    If estado.Value = "AC" Then
        testado = "Acre"
    ElseIf estado.Value = "AL" Then
        testado = "Alagoas"
    ElseIf estado.Value = "AP" Then
        testado = "Amapá"
    ElseIf estado.Value = "AM" Then
        testado = "Amazonas"
    ElseIf estado.Value = "BA" Then
        testado = "Bahia"
    ElseIf estado.Value = "CE" Then
        testado = "Ceará"
    ElseIf estado.Value = "ES" Then
        testado = "Espírito Santo"
    ElseIf estado.Value = "GO" Then
        testado = "Goiás"
    ElseIf estado.Value = "MA" Then
        testado = "Maranhão"
    ElseIf estado.Value = "MT" Then
        testado = "Mato Grosso"
    ElseIf estado.Value = "MS" Then
        testado = "Mato Grosso do Sul"
    ElseIf estado.Value = "MG" Then
        testado = "Minas Gerais"
    ElseIf estado.Value = "PA" Then
        testado = "Pará"
    ElseIf estado.Value = "PB" Then
        testado = "Paraíba"
    ElseIf estado.Value = "PR" Then
        testado = "Paraná"
    ElseIf estado.Value = "PE" Then
        testado = "Pernambuco"
    ElseIf estado.Value = "PI" Then
        testado = "Piauí"
    ElseIf estado.Value = "RJ" Then
        testado = "Rio de Janeiro"
    ElseIf estado.Value = "RN" Then
        testado = "Rio Grande do Norte"
    ElseIf estado.Value = "RS" Then
        testado = "Rio Grande do Sul"
    ElseIf estado.Value = "RO" Then
        testado = "Rondônia"
    ElseIf estado.Value = "RR" Then
        testado = "Roraima"
    ElseIf estado.Value = "SC" Then
        testado = "Santa Catarina"
    ElseIf estado.Value = "SP" Then
        testado = "São Paulo"
    ElseIf estado.Value = "SE" Then
        testado = "Sergipe"
    ElseIf estado.Value = "DF" Then
        testado = "Distrito Federal"
    
    End If
End Sub
Private Sub UserForm_Activate()
    abrir_novo_orçamento
    

    
    estado.AddItem "AC"
    estado.AddItem "AL"
    estado.AddItem "AP"
    estado.AddItem "AM"
    estado.AddItem "BA"
    estado.AddItem "CE"
    estado.AddItem "ES"
    estado.AddItem "GO"
    estado.AddItem "MA"
    estado.AddItem "MT"
    estado.AddItem "MS"
    estado.AddItem "MG"
    estado.AddItem "PA"
    estado.AddItem "PB"
    estado.AddItem "PR"
    estado.AddItem "PE"
    estado.AddItem "PI"
    estado.AddItem "RJ"
    estado.AddItem "RN"
    estado.AddItem "RS"
    estado.AddItem "RO"
    estado.AddItem "RR"
    estado.AddItem "SC"
    estado.AddItem "SP"
    estado.AddItem "SE"
    estado.AddItem "TO"
    estado.AddItem "DF"    
End Sub

