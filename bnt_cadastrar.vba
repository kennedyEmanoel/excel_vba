Private Sub btn_Cadastrar_Click()
    Dim wsBD As Worksheet
    Dim wsHistorico As Worksheet
    Dim ultLinhaPos As Long, ultLinhaHist As Long
    
    Set wsBD = ThisWorkbook.Sheets("BD_estoque")
    Set wsHistorico = ThisWorkbook.Sheets("Historico_Producao")
    
    If Me.MultiPage1.Value = 0 Then
        Dim rEncontrado As Range
        If Me.txt_ID.Value = "" Or Me.txt_Modelo.Value = "" Or Me.txt_Qtd.Value = "" Then
            MsgBox "Preencha pelo menos o ID, o Modelo e a Quantidade!", vbExclamation, "Campos Obrigatórios"
            Exit Sub
        End If
        
        Set rEncontrado = wsBD.Columns("A").Find(What:=Me.txt_ID.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not rEncontrado Is Nothing Then
            MsgBox "ERRO: O ID " & Me.txt_ID.Value & " já existe no estoque!", vbCritical, "ID Duplicado"
            Me.txt_ID.SetFocus
            Exit Sub
        End If
        
        ultLinhaPos = wsBD.Cells(Rows.Count, 1).End(xlUp).Row + 1
        wsBD.Cells(ultLinhaPos, 1).Value = Me.txt_ID.Value
        wsBD.Cells(ultLinhaPos, 2).Value = Me.txt_Modelo.Value
        wsBD.Cells(ultLinhaPos, 3).Value = Me.txt_Qtd.Value
        wsBD.Cells(ultLinhaPos, 4).Value = Me.txt_Etapa.Value
        wsBD.Cells(ultLinhaPos, 5).Value = Me.txt_Local.Value
        wsBD.Cells(ultLinhaPos, 6).Value = Me.txt_Operador.Value
        wsBD.Cells(ultLinhaPos, 7).Value = Me.txt_peso.Value
        wsBD.Cells(ultLinhaPos, 8).Value = Now()
        
        ' 4. Gravando no Histórico (10 Colunas)
        ultLinhaHist = wsHistorico.Cells(Rows.Count, 1).End(xlUp).Row + 1
        wsHistorico.Cells(ultLinhaHist, 1).Value = ultLinhaHist - 1
        wsHistorico.Cells(ultLinhaHist, 2).Value = Me.txt_ID.Value
        wsHistorico.Cells(ultLinhaHist, 3).Value = Now()
        wsHistorico.Cells(ultLinhaHist, 4).Value = "CRIAÇÃO"
        wsHistorico.Cells(ultLinhaHist, 5).Value = Me.txt_Local.Value
        wsHistorico.Cells(ultLinhaHist, 6).Value = Me.txt_Etapa.Value
        wsHistorico.Cells(ultLinhaHist, 7).Value = Me.txt_Etapa.Value
        wsHistorico.Cells(ultLinhaHist, 8).Value = Me.txt_Operador.Value
        wsHistorico.Cells(ultLinhaHist, 9).Value = Me.txt_Operador.Value
        wsHistorico.Cells(ultLinhaHist, 10).Value = "00:00:00"
        
        MsgBox "Caixa " & Me.txt_ID.Value & " cadastrada com sucesso!", vbInformation, "NB2 Controle"
        
        ' 5. Limpeza
        Me.txt_ID.Value = ""
        Me.txt_peso.Value = ""
        Me.txt_ID.SetFocus
        
        
    ' ==========================================================
    ' ABA 2: CRIAÇÃO EM LOTE (MultiPage Index 1)
    ' ==========================================================
    ElseIf Me.MultiPage1.Value = 1 Then
        Dim i As Integer, qtdLotes As Integer
        Dim ultimoID As Long, novoID As Long
        
        ' 1. Validações
        If Me.cmb_Modelo_Lote.Value = "" Or Me.txt_Qtd_Lote.Value = "" Or Me.txt_Qtd_Por_Caixa.Value = "" Then
            MsgBox "Preencha todos os campos do Lote!", vbExclamation
            Exit Sub
        End If
        
        If Not IsNumeric(Me.txt_Qtd_Lote.Value) Then
            MsgBox "A quantidade de caixas deve ser um número!", vbCritical
            Exit Sub
        End If
        
        qtdLotes = CInt(Me.txt_Qtd_Lote.Value)
        
        ' 2. Descobrindo o último ID para sequência
        ultimoID = Application.WorksheetFunction.Max(wsBD.Columns("A"))
        If ultimoID = 0 Then ultimoID = 1000 ' Começa no 1000 se estiver vazia
        
        Application.ScreenUpdating = False ' Acelera a macro
        
        ' 3. Loop de Criação
        For i = 1 To qtdLotes
            novoID = ultimoID + i
            
            ' DB_estoque
            ultLinhaPos = wsBD.Cells(Rows.Count, 1).End(xlUp).Row + 1
            wsBD.Cells(ultLinhaPos, 1).Value = novoID
            wsBD.Cells(ultLinhaPos, 2).Value = Me.cmb_Modelo_Lote.Value
            wsBD.Cells(ultLinhaPos, 3).Value = Me.txt_Qtd_Por_Caixa.Value
            wsBD.Cells(ultLinhaPos, 4).Value = Me.cmb_Etapa_Lote.Value
            wsBD.Cells(ultLinhaPos, 5).Value = "Estoque"
            wsBD.Cells(ultLinhaPos, 6).Value = Me.cmb_Operador_Lote.Value
            wsBD.Cells(ultLinhaPos, 7).Value = ""
            wsBD.Cells(ultLinhaPos, 8).Value = Now()
            
            ' Histórico (10 Colunas)
            ultLinhaHist = wsHistorico.Cells(Rows.Count, 1).End(xlUp).Row + 1
            wsHistorico.Cells(ultLinhaHist, 1).Value = ultLinhaHist - 1
            wsHistorico.Cells(ultLinhaHist, 2).Value = novoID
            wsHistorico.Cells(ultLinhaHist, 3).Value = Now()
            wsHistorico.Cells(ultLinhaHist, 4).Value = "CRIAÇÃO LOTE"
            wsHistorico.Cells(ultLinhaHist, 5).Value = "Estoque"
            wsHistorico.Cells(ultLinhaHist, 6).Value = Me.cmb_Etapa_Lote.Value
            wsHistorico.Cells(ultLinhaHist, 7).Value = Me.cmb_Etapa_Lote.Value
            wsHistorico.Cells(ultLinhaHist, 8).Value = Me.cmb_Operador_Lote.Value
            wsHistorico.Cells(ultLinhaHist, 9).Value = Me.cmb_Operador_Lote.Value
            wsHistorico.Cells(ultLinhaHist, 10).Value = "00:00:00"
        Next i
        
        Application.ScreenUpdating = True
        
        MsgBox qtdLotes & " caixas cadastradas com sucesso!" & vbCrLf & _
               "IDs gerados: de " & (ultimoID + 1) & " até " & novoID, vbInformation, "Sucesso"
               
        ' 4. Limpeza
        Me.txt_Qtd_Lotes.Value = ""
        Me.txt_Qtd_PorLote.Value = ""
    End If
End Sub
