' ==========================================================
' BOTÃO BUSCAR (Nome do objeto: bnt_buscar)
' ==========================================================
Private Sub bnt_buscar_Click()
    Dim wsBD As Worksheet
    Dim rEnc As Range
    
    ' Valida se tem algo escrito no scanner
    If Me.txt_Scanner.Value = "" Then
        MsgBox "Digite ou bipe o ID da caixa!", vbExclamation
        Me.txt_Scanner.SetFocus
        Exit Sub
    End If
    
    ' Definindo a aba correta (DB_estoque)
    Set wsBD = ThisWorkbook.Sheets("DB_estoque")
    
    ' Procura o ID na Coluna A
    Set rEnc = wsBD.Columns("A").Find(What:=Me.txt_Scanner.Value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rEnc Is Nothing Then
        ' Preenche os Labels de Verificação
        Me.lbl_ID_Verif.Caption = rEnc.Value
        Me.lbl_Modelo_Verif.Caption = rEnc.Offset(0, 1).Value ' Coluna B
        Me.lbl_Qtd_Verif.Caption = rEnc.Offset(0, 2).Value    ' Coluna C
        Me.lbl_Etapa_Verif.Caption = rEnc.Offset(0, 3).Value  ' Coluna D
        Me.lbl_Local_Verif.Caption = rEnc.Offset(0, 4).Value  ' Coluna E
        Me.lbl_Operador_Verif.Caption = rEnc.Offset(0, 5).Value ' Coluna F
        Me.lbl_Peso_Verif.Caption = rEnc.Offset(0, 6).Value   ' Coluna G
        Me.lbl_Data_Verif.Caption = rEnc.Offset(0, 7).Value   ' Coluna H
        
        ' Feedback visual de sucesso (ID fica verde)
        Me.lbl_ID_Verif.ForeColor = RGB(0, 255, 0)
    Else
        MsgBox "ID " & Me.txt_Scanner.Value & " não encontrado!", vbCritical
        LimparTela
    End If
End Sub

' ==========================================================
' DISPARO AUTOMÁTICO (Opcional: busca ao bipar sem clicar)
' ==========================================================
Private Sub txt_Scanner_AfterUpdate()
    ' Isso faz o botão "clicar sozinho" ao dar Enter no Scanner
    Call bnt_buscar_Click
End Sub

' ==========================================================
' BOTÃO INICIAR
' ==========================================================
Private Sub btn_Iniciar_Click()
    Dim wsBD As Worksheet, wsHist As Worksheet
    Dim rEnc As Range
    Dim ultLinhaHist As Long
    
    If Me.lbl_ID_Verif.Caption = "---" Then
        MsgBox "Busque uma caixa válida primeiro!", vbExclamation
        Exit Sub
    End If

    If Me.cmb_Etapa.Value = "" Then
        MsgBox "Selecione a Nova Etapa!", vbExclamation
        Exit Sub
    End If

    Set wsBD = ThisWorkbook.Sheets("DB_estoque")
    Set wsHist = ThisWorkbook.Sheets("Historico")
    
    Set rEnc = wsBD.Columns("A").Find(What:=Me.lbl_ID_Verif.Caption, LookAt:=xlWhole)
    
    If Not rEnc Is Nothing Then
        Dim etpAnt As String: etpAnt = rEnc.Offset(0, 3).Value
        Dim opAnt As String: opAnt = rEnc.Offset(0, 5).Value
        
        rEnc.Offset(0, 3).Value = Me.cmb_Etapa.Value
        rEnc.Offset(0, 4).Value = "Produção"
        rEnc.Offset(0, 5).Value = Me.cmb_Operador.Value
        If Me.txt_Peso_Novo.Value <> "" Then rEnc.Offset(0, 6).Value = Me.txt_Peso_Novo.Value
        rEnc.Offset(0, 7).Value = Now()
        
        ultLinhaHist = wsHist.Cells(Rows.Count, 1).End(xlUp).Row + 1
        wsHist.Cells(ultLinhaHist, 1).Value = ultLinhaHist - 1
        wsHist.Cells(ultLinhaHist, 2).Value = rEnc.Value
        wsHist.Cells(ultLinhaHist, 3).Value = Now()
        wsHist.Cells(ultLinhaHist, 4).Value = "INÍCIO"
        wsHist.Cells(ultLinhaHist, 5).Value = "Produção"
        wsHist.Cells(ultLinhaHist, 6).Value = etpAnt
        wsHist.Cells(ultLinhaHist, 7).Value = Me.cmb_Etapa.Value
        wsHist.Cells(ultLinhaHist, 8).Value = opAnt
        wsHist.Cells(ultLinhaHist, 9).Value = Me.cmb_Operador.Value
        wsHist.Cells(ultLinhaHist, 10).Value = "00:00:00"
        
        MsgBox "Etapa iniciada!", vbInformation
        LimparTela
    End If
End Sub

' ==========================================================
' BOTÃO FINALIZAR
' ==========================================================
Private Sub btn_Finalizar_Click()
    Dim wsBD As Worksheet, wsHist As Worksheet
    Dim rEnc As Range
    Dim ultLinhaHist As Long
    Dim tempoGasto As Double, inicioEtapa As Date
    
    If Me.lbl_ID_Verif.Caption = "---" Then Exit Sub

    Set wsBD = ThisWorkbook.Sheets("DB_estoque")
    Set wsHist = ThisWorkbook.Sheets("Historico")
    
    Set rEnc = wsBD.Columns("A").Find(What:=Me.lbl_ID_Verif.Caption, LookAt:=xlWhole)
    
    If Not rEnc Is Nothing Then
        inicioEtapa = rEnc.Offset(0, 7).Value
        tempoGasto = Now() - inicioEtapa
        
        rEnc.Offset(0, 3).Value = "Concluído: " & rEnc.Offset(0, 3).Value
        rEnc.Offset(0, 4).Value = "Estoque"
        rEnc.Offset(0, 7).Value = Now()
        
        ultLinhaHist = wsHist.Cells(Rows.Count, 1).End(xlUp).Row + 1
        wsHist.Cells(ultLinhaHist, 1).Value = ultLinhaHist - 1
        wsHist.Cells(ultLinhaHist, 2).Value = rEnc.Value
        wsHist.Cells(ultLinhaHist, 3).Value = Now()
        wsHist.Cells(ultLinhaHist, 4).Value = "FINALIZAÇÃO"
        wsHist.Cells(ultLinhaHist, 5).Value = "Estoque"
        wsHist.Cells(ultLinhaHist, 6).Value = Me.lbl_Etapa_Verif.Caption
        wsHist.Cells(ultLinhaHist, 7).Value = "Finalizado"
        wsHist.Cells(ultLinhaHist, 8).Value = Me.lbl_Operador_Verif.Caption
        wsHist.Cells(ultLinhaHist, 9).Value = Me.cmb_Operador.Value
        wsHist.Cells(ultLinhaHist, 10).Value = Format(tempoGasto, "hh:mm:ss")
        
        MsgBox "Finalizado! Tempo: " & Format(tempoGasto, "hh:mm:ss"), vbInformation
        LimparTela
    End If
End Sub