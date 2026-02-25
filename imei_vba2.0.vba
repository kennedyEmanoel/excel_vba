Private Sub Worksheet_Change(ByVal Target As Range)
    ' 1. Verifica se a mudança foi na célula B2
    If Not Intersect(Target, Range("B2")) Is Nothing And Range("B2").Value <> "" Then
        
        ' Limpa o status visual anterior para uma nova leitura
        With Range("F2")
            .Value = ""
            .Interior.ColorIndex = xlNone
        End With

        ' 2. Verifica se o IMEI da Coluna C é IGUAL ao da Coluna D
        If Trim(Range("C2").Value) = Trim(Range("D2").Value) Then
            
            Dim wsHist As Worksheet
            Dim proximaLinha As Long
            
            On Error Resume Next
            Set wsHist = ThisWorkbook.Worksheets("Historico")
            On Error GoTo 0
            
            If wsHist Is Nothing Then
                MsgBox "Erro: A aba 'Historico' não foi encontrada!", vbCritical
                Exit Sub
            End If
            
            Application.EnableEvents = False
            
            ' Encontra a próxima linha no histórico
            proximaLinha = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
            
            ' Salva no histórico
            wsHist.Cells(proximaLinha, 1).Value = Range("E2").Value
            wsHist.Cells(proximaLinha, 2).Value = Now
            
            ' Copia para a área de transferência
            Range("E2").Copy
            
            ' --- FEEDBACK VISUAL ---
            With Range("F2")
                .Value = "COPIADO!"
                .Interior.Color = RGB(0, 255, 0) ' Verde brilhante
                .Font.Color = RGB(0, 0, 0)       ' Texto preto
                .Font.Bold = True
            End With
            
            ' Salva o arquivo
            ThisWorkbook.Save 
            
            Application.EnableEvents = True
            
        Else
            ' Caso os IMEIs sejam diferentes, mostra um erro visual em vez de nada
            With Range("F2")
                .Value = "DIVERGENTE!"
                .Interior.Color = RGB(255, 0, 0) ' Vermelho
                .Font.Color = RGB(255, 255, 255) ' Texto branco
                .Font.Bold = True
            End With
        End If
    End If
End Sub