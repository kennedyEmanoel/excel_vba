Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("B2")) Is Nothing And Range("B2").Value <> "" Then

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
            
            proximaLinha = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
            
            wsHist.Cells(proximaLinha, 1).Value = Range("E2").Value
            wsHist.Cells(proximaLinha, 2).Value = Now

            'copiar
            'Range("E2").Copy
            
            ThisWorkbook.Save 
            
            Application.EnableEvents = True
            
        End If
    End If
End Sub 