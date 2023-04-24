Public doc As Object
Private word As New word.Application
Private cam As String
Private CamDest As String
Private Plan2 As Worksheet
Private colmodelo As String
Private ColNomeArq As String
Private colTipArq As String
Private lf As Integer
Private li As Integer
Private colarr
Private dados
Private tempo_espera

Public roboStart As Boolean


Public Sub Minuta()

    colarr = colConvert(Array("C", "D", "E", "F", "G", "H", "K", "L", "K", "L", "M", "N", "O", "P", "Q", "R", "S"))
    Call Preenche_DOC_Lote_plan(colarr, "Suspensao", "B", "G")

End Sub

Public Sub Gerar_MalaDireta()

    colarr = colConvert(Array("E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"))
    
   ' Call Preenche_DOC_Lote_plan(colarr, "MalaDireta", "B", "C", "D")
    Call Preenche_DOC_Lote_plan(colarr, "MalaDireta", 2, 3, 4)

End Sub





Public Sub Preenche_DOC_Lote_plan(colarr1, plan, colmodelo1, ColNomeArq1, colTipArq1)
    On Error GoTo erro
    
    colmodelo = colmodelo1
    ColNomeArq = ColNomeArq1
    colTipArq = colTipArq1
    colarr = colarr1
    tempo = Planilha6.[B1].Value
    
    If tempo Then tempo_espera = "00:00:01" Else: tempo_espera = "00:00:00"
    
    Set Plan2 = ThisWorkbook.Sheets(plan)
    
    lf = Plan2.[A1000000].End(xlUp).Row
    dados = Plan2.Range("A1:T" & lf).Value
    
    cam = ThisWorkbook.Path & "\Peticao"
    CamDest = ThisWorkbook.Path & "\MalaDireta"
    
    lf = ultlin(Plan2, "A")
    UserForm1.Ttotal = lf - 1
    li = ultlin(Plan2, "T") + 1
    
    Call Preenche_DOC_loop

Exit Sub
erro: MsgBox "Erro na geração do Arquivo"
End Sub

Private Function Preenche_DOC_loop()
    
    Dim texto(1000), TextoNovo(1000)
    'arq = Plan2.Cells(li, colmodelo)
    arq = dados(li, colmodelo)
    
    UserForm1.Tcont = li - 1
    
    For c = 0 To UBound(colarr)
        'texto(c) = Plan2.Cells(1, colarr(c))
        'TextoNovo(c) = Plan2.Cells(li, colarr(c))
        texto(c) = dados(1, colarr(c))
        TextoNovo(c) = dados(li, colarr(c))
    Next
    
    'TipoArq = Plan2.Cells(li, colTipArq)
    'ArqNovo = Plan2.Cells(li, ColNomeArq)
    TipoArq = dados(li, colTipArq)
    ArqNovo = dados(li, ColNomeArq)
    
    Call Preenche_DOC(arq, texto, TextoNovo, ArqNovo, TipoArq)
    
    li = li + 1

    
    If li <= lf And roboStart Then
        Application.OnTime Now + TimeValue(tempo_espera), "Preenche_DOC_loop"
    Else
        word.Quit
        MsgBox "Arquivo Gerado"
        Call abrir_pasta(CamDest)
    End If
    
End Function


Private Function Preenche_DOC(arq, texto, TextoNovo, ArqNovo, TipoArq)
    On Error GoTo erro
    
    Set doc = word.Documents.Open(cam & "\" & arq, ReadOnly:=True)
    
        doc.Windows(1).Visible = TipoArq = "Aberto"
    
    For c = 0 To UBound(colarr)
        Call FindRangeWord(doc, texto(c), TextoNovo(c))
    Next
    
    Call salvar_Arq(TipoArq, ArqNovo)
    
    Plan2.Cells(li, "T").Value = "Gerado"
    
Exit Function
erro: doc.Saved = True: doc.Close: Plan2.Cells(li, "T").Value = "Erro"
End Function


Private Function FindRangeWord(doc, texto, TextoNovo)
    On Error GoTo vazio
    
    For d = 1 To 30
        testebool = doc.Content.Find.Execute(FindText:=texto, Forward:=True, ReplaceWith:=TextoNovo)
        If testebool = False Then Exit For
    Next

Exit Function
vazio:  doc.Content.Find.Execute FindText:=texto, Forward:=True, ReplaceWith:=""
End Function


Private Function salvar_Arq(TipoArq, ArqNovo)

    If TipoArq = "PDF" Then
        Call Salvar_PDF3(ArqNovo)
    ElseIf TipoArq = "Word" Then
        doc.SaveAs (CamDest & "\" & ArqNovo & ".doc")
    End If
    
    If TipoArq <> "Aberto" Then
        doc.Saved = True
        doc.Close
    End If
    
End Function

Private Function Salvar_PDF3(NomeArq)
    CamSalvar = CamDest & "\" & NomeArq & ".pdf"
    
    word.Documents(word.ActiveDocument).ExportAsFixedFormat OutputFileName:=CamSalvar _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        
End Function


Private Function colConvert(colcoverter)
    
    'Dim col(0 To 26) As String
    col = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    ReDim ret(UBound(colcoverter))
    
    For i = 0 To UBound(colcoverter)
        For f = 0 To UBound(col)
            If colcoverter(i) = col(f) Then ret(i) = f
        Next
    Next
    
    colConvert = ret
    
End Function


Private Function gerar()


    ti = "(" & """"
    
    For i = 1 To 26
        ti = ti & """,""" & Chr(i + 64)
    Next
    
    ti = ti & ")"
    
    Selection.Value = ti
    

End Function




