Sub GerarTXT_CENTRUSRF_Dividido()

    Dim ws As Worksheet
    Dim linha As Long
    Dim dataTexto As Variant, dataFinal As String
    Dim operacao As String, valor As Double
    Dim valorText As String
    Dim banco As String, agencia As String, conta As String
    Dim textoFinal As String
    Dim arquivoResg As String, arquivoTransf As String
    Dim temp As String
    Dim d As Date
    Dim tryText As String
    Dim re As Object, matches As Object
    Dim dayPart As String, monPart As String, yearPart As String
    Dim pastaDestino As String
    Dim dlg As FileDialog
    Dim cleaned As String, i As Long, ch As String
    
    Set ws = Sheets("Teste 01")

    '  LER E TRATAR A DATA

    dataTexto = ws.Range("B31").Value
    
    If IsDate(dataTexto) Then
        d = CDate(dataTexto)
    Else
        tryText = Trim(CStr(dataTexto))
        If InStr(tryText, ".") > 0 Then tryText = Replace(tryText, ".", "/")
        If InStr(tryText, "-") > 0 Then tryText = Replace(tryText, "-", "/")
        
        If IsDate(tryText) Then
            d = CDate(tryText)
        Else
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = "(\d{1,2})[^\d](\d{1,2})[^\d](\d{2,4})"
            re.IgnoreCase = True
            
            If re.Test(CStr(dataTexto)) Then
                Set matches = re.Execute(CStr(dataTexto))
                dayPart = matches(0).SubMatches(0)
                monPart = matches(0).SubMatches(1)
                yearPart = matches(0).SubMatches(2)
                
                If Len(yearPart) = 2 Then
                    If CInt(yearPart) < 50 Then yearPart = "20" & yearPart Else yearPart = "19" & yearPart
                End If
                
                d = DateSerial(CInt(yearPart), CInt(monPart), CInt(dayPart))
            Else
                MsgBox "Não foi possível interpretar a data da célula B31." & vbCrLf & _
                       "Use formato 28.11.2025 ou 28/11/2025.", vbCritical
                Exit Sub
            End If
        End If
    End If

    dataFinal = Format(d, "yyyymmdd")

   
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Selecione a pasta onde deseja salvar os arquivos"
    
    If dlg.Show <> -1 Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    pastaDestino = dlg.SelectedItems(1)

    ' DEFINIR ARQUIVOS SEPARADOS
 
    arquivoResg = pastaDestino & "\CENTRUSRF_RESGATES_" & dataFinal & ".txt"
    arquivoTransf = pastaDestino & "\CENTRUSRF_TRANSFERENCIAS_" & dataFinal & ".txt"

    Open arquivoResg For Output As #1
    Open arquivoTransf For Output As #2

    For linha = 33 To 39

        operacao = CStr(ws.Range("B" & linha).Value)
        valorText = CStr(ws.Range("C" & linha).Value)
        valor = 0
        
        If Trim(operacao) <> "" Then

            valorText = Trim(valorText)
            valorText = Replace(valorText, Chr(160), "")
            valorText = Replace(valorText, "R$", "")
            valorText = Replace(valorText, "$", "")

            If InStr(valorText, ",") > 0 And InStr(valorText, ".") > 0 Then
                valorText = Replace(valorText, ".", "")
                valorText = Replace(valorText, ",", ".")
            ElseIf InStr(valorText, ",") > 0 Then
                valorText = Replace(valorText, ",", ".")
            End If

            cleaned = ""
            For i = 1 To Len(valorText)
                ch = Mid(valorText, i, 1)
                If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then cleaned = cleaned & ch
                Next i
            valorText = cleaned

            If valorText <> "" And IsNumeric(valorText) Then
                valor = CDbl(valorText)
            Else
                GoTo ProximaLinha
            End If

            ' IDENTIFICAR BANCO
    
            banco = ""
            If InStr(1, operacao, "Bradesco", vbTextCompare) > 0 Then banco = "BRADESCO"
            If InStr(1, operacao, "Safra", vbTextCompare) > 0 Then banco = "SAFRA"
            If InStr(1, operacao, "Itaú", vbTextCompare) > 0 Then banco = "ITAU"
            If InStr(1, operacao, "Brasil", vbTextCompare) > 0 Then banco = "BANCO DO BRASIL"

            ' AGENCIA E CONTA

            agencia = ""
            conta = ""

            If InStr(1, operacao, "agência:", vbTextCompare) > 0 Then
                temp = Mid(operacao, InStr(1, operacao, "agência:", vbTextCompare) + 8)
                agencia = Trim(Split(temp, ")")(0))
            End If
            
            If InStr(1, operacao, "conta:", vbTextCompare) > 0 Then
                temp = Mid(operacao, InStr(1, operacao, "conta:", vbTextCompare) + 6)
                conta = Trim(Split(temp, " ")(0))
            End If

    
            If InStr(1, operacao, "Resgate", vbTextCompare) > 0 Then
            
                textoFinal = "CENTRUSRF;" & dataFinal & _
                             ";Não Selecionado;V;S;1BRADFOC;" & _
                             "46618087993;02561193000185;" & _
                             banco & " FI RF FOCO;;;0000;0000;0000;;;;" & _
                             banco & ";46618449;N;N;" & _
                             Replace(Format(valor, "0.00"), ",", ".") & _
                             ";0.00000000;0.00000000;0.00;0.00000000;Cetip;Bruta;;;;;;;;;;S;"

                Print #1, textoFinal

        
            ElseIf InStr(1, operacao, "Transferência", vbTextCompare) > 0 Then

                textoFinal = "CENTRUSRF;" & dataFinal & ";N;R;" & _
                             Replace(Format(valor, "0.00"), ",", ".") & _
                             ";Doc;Str;N;;;C;CONTAS BANCÁRIAS;" & _
                             "00580571000142;01;" & _
                             Replace(agencia, "-", "") & ";" & _
                             Replace(conta, "-", "") & ";" & _
                             Replace(Format(valor, "0.00"), ",", ".") & _
                             ";J;C;"

                Print #2, textoFinal

            End If

        End If
ProximaLinha:
    Next linha

    Close #1
    Close #2

    MsgBox "Arquivos gerados:" & vbCrLf & _
           arquivoResg & vbCrLf & arquivoTransf, vbInformation

End Sub


