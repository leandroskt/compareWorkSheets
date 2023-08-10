Sub CompareWorksheets()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim r1 As Range, r2 As Range
    Dim cell1 As Range, cell2 As Range
    Dim diffCount As Long
    Dim found As Range
    Dim Response
    Dim collection As New collection
    
    'Call ConsolidateData
    
    
    ' Defina aqui as planilhas que você deseja comparar
    Set ws1 = ThisWorkbook.Worksheets(2)
    Set ws2 = ThisWorkbook.Worksheets(3)

    ' Defina os intervalos que você deseja comparar nas planilhas
    Set r1 = ws1.UsedRange
    Set r2 = ws2.UsedRange

    ' Verifique se os intervalos têm o mesmo tamanho
    If r1.Rows.Count <> r2.Rows.Count And r1.Columns.Count <> r2.Columns.Count Then
        MsgBox "O intervalo de linhas e colunas são diferentes!", vbExclamation
        Exit Sub
        
    ElseIf r1.Rows.Count <> r2.Rows.Count Then
        Dim rowDiff As New collection
        diffCount = 0
        For Each r1 In ws1.UsedRange.Rows
            ' Verificar se a linha existe em ws2
            Set found = Nothing
            For Each r2 In ws2.UsedRange.Rows
                  
                If r1.Value2(1, 1) = r2.Value2(1, 1) Then
                     Set found = r1
                     Exit For
                End If
        
                If Not found Is Nothing Then
                    Exit For
                End If
                
            Next r2
    
            ' Se a linha não foi encontrada, destaque-a
            If found Is Nothing Then
                diffCount = diffCount + 1
                collection.Add r1.Row
                rowDiff.Add r1
                For Each cell In r1.Cells
                    cell.Interior.Color = RGB(250, 150, 150) ' Cor em vermelho claro
                Next cell
            End If
            
        Next r1
        
        For Each Item In collection
            Message = Message & "Linha: " & Item & vbNewLine
        Next Item
        
        Response = MsgBox("Deseja Apagar as linhas:" & vbNewLine & Message & "E executar compare novamente?", vbYesNo, "Ajuste no compare")

        If Response = vbYes Then
            For Each linha In rowDiff
                linha.Delete
            Next
            Call CompareWorksheets
            Exit Sub
        End If
        
        MsgBox "A planilha possui " & diffCount & " Linhas diferentes!", vbExclamation
        
        Exit Sub
        
    ElseIf r1.Columns.Count <> r2.Columns.Count Then
        ' Loop através de cada coluna em ws1
        Dim c1 As Range
        Dim c2 As Range
        Dim columnDiff As New collection

        diffCount = 0

        For Each c1 In ws1.UsedRange.Columns

            ' Verifique se a coluna existe em ws2
            Set found = Nothing
            For Each c2 In ws2.UsedRange.Columns
                 
                If c1.Value2(1, 1) = c2.Value2(1, 1) Then
                     Set found = c1
                     Exit For
                End If
        
                If Not found Is Nothing Then
                    Exit For
                End If
            Next c2
    
            ' Se a coluna não foi encontrada, destaque-a
            If found Is Nothing Then
                columnDiff.Add c1
                collection.Add c1.Column
    
                diffCount = diffCount + 1
                For Each cell In c1.Cells
                    cell.Interior.Color = RGB(250, 150, 150) ' Cor em vermelho claro
                Next cell
            End If
            
        Next c1
        
        For Each Item In collection
            Message = Message & "Coluna: " & Item & vbNewLine
        Next Item
        
        Response = MsgBox("Deseja Apagar as colunas:" & vbNewLine & Message & "E executar compare novamente?", vbYesNo, "Ajuste no compare")
        
        If Response = vbYes Then
            For Each coluna In columnDiff
                coluna.Delete
            Next
            Call CompareWorksheets
            Exit Sub
        End If
        
        MsgBox "A planilha possui " & diffCount & " Colunas diferentes!", vbExclamation

        Exit Sub
        
    End If


    ' Compare as células
    diffCount = 0
    For Each cell1 In r1
        Set cell2 = r2.Cells(cell1.Row, cell1.Column)
        If cell1.Value <> cell2.Value Then
            ' Destaque as células diferentes
            cell1.Interior.Color = RGB(250, 150, 150)
            cell2.Interior.Color = RGB(250, 150, 150)
            diffCount = diffCount + 1
        End If
    Next cell1

    ' Exiba o resultado da comparação
    If diffCount > 0 Then
        MsgBox diffCount & " células são diferentes entre as planilhas!", vbInformation
    Else
        MsgBox "As planilhas são idênticas!", vbInformation
    End If
        
        
    ThisWorkbook.Worksheets(3).Activate
    ThisWorkbook.Worksheets(3).Range("A1").Select
    ThisWorkbook.Worksheets(2).Activate
    ThisWorkbook.Worksheets(2).Range("A1").Select
    
End Sub

Sub ClearAllSheets(Optional aviso As Variant)
    
    If IsMissing(aviso) Then
        aviso = True
    End If
    
    If aviso Then
        Dim Response
        
        Response = MsgBox("Deseja limpar e apagar tudo?", vbYesNo, "Script de Limpeza")
        
        If Response = vbNo Then
            Exit Sub
        End If
    End If


    
    Dim ws As Worksheet
    Dim sheetToKeep As String

    sheetToKeep = ThisWorkbook.Worksheets(1).Name
    
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetToKeep Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True
    
    ThisWorkbook.Worksheets(1).Range("A1").Select
    
    ' Loop through all worksheets in the workbook
    'For Each ws In ThisWorkbook.Worksheets
        ' Clear everything (contents, formats, comments, etc.)
    '    ws.Activate
    '    ws.Cells.Clear
    '    ws.Range("A1").Select
   ' Next ws
End Sub

Function IsValidFileType(fileName As Variant, fileExtension As String) As Boolean
    If Right(fileName, Len(fileExtension)) = fileExtension Then
        IsValidFileType = True
    Else
        IsValidFileType = False
    End If
End Function

Sub Import2Files()
    Dim fileNames As Variant
    Dim wb As Workbook, wb1 As Workbook
    Dim ws1 As Worksheet
    Dim tipoXls As String, tabName As String
    Dim isXlsFile As Boolean
    

    
    ' Prompt user to select multiple files
    fileNames = Application.GetOpenFilename(Title:="Selecione os arquivos", MultiSelect:=True)
    
    If Not IsArray(fileNames) Then
        MsgBox "Erro: Nenhum arquivo selecionado. Saindo..."
        Exit Sub
    ElseIf UBound(fileNames) <> 2 Then
        MsgBox "Erro: Selecione exatamente 2 arquivos para comparação!"
        Exit Sub
    End If

    ' Set the current workbook
    Set wb = ThisWorkbook

    ' Clear all sheets
    Call ClearAllSheets(False)

    ' Loop through the selected files and import data
    Dim fileName As Variant
    
    Dim i As Integer
    i = 1
    
    For Each fileName In fileNames
        If IsValidFileType(fileName, "xls") Then
            Set wb1 = Workbooks.Open(fileName)
            ' Import the sheet from the current workbook
            wb1.Sheets(1).Copy after:=wb.Worksheets(wb.Worksheets.Count())
            wb1.Close SaveChanges:=False
        ElseIf IsValidFileType(fileName, "xlsx") Then
            Set wb1 = Workbooks.Open(fileName, corruptLoad:=xlExtractData) 'xlRepairFile)
            ' Import the sheet from the current workbook
            wb1.Sheets(1).Copy after:=wb.Worksheets(wb.Worksheets.Count())
            wb1.Close SaveChanges:=False
        
        'Tentativa de correção para arquivos com enconding inválido em csv
        ElseIf IsValidFileType(fileName, "csv") Then
            
            ' Cria uma nova planilha
            Set ws1 = ThisWorkbook.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
        
            Dim fileNum As Integer
            Dim firstLine As String
            Dim delimiter As String
            
            fileNum = FreeFile
            Open fileName For Input As #fileNum
            Line Input #fileNum, firstLine
            Close #fileNum
            
             If InStr(firstLine, ",") > 0 Then
                delimiter = ","
            ElseIf InStr(firstLine, ";") > 0 Then
                delimiter = ";"
            Else
                ' Default to comma if no comma or semicolon found
                delimiter = ","
            End If
            
            
            ' Importa os dados do arquivo CSV para a nova planilha
            With ws1.QueryTables.Add(Connection:="TEXT;" & fileName, Destination:=ws1.Range("A1"))
                .Name = "Import"
                .FieldNames = True
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .TextFilePromptOnRefresh = False
                .TextFilePlatform = 65001
                .TextFileStartRow = 1
                .TextFileParseType = xlDelimited
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileConsecutiveDelimiter = False
                .TextFileTabDelimiter = False
                .TextFileSemicolonDelimiter = (delimiter = ";")
                .TextFileCommaDelimiter = (delimiter = ",")
                .TextFileSpaceDelimiter = False
                .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1)
                .TextFileTrailingMinusNumbers = True
                .Refresh BackgroundQuery:=False
            End With
        Else
            Set wb1 = Workbooks.Open(fileName, corruptLoad:=xlExtractData)
            ' Import the sheet from the current workbook
            wb1.Sheets(1).Copy after:=wb.Worksheets(wb.Worksheets.Count())
            wb1.Close SaveChanges:=False
        End If
        


        ' Name the imported sheets
        If (Len(Dir(fileName)) < 31) Then
            ThisWorkbook.Sheets(i + 1).Name = Dir(fileName)
        Else
            ThisWorkbook.Sheets(i + 1).Name = Right(Dir(fileName), 30)
        End If

        i = i + 1
        
    Next fileName

    Call FilterAllWorkSheets
    
End Sub
Function AddQueryAndTransformData()
   
   
    ' Adicionar a consulta (Query)
    
        
    ' Adicionar uma nova planilha (Worksheet)
    Set objWorksheet = objWorkbook.Worksheets.Add
    
    ' Salvar e fechar o arquivo
    objWorkbook.Save
    objWorkbook.Close
    
    ' Fechar a aplicação Excel
    objExcel.Quit
    
    ' Liberar memória dos objetos
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Function



Sub FilterAllWorkSheets()
    Dim ws As Worksheet
    'Add filter to all worksheets except the First one
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode And ws.Index <> 1 Then
            ws.AutoFilterMode = False
        End If
        If ws.Index <> 1 Then
            ws.Rows(1).AutoFilter
        End If
        ws.Activate
        ws.Range("A1").Select
    Next ws
End Sub



Sub SalvarWorkbook()
    Dim nomeArquivo As String
    Dim pastaFixa As String
    Dim pasta As Object
    Dim fs As Object
    Dim arquivoSalvo As String

    ' Solicitar o nome do arquivo ao usuário
    nomeArquivo = InputBox("Digite o nome do Relatório:", "Salvar Workbook")

    ' Verificar se o usuário digitou um nome de arquivo válido
    If nomeArquivo = "" Then
        MsgBox "Nome do arquivo inválido. Operação cancelada."
        Exit Sub
    End If

    ' Definir a pasta fixa onde o arquivo será salvo
    pastaFixa = "G:\Meu Drive\Documentos Leandro\Evidencias\Empresarial\Adaptação dos relatórios para spreadsheet\"

    ' Verificar se a pasta é acessível
    Set fs = CreateObject("Scripting.FileSystemObject")

    If Not fs.FolderExists(pastaFixa) Then
        MsgBox "A pasta fixa não existe ou não é acessível. Operação cancelada."
        Exit Sub
    End If

    ' Verificar se já existe um arquivo com o mesmo nome
    If fs.FileExists(pastaFixa & nomeArquivo & " compare.xlsm") Then
        MsgBox "Já existe um arquivo com o mesmo nome na pasta fixa. Operação cancelada."
        Exit Sub
    End If


    ' Salvar o workbook na pasta fixa com o nome fornecido pelo usuário
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=pastaFixa & nomeArquivo & " compare.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    arquivoSalvo = ActiveWorkbook.Name()
    
    Shell "explorer.exe /select, /n, /e," & pastaFixa & arquivoSalvo, vbNormalFocus

    'MsgBox "Workbook salvo com sucesso em: " & pastaFixa & arquivoSalvo, vbInformation
    
End Sub

Sub ConsolidateData()
       Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cols As Variant
    Dim i As Variant

    ' Define the columns to edit
    cols = Array("C")  ', "N", "O")

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Skip the first worksheet
        ws.Activate
        If ws.Index > 2 Then
            ' Loop through specified columns
            For Each i In cols
                ' Loop through each cell in the column
                For Each cell In ws.Columns(i).Cells
                    If cell.Row > ws.UsedRange.Rows.Count Then
                        Exit For
                    
                    End If
                    
                    cell.Select
                    cell.Value = cell.Value
                
                Next cell
            Next i
        End If
    Next ws
End Sub
