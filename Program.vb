Imports System
Imports System.IO
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Module Program
    Sub Main(ByVal args() As String)
        Console.Clear()
        Console.WriteLine("=======================================================")
        Console.WriteLine("      Conversao de planilha excel (XLSX) para CSV      ")
        Console.WriteLine("=======================================================")
        Console.WriteLine("")

        ' Verifique se o número correto de argumentos foi fornecido
        If args.Length <= 0 Then
            Console.WriteLine("Uso: converta-xlsx-csv.exe <diretorio + arquivo> <destino>")
            Console.WriteLine("Exmplo: converta-xlsx-csv.exe c:\pasta\arquivo.xlsx c:\pasta\destino\")
            Console.ReadKey()
            Return
        End If

        ' Caminho do arquivo XLSX a ser convertido
        Dim caminhoDoArquivoXlsx As String = args(0)
        Console.WriteLine("Planilha: " & Path.GetFileName(caminhoDoArquivoXlsx))

        ' Verifique se o arquivo XLSX existe
        If Not File.Exists(caminhoDoArquivoXlsx) Then
            Console.WriteLine("Erro....: O arquivo XLSX não foi encontrado.")
            Return
        End If

        Dim Destino As String = Path.GetDirectoryName(caminhoDoArquivoXlsx)

        If args.Length = 2 Then
            If Directory.Exists(args(1)) Then
                Destino = args(1)
            Else
                Directory.CreateDirectory(args(1))
                Destino = args(1)
                Console.WriteLine("Aviso...: Destino não exite. Criando destino!")
            End If
        End If

        Console.WriteLine("Destino.: " & Destino)

        ' Crie uma instância do Excel e abra o arquivo XLSX
        Dim excelApp As Application = Nothing
        Dim excelWorkbook As Workbook = Nothing

        Try
            excelApp = New Application()
            excelWorkbook = excelApp.Workbooks.Open(caminhoDoArquivoXlsx)

            For Each excelWorksheet As Worksheet In excelWorkbook.Sheets
                Dim usedRange As Microsoft.Office.Interop.Excel.Range = excelWorksheet.UsedRange
                Dim nomeDaPlanilha As String = excelWorksheet.Name
                Console.WriteLine("")
                Console.WriteLine("Aba.....: " & nomeDaPlanilha)

                ' Crie um StreamWriter para escrever no arquivo CSV
                Dim caminhoDoArquivoCsv As String = Path.Combine(Destino, $"{nomeDaPlanilha}.csv") 'Path.Combine(Path.GetDirectoryName(caminhoDoArquivoXlsx), $"{nomeDaPlanilha}.csv")

                'Console.WriteLine("CSV.....: " & caminhoDoArquivoCsv)

                Using writer As New StreamWriter(caminhoDoArquivoCsv)
                    For Each row As Microsoft.Office.Interop.Excel.Range In usedRange.Rows
                        Dim csvLine As String = ""
                        Dim isFirstCell As Boolean = True

                        For Each cell As Microsoft.Office.Interop.Excel.Range In row.Cells
                            If IsDate(cell.Value) Then
                                Dim formattedDate As String = CType(cell.Value, DateTime).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                                csvLine &= If(isFirstCell, "", ";") & formattedDate
                            Else
                                Dim cellValue As String = cell.Value.ToString().Replace(",", ".")
                                csvLine &= If(isFirstCell, "", ";") & cellValue
                            End If

                            isFirstCell = False
                        Next

                        writer.WriteLine(csvLine)
                    Next
                End Using

                Console.WriteLine("Aviso...: Conversão concluída")
                Console.WriteLine("CSV.....: " + $"{caminhoDoArquivoCsv}")
                Console.WriteLine("")
            Next
        Catch ex As Exception
            Console.WriteLine("Erro....: Ocorreu um erro durante a conversão: " & ex.Message)
            Console.WriteLine("")
        Finally
            Console.WriteLine("=======================================================")
            Console.WriteLine("                     ~~~ Fim ~~~                       ")
            Console.WriteLine("=======================================================")
            ' Feche o Excel
            If excelWorkbook IsNot Nothing Then
                excelWorkbook.Close(False)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
            End If
        End Try
    End Sub
End Module
