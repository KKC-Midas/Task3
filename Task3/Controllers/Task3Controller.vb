Imports System.Net
Imports System.Net.Http
Imports System.Web.Http
Imports System.Web.Http.Cors
Imports OfficeOpenXml
Imports System.IO
Imports OfficeOpenXml.Style

<EnableCors("*", "*", "*")>
Public Class Task3Controller
    Inherits ApiController

    Private Shared logFilePath As String = HttpContext.Current.Server.MapPath("~/App_Data/Task3Controller" & Format(Now, "yyyyMMdd") & ".log")

    ' Helper method to log messages
    Private Sub LogMessage(message As String)
        Try
            Using writer As New StreamWriter(logFilePath, True)
                writer.WriteLine($"{DateTime.Now}: {message}")
            End Using
        Catch ex As Exception
            Console.WriteLine("Exception From LogMessage: " + ex.Message)
        End Try
    End Sub

    ' POST api/Task3
    <HttpPost>
    Public Function GetNameRanges() As HttpResponseMessage
        LogMessage("GetNameRanges() Triggered")
        Try
            Dim httpRequest = HttpContext.Current.Request
            Dim baseFile = HttpContext.Current.Server.MapPath("~/App_Data/Copy of PSC_AASHTO_LRFD_Report.xlsx")
            Dim outputFile = httpRequest.Files("Output File")
            Dim missingNameRanges As New List(Of String)
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial

            ' Load the base Excel file
            Using package As New ExcelPackage(New FileInfo(baseFile))
                Dim workbook = package.Workbook
                Dim actualNames = workbook.Names.Where(Function(n) n.Address.Contains("Detail")).Select(Function(n) n.Name).ToList()
                Using newPackage As New ExcelPackage(outputFile.InputStream)
                    Dim newWorkbook = newPackage.Workbook
                    For Each ws In newWorkbook.Worksheets
                        If Not (ws.Name.Contains("_I") Or ws.Name.Contains("_J")) Then
                            Continue For
                        End If
                        Dim nameRange = workbook.Names.Where(Function(n) n.Address.Contains("Detail"))
                        For Each name As ExcelNamedRange In nameRange
                            If Not (IsSubArray(ws, name)) Then
                                missingNameRanges.Add(name.Name)
                            End If
                        Next
                    Next
                    Dim difference As List(Of String) = actualNames.Except(missingNameRanges).ToList()
                    LogMessage("NamedRanges in Base File Which are also present in Output File: ")
                    For Each name As String In difference
                        LogMessage(name)
                    Next
                    LogMessage("NamedRanges Which are Missing in Output File: ")
                    For Each name As String In missingNameRanges
                        LogMessage(name)
                    Next
                    LogMessage("GetNameRanges() Executed successfully")
                    Return Request.CreateResponse(HttpStatusCode.OK, missingNameRanges)
                End Using
            End Using

        Catch ex As Exception
            LogMessage("GetNameRanges() Thrown Exception: " + ex.Message)
            Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message)
        End Try
    End Function

    'Helper method To compare cell formats
    Private Function CompareCellFormats(baseFormat As ExcelStyle, outputFormat As ExcelStyle) As Boolean
        Return baseFormat.Font.Bold = outputFormat.Font.Bold AndAlso
               baseFormat.Font.Color.Rgb = outputFormat.Font.Color.Rgb AndAlso
               baseFormat.Fill.BackgroundColor.Rgb = outputFormat.Fill.BackgroundColor.Rgb AndAlso
               baseFormat.Border.Top.Style = outputFormat.Border.Top.Style AndAlso
               baseFormat.Border.Bottom.Style = outputFormat.Border.Bottom.Style AndAlso
               baseFormat.Border.Left.Style = outputFormat.Border.Left.Style AndAlso
               baseFormat.Border.Right.Style = outputFormat.Border.Right.Style
    End Function

    'Helper method To check whether the namedRange is a sub-set of the output file
    Private Function IsSubArray(mainArray As ExcelWorksheet, subArray As ExcelNamedRange) As Boolean
        Dim mainRows As Integer = mainArray.Dimension.Rows
        Dim mainCols As Integer = mainArray.Dimension.Columns
        Dim subRows As Integer = subArray.Rows
        Dim subCols As Integer = subArray.Columns

        For i As Integer = 1 To mainRows - subRows + 1

            For j As Integer = 1 To mainCols

                Dim found As Boolean = True


                For k As Integer = subArray.Start.Row To subArray.Start.Row + subRows - 1
                    For l As Integer = subArray.Start.Column To subArray.Start.Column + subCols - 1
                        If subArray.Worksheet.Cells(k, l).Value Is Nothing Then
                            Continue For
                        End If
                        If subArray.Worksheet.Cells(k, l).Value.ToString = "0" Then
                            Continue For
                        End If
                        If (mainArray.Cells(i + k - subArray.Start.Row, j + l - subArray.Start.Column).Value Is Nothing) AndAlso (subArray.Worksheet.Cells(k, l).Value IsNot Nothing) Then
                            'Continue For
                            found = False
                            Exit For
                        End If
                        If (mainArray.Cells(i + k - subArray.Start.Row, j + l - subArray.Start.Column).Value Is Nothing) Then
                            Continue For
                            'found = False
                            'Exit For
                        End If
                        If (mainArray.Cells(i + k - subArray.Start.Row, j + l - subArray.Start.Column).Value.GetType.Name <> subArray.Worksheet.Cells(k, l).Value.GetType.Name) Then
                            found = False
                            Exit For
                        End If
                        If (mainArray.Cells(i + k - subArray.Start.Row, j + l - subArray.Start.Column).Value <> subArray.Worksheet.Cells(k, l).Value) Then
                            found = False
                            Exit For
                        End If
                        If Not CompareCellFormats(mainArray.Cells(i + k - subArray.Start.Row, j + l - subArray.Start.Column).Style, subArray.Worksheet.Cells(k, l).Style) Then
                            found = False
                            Exit For
                        End If
                    Next
                    If Not found Then Exit For
                Next

                ' If all elements match, return True
                If found Then
                    Return True
                End If
                If Not found Then
                    Exit For
                End If
            Next
        Next

        ' No matching subArray found
        Return False
    End Function
End Class
