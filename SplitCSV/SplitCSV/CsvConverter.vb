Imports ADSFramework.Misc
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Text

Public Class CsvConverter
#Region "ToDataTable"

    ''' <summary>
    ''' convert a csv file into a Generic datatable
    ''' </summary>
    ''' <param name="csv">the file path to the csv file</param>
    ''' <param name="delimiter">the separator used to separate data in the file</param>
    ''' <param name="HeaderRowIndex">Specify the Header Row in the given File by it's 0-based index.  The Header Row will be used for Column Names in the resultant DataTable.  Use -1 to select the first 'Full' row of data as the Header Row, -2 to use no header row</param>
    ''' <returns></returns>
    Public Shared Function ToDatatable(ByVal csv As String, Optional ByVal delimiter As Char = ","c, Optional ByVal HeaderRowIndex As Integer = -1) As DataTable
        Dim useless As String = ""
        Return ToDatatable(csv, delimiter, HeaderRowIndex, useless)
    End Function

    ''' <summary>
    ''' convert a csv file into a Generic datatable
    ''' </summary>
    ''' <param name="csv">the file path to the csv file</param>
    ''' <param name="delimiter">the separator used to separate data in the file</param>
    ''' <param name="HeaderRowIndex">Specify the Header Row in the given File by it's 0-based index.  The Header Row will be used for Column Names in the resultant DataTable.  Use -1 to select the first 'Full' row of data as the Header Row</param>
    ''' <returns></returns>
    Public Shared Function ToDatatable(Of T As {Data.DataTable, New})(ByVal csv As String, Optional ByVal delimiter As Char = ","c, Optional ByVal HeaderRowIndex As Integer = -1) As T
        Return ToTyped(Of T)(ToDatatable(csv, delimiter, HeaderRowIndex))
    End Function

    ''' <summary>
    ''' convert a csv file into a Generic datatable
    ''' </summary>
    ''' <param name="csv">the file path to the csv file</param>
    ''' <param name="delimiter">the separator used to separate data in the file</param>
    ''' <param name="HeaderRowIndex">Specify the Header Row in the given File by it's 0-based index.  The Header Row will be used for Column Names in the resultant DataTable.  Use -1 to select the first 'Full' row of data as the Header Row</param>
    ''' <param name="result">Result of the Operation.  Use for advanced Debugging.</param>
    ''' <returns></returns>
    Public Shared Function ToDatatable(ByVal csv As String, ByVal delimiter As Char, ByVal HeaderRowIndex As Integer, ByRef result As String) As DataTable

        '' validate csv file path exists
        If (File.Exists(csv) = False) Then
            result = "Csv file path could not be verified or does not exist"
            Return New DataTable
        End If

        Try
            Dim tp As TextFieldParser

            '' Get width of csv file (column count)
            tp = New TextFieldParser(csv) : With tp
                .TextFieldType = FieldType.Delimited
                .Delimiters = {delimiter}
            End With
            Dim cc As Integer = 0 '' col count
            While Not (tp.EndOfData)
                Dim tpcc As Integer = tp.ReadFields.Length - 1
                If (tpcc > cc) Then cc = tpcc
            End While

            '' Init datatable
            Dim dt = New DataTable
            For i As Integer = 0 To cc
                dt.Columns.Add(i, GetType(String))
            Next

            '' Copy data to datatable
            tp = New TextFieldParser(csv) : With tp
                .TextFieldType = FieldType.Delimited
                .Delimiters = {delimiter}
            End With
            While Not (tp.EndOfData)
                Dim cr() As String = tp.ReadFields '' Current Row
                cr = UnifyCol(cr, cc)
                dt.Rows.Add(cr)
            End While

            '' Begin Format DataTable
            '' Find Header row
            'Dim j As Integer = 0
            'While j < dt.Rows.Count
            '    If Not (isRowFull(dt.Rows(0))) Then
            '        dt.Rows.RemoveAt(0)
            '    Else
            '        Exit While
            '    End If
            '    j += 1
            'End While
            If (HeaderRowIndex <> -2) Then
                Dim j As Integer = 0
                While ((dt.Rows.Count > 0) AndAlso (((HeaderRowIndex > -1) AndAlso (j <> HeaderRowIndex)) OrElse ((HeaderRowIndex = -1) AndAlso (IsRowFull(dt.Rows(0)) = False))))
                    dt.Rows.RemoveAt(0)
                    j += 1
                End While

                '' Move Headers to DataTable ColumnNames
                For k As Integer = 0 To dt.Columns.Count - 1
                    If (dt.Rows(0).Item(k).ToString.Length = 0) Then
                        dt.Rows(0).Item(k) = $"DataColumn{k + 1}"
                    End If
                    dt.Columns(k).ColumnName = dt.Rows(0).Item(k)
                Next
                dt.Rows.RemoveAt(0)
            End If

            '' Remove Empty Rows
            dt = RmEmptyRows(dt)

            tp.Dispose()
            Return dt
        Catch ex As Exception
            result = ex.Message
            Return New DataTable
        End Try
    End Function


    ''' <summary>
    ''' convert a csv file into a Generic datatable
    ''' </summary>
    ''' <typeparam name="T">a Typed Datatable generic</typeparam>
    ''' <param name="csv">the file path to the csv file</param>
    ''' <param name="delimiter">the separator used to separate data in the file</param>
    ''' <param name="HeaderRowIndex">Specify the Header Row in the given File by it's 0-based index.  The Header Row will be used for Column Names in the resultant DataTable.  Use -1 to select the first 'Full' row of data as the Header Row</param>
    ''' <param name="result">Result of the Operation.  Use for advanced Debugging.</param>
    ''' <returns></returns>
    Public Shared Function ToDatatable(Of T As {Data.DataTable, New})(ByVal csv As String, ByVal delimiter As Char, ByVal HeaderRowIndex As Integer, ByRef result As String) As T
        Return ToTyped(Of T)(ToDatatable(csv, delimiter, HeaderRowIndex, result))
    End Function

    ''' <summary>
    ''' Standardize Row Array Length to eliminate possibility of jagged Rows
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="c"></param>
    ''' <param name="setTo"></param>
    ''' <returns></returns>
    Private Shared Function UnifyCol(ByVal r() As String, c As Integer, Optional ByVal setTo As String = "") As String()

        While r.Length < c
            ReDim Preserve r(r.Length)
            r(r.Length - 1) = setTo
        End While

        Return r
    End Function

    ''' <summary>
    ''' Validate if All items in Row Array Contain a non-Empty String
    ''' </summary>
    ''' <param name="r"></param>
    ''' <returns></returns>
    Private Shared Function IsRowFull(ByVal r As DataRow) As Boolean
        For Each i As String In r.ItemArray
            If (IsNothing(i)) Then
                Return False
            End If
            If (Trim(i) = "") Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' Validate if All items in Row Array contain a empty String
    ''' </summary>
    ''' <param name="r"></param>
    ''' <returns></returns>
    Private Shared Function IsRowEmpty(ByVal r As DataRow) As Boolean
        For Each i As String In r.ItemArray
            If Not (Trim(i) = "") Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' for each row, Remove all rows where isRowEmpty() = true
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    Private Shared Function RmEmptyRows(ByVal dt As DataTable) As DataTable
        Dim j As Integer = 0
        While j < dt.Rows.Count
            If (IsRowEmpty(dt.Rows(j))) Then
                dt.Rows.RemoveAt(j)
            Else
                j += 1
            End If
        End While
        Return dt
    End Function

#End Region

#Region "ToCsv"
    ''' <summary>
    ''' ADSFramework: Convert a Datatable object to string of Csv
    ''' </summary>
    ''' <param name="dt">the datatable to convert to csv</param>
    ''' <returns></returns>
    Public Shared Function ToCsv(ByVal dt As DataTable) As String
        Return ToCsv(dt, ","c)
    End Function

    ''' <summary>
    ''' ADSFramework: Convert a Datatable object to string of Csv
    ''' </summary>
    ''' <param name="dt">the datatable to convert to csv</param>
    ''' <param name="Delimiter">the delimiter to use for the csv</param>
    ''' <returns></returns>
    Public Shared Function ToCsv(ByVal dt As DataTable, ByVal Delimiter As Char) As String
        '' Add ColumnNames
        Dim sb As New StringBuilder()
        For i As Integer = 0 To dt.Columns.Count - 1
            sb.Append(dt.Columns(i))
            If (i < dt.Columns.Count - 1) Then
                sb.Append(Delimiter)
            End If
        Next
        sb.AppendLine()

        '' Add DataRows
        For Each r As DataRow In dt.Rows
            For i As Integer = 0 To dt.Columns.Count - 1
                If ((IsDBNull(r(i)) = False) AndAlso (r(i).ToString.Contains(Delimiter.ToString))) Then
                    '' must escape delimiter if exists in string
                    '' encapsulate data in quotes
                    sb.Append("""" & r(i).ToString() & """")
                Else
                    sb.Append(r(i).ToString())
                End If

                If (i < dt.Columns.Count - 1) Then
                    sb.Append(Delimiter)
                End If
            Next
            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' ADSFramework: Convert a Datatable object to Csv file.
    ''' </summary>
    ''' <param name="dt">the datatable to convert to csv</param>
    ''' <param name="Delimiter">the delimiter to use for the csv</param>
    ''' <param name="FullPath">the directory to write the file to</param>
    ''' <returns>boolean result of operation</returns>
    Public Shared Function ToCsv(ByVal dt As DataTable, ByVal Delimiter As Char, ByVal FullPath As String) As Boolean
        Return ToCsv(dt, Delimiter, FullPath, True)
    End Function

    ''' <summary>
    ''' ADSFramework: Convert a Datatable object to Csv file.
    ''' </summary>
    ''' <param name="dt">the datatable to convert to csv</param>
    ''' <param name="Delimiter">the delimiter to use for the csv</param>
    ''' <param name="FullPath">the directory to write the file to</param>
    ''' <returns>boolean result of operation</returns>
    Public Shared Function ToCsv(ByVal dt As DataTable, ByVal Delimiter As Char, ByVal FullPath As String, ByVal UseColumnHeaders As Boolean) As Boolean
        Dim useless As String = ""
        Return ToCsv(dt, Delimiter, FullPath, UseColumnHeaders, useless)
    End Function

    ''' <summary>
    ''' ADSFramework: Convert a Datatable object to Csv file.
    ''' </summary>
    ''' <param name="dt">the datatable to convert to csv</param>
    ''' <param name="Delimiter">the delimiter to use for the csv</param>
    ''' <param name="FullPath">the directory to write the file to</param>
    ''' <param name="result">Result of the Operation.  Use for advanced Debugging.</param>
    ''' <returns></returns>
    Public Shared Function ToCsv(ByVal dt As DataTable, ByVal Delimiter As Char, ByVal FullPath As String, ByVal UseColumnHeaders As Boolean, ByRef result As String) As Boolean
        ToCsv = False
        result = ""

        '' Make sure is csv file
        If (FullPath.SafeSubstring(FullPath.Length - 4, 4) <> ".csv") Then FullPath += ".csv"

        Using sw As StreamWriter = New StreamWriter(FullPath, True, Encoding.UTF8, 4096)
            Dim sb As New StringBuilder

            Try
                '' Add ColumnNames
                If (UseColumnHeaders) Then
                    For i As Integer = 0 To dt.Columns.Count - 1
                        sb.Append(dt.Columns(i))
                        If (i < dt.Columns.Count - 1) Then
                            sb.Append(Delimiter)
                        End If
                    Next
                    sw.WriteLine(sb) : sw.Flush()
                    sb.Clear()
                End If

                '' Add DataRows
                For Each r As DataRow In dt.Rows
                    For i As Integer = 0 To dt.Columns.Count - 1
                        If ((IsDBNull(r(i)) = False) AndAlso (r(i).ToString.Contains(Delimiter.ToString))) Then
                            '' must escape delimiter if exists in string, encapsulate in quotes
                            sb.Append("""" & r(i).ToString() & """")
                        Else
                            sb.Append(r(i).ToString())
                        End If

                        If (i < dt.Columns.Count - 1) Then
                            sb.Append(Delimiter)
                        End If
                    Next
                    sw.WriteLine(sb) : sw.Flush()
                    sb.Clear()
                Next

                ToCsv = True
            Catch ex As Exception
                result = ex.Message
            Finally
                sb = Nothing
            End Try
        End Using
    End Function
#End Region
End Class
