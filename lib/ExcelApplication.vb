Option Strict Off

Namespace ExcelUtility
    ''' <summary>
    ''' Excelのラッパークラス
    ''' </summary>
    ''' <remarks>
    ''' <example>
    ''' 通常の使用例
    ''' <code>
    ''' Using cEX As New ExcelApplication
    '''     Using cWB As ExcelWorkbook = cEX.Workbooks.Open("C:\Temp\出力エクセル.xls")
    '''         'マクロ呼び出し
    '''         cEX.Run("SetData", P_JigyoID, MDBNAME)
    '''         'セルの値取得
    '''         cWB.Worksheets("WORK").Cells(1, 1).Value = P_JigyoID
    '''         '保存して画面を閉じる
    '''         cWB.Close(True)
    '''     End Using
    ''' End Using
    ''' </code>
    ''' データテーブルの単純な出力
    ''' <code>
    ''' 'DataTableは別途取得済み
    ''' Using cEX As New ExcelApplication
    '''     cEX.ExportFile(dt, "C:\Temp\出力エクセル.xls")
    ''' End Using
    ''' </code>
    ''' </example>
    ''' </remarks>
    Public Class ExcelApplication
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelWorkbooks As ExcelWorkbooks

        'ネイティブリソース
        Private _NativeExcelApplication As Object

        Public Sub New()
            _NativeExcelApplication = CreateObject("Excel.Application")
            _ManagedExcelWorkbooks = New ExcelWorkbooks(_NativeExcelApplication)
        End Sub

        Public Property Visible As Boolean
            Get
                Return _NativeExcelApplication.Visible
            End Get
            Set(value As Boolean)
                _NativeExcelApplication.Visible = value
            End Set
        End Property

        Public ReadOnly Property Caption As String
            Get
                Return _NativeExcelApplication.Caption
            End Get
        End Property

        Public ReadOnly Property Workbooks() As ExcelWorkbooks
            Get
                Return _ManagedExcelWorkbooks
            End Get
        End Property

        Public Sub Run(ByVal Macro As String, ByVal Arg1 As Object)
            Try
                _NativeExcelApplication.Run(Macro, Arg1)
            Catch ex As System.Runtime.InteropServices.COMException
                If ex.Message.Contains("0x800A9C68") Then
                    'マクロ内でブック自体を閉じる処理でエラーが発生することがある。これは無視してもよいので無視する、
                Else
                    Debug.Print(ex.ToString)
                End If
            End Try
        End Sub

        Public Sub Run(ByVal Macro As String, ByVal Arg1 As Object, ByVal Arg2 As Object)
            Try
                _NativeExcelApplication.Run(Macro, Arg1, Arg2)
            Catch ex As System.Runtime.InteropServices.COMException
                If ex.Message.Contains("0x800A9C68") Then
                    'マクロ内でブック自体を閉じる処理でエラーが発生することがある。これは無視してもよいので無視する、
                Else
                    Debug.Print(ex.ToString)
                End If
            End Try
        End Sub

        Public Sub Run(ByVal Macro As String, ByVal Arg1 As Object, ByVal Arg2 As Object, ByVal Arg3 As Object)
            Try
                _NativeExcelApplication.Run(Macro, Arg1, Arg2, Arg3)
            Catch ex As System.Runtime.InteropServices.COMException
                If ex.Message.Contains("0x800A9C68") Then
                    'マクロ内でブック自体を閉じる処理でエラーが発生することがある。これは無視してもよいので無視する、
                Else
                    Debug.Print(ex.ToString)
                End If
            End Try
        End Sub

        ''' <summary>
        ''' 指定したデータテーブルをExcel2003形式でファイルに出力する。
        ''' </summary>
        ''' <param name="TargetData">データテーブル。テーブル名が指定されている場合は、シート名となる</param>
        ''' <param name="FileName">ファイル名</param>
        ''' <param name="RowLimit">（省略可）１ファイルへの出力行数の制限値</param>
        ''' <remarks></remarks>
        Public Sub ExportFile(ByVal TargetData As DataTable, ByVal FileName As String, Optional ByVal RowLimit As Integer = 10000)
            If TargetData.Rows.Count = 0 Then
                Using book As ExcelWorkbook = Me.Workbooks.Add()
                    book.Worksheets.Add()
                    Dim ws As ExcelWorksheet = book.Worksheets(1)
                    If TargetData.TableName <> "" Then
                        ws.Name = TargetData.TableName
                    End If

                    If TargetData.Columns.Count > 0 Then
                        Dim rangeHani As ExcelRange = ws.Range(1, 1, TargetData.Rows.Count + 1, TargetData.Columns.Count)
                        'rangeHani.NumberFormat = "@" '文字列
                        'rangeHani.Font.Name = "ＭＳ ゴシック"

                        For Each arg As Integer In New Integer() {ExcelBorders.XlBordersIndex.xlEdgeTop,
                                                                  ExcelBorders.XlBordersIndex.xlEdgeBottom,
                                                                  ExcelBorders.XlBordersIndex.xlEdgeLeft,
                                                                  ExcelBorders.XlBordersIndex.xlEdgeRight,
                                                                  ExcelBorders.XlBordersIndex.xlInsideHorizontal,
                                                                  ExcelBorders.XlBordersIndex.xlInsideVertical}
                            Dim border As ExcelBorder = rangeHani.Borders(arg)
                            border.LineStyle = ExcelBorder.XlLineStyle.xlContinuous
                            border.Weight = ExcelBorder.XlBorderWeight.xlThin
                        Next

                        Dim rowIndex As Integer = 0

                        'ヘッダ
                        rowIndex += 1
                        For columnIndex As Integer = 1 To TargetData.Columns.Count
                            Dim range As ExcelRange = ws.Cells(rowIndex, columnIndex)

                            range.Value = TargetData.Columns(columnIndex - 1).ColumnName

                        Next
                    End If

                    book.SaveAs(FileName, ExcelWorkbook.XlFileFormat.xlExcel8)
                End Using
            Else
                Dim book As ExcelWorkbook = Nothing
                Dim ws As ExcelWorksheet = Nothing

                Dim columnCount As Integer = TargetData.Columns.Count
                Dim rowCount As Integer = TargetData.Rows.Count


                Dim liColumns As Generic.List(Of Object) = Nothing
                Dim liValues As Generic.List(Of Object()) = Nothing

                Try
                    Dim fileCount As Integer = 0
                    For dbRowIndex As Integer = 0 To TargetData.Rows.Count - 1
                        Dim dbRow As DataRow = TargetData.Rows(dbRowIndex)

                        Dim createBook As Boolean = False
                        If dbRowIndex = 0 Then
                            createBook = True
                        ElseIf dbRowIndex Mod RowLimit = 0 Then
                            createBook = True
                        End If

                        If createBook Then
                            If book IsNot Nothing Then
                                ExportFile_SetAndExport(TargetData, FileName, book, ws, liValues, fileCount)
                                book.Dispose()
                                book = Nothing
                            End If
                            book = Me.Workbooks.Add()
                            book.Worksheets.Add()
                            ws = book.Worksheets(1)
                            If TargetData.TableName <> "" Then
                                ws.Name = TargetData.TableName
                            End If
                            liValues = New Generic.List(Of Object())

                            'ヘッダ
                            liColumns = New Generic.List(Of Object)
                            For excelColumnIndex As Integer = 1 To TargetData.Columns.Count
                                liColumns.Add(TargetData.Columns(excelColumnIndex - 1).ColumnName)
                            Next
                            liValues.Add(liColumns.ToArray)
                        End If

                        liColumns = New Generic.List(Of Object)
                        For excelColumnIndex As Integer = 1 To TargetData.Columns.Count
                            liColumns.Add(dbRow(excelColumnIndex - 1))
                        Next
                        liValues.Add(liColumns.ToArray)
                    Next

                    ExportFile_SetAndExport(TargetData, FileName, book, ws, liValues, fileCount)
                    book.Dispose()
                    book = Nothing
                Finally
                    If book IsNot Nothing Then
                        book.Dispose()
                        book = Nothing
                    End If
                End Try
            End If
        End Sub

        Private Sub ExportFile_SetAndExport(ByVal dt As DataTable,
                                            ByVal FileName As String,
                                            ByVal book As ExcelWorkbook,
                                            ByVal ws As ExcelWorksheet,
                                            ByVal liValues As Generic.List(Of Object()),
                                            ByRef fileCount As Integer)
            Dim saveFileName As String
            Dim valueArr(,) As Object
            Dim rangeHani As ExcelRange
            Dim oneLineValues() As Object
            ReDim oneLineValues(dt.Columns.Count - 1)

            ReDim valueArr(liValues.Count - 1, dt.Columns.Count - 1)
            For valRowIndex As Integer = 0 To liValues.Count - 1
                oneLineValues = liValues(valRowIndex)
                For valColumnIndex As Integer = 0 To oneLineValues.Length - 1
                    valueArr(valRowIndex, valColumnIndex) = oneLineValues(valColumnIndex)
                Next
            Next

            rangeHani = ws.Range(1, 1, liValues.Count, dt.Columns.Count)
            '値
            rangeHani.Value = valueArr
            '書式
            'rangeHani.NumberFormat = "@" '文字列
            'rangeHani.Font.Name = "ＭＳ ゴシック"
            '罫線
            For Each arg As Integer In New Integer() {ExcelBorders.XlBordersIndex.xlEdgeTop,
                                                      ExcelBorders.XlBordersIndex.xlEdgeBottom,
                                                      ExcelBorders.XlBordersIndex.xlEdgeLeft,
                                                      ExcelBorders.XlBordersIndex.xlEdgeRight,
                                                      ExcelBorders.XlBordersIndex.xlInsideHorizontal,
                                                      ExcelBorders.XlBordersIndex.xlInsideVertical}
                Dim border As ExcelBorder = rangeHani.Borders(arg)
                border.LineStyle = ExcelBorder.XlLineStyle.xlContinuous
                border.Weight = ExcelBorder.XlBorderWeight.xlThin
            Next
            '
            fileCount += 1
            saveFileName = IO.Path.GetDirectoryName(FileName) & "\" &
                            IO.Path.GetFileNameWithoutExtension(FileName) &
                            "_" & fileCount.ToString.PadLeft(5, "0") &
                            IO.Path.GetExtension(FileName)
            book.SaveAs(saveFileName, ExcelWorkbook.XlFileFormat.xlExcel8)
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' NOTE: マネージ状態を破棄します (マネージ オブジェクト)。
                    '破棄：マネージドクラス
                    '_ManagedXXXX = Nothing
                End If

                ' NOTE: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下の Finalize() をオーバーライドします。
                ' NOTE: 大きなフィールドを null に設定します。

                '破棄：アンマネージドリソースを管理しているマネージドクラス（IDisposable）
                If _ManagedExcelWorkbooks IsNot Nothing Then
                    _ManagedExcelWorkbooks.Dispose()
                    _ManagedExcelWorkbooks = Nothing
                End If

                If _NativeExcelApplication IsNot Nothing Then
                    _NativeExcelApplication.Quit()
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelApplication)
                    _NativeExcelApplication = Nothing
                End If
            End If
            Me.disposedValue = True
        End Sub

        ' NOTE: 上の Dispose(ByVal disposing As Boolean) にアンマネージ リソースを解放するコードがある場合にのみ、Finalize() をオーバーライドします。
        Protected Overrides Sub Finalize()
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(False)
            MyBase.Finalize()
        End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Namespace
