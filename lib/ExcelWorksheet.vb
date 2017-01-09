Option Infer On
Option Strict Off

Namespace ExcelUtility
    Public Class ExcelWorksheet
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelRangeDictionary As Generic.Dictionary(Of String, ExcelRange)

        'ネイティブリソース
        Private _NativeExcelWorksheet As Object


        Public Sub New(ByVal nativeWorksheet As Object)
            _NativeExcelWorksheet = nativeWorksheet
            _ManagedExcelRangeDictionary = New Generic.Dictionary(Of String, ExcelRange)
        End Sub

        ''' <summary>
        ''' セルのExcelRangeを返却する
        ''' </summary>
        ''' <param name="RowIndex">開始行 1から始まる</param>
        ''' <param name="ColumnIndex">開始列 1から始まる</param>
        Public Function Cells(ByVal RowIndex As Integer, ByVal ColumnIndex As Integer) As ExcelRange
            Dim key As String = String.Format("{0}:{1}",
                                              New Object() {RowIndex, ColumnIndex})
            If _ManagedExcelRangeDictionary.ContainsKey(key) Then
                Return _ManagedExcelRangeDictionary(key)
            Else
                Dim r As New ExcelRange(_NativeExcelWorksheet.Cells(RowIndex, ColumnIndex))
                _ManagedExcelRangeDictionary.Add(key, r)
                Return r
            End If
        End Function

        ''' <summary>
        ''' セルのExcelRangeを返却する
        ''' </summary>
        ''' <param name="RowIndex">開始行</param>
        ''' <param name="ColumnIndex">開始列</param>
        Public Function Cells(ByVal RowIndex As Integer, ByVal ColumnIndex As String) As ExcelRange
            Dim key As String = String.Format("{0}:{1}",
                                              New Object() {RowIndex, ColumnIndex})
            If _ManagedExcelRangeDictionary.ContainsKey(key) Then
                Return _ManagedExcelRangeDictionary(key)
            Else
                Dim r As New ExcelRange(_NativeExcelWorksheet.Cells(RowIndex, ColumnIndex))
                _ManagedExcelRangeDictionary.Add(key, r)
                Return r
            End If
        End Function

        ''' <summary>
        ''' 指定した範囲のExcelRangeを返却する
        ''' </summary>
        ''' <param name="RowIndex1">範囲開始セルの行 1から始まる</param>
        ''' <param name="ColumnIndex1">範囲開始セルの列 1から始まる</param>
        ''' <param name="RowIndex2">範囲終了セルの行 1から始まる</param>
        ''' <param name="ColumnIndex2">範囲終了セルの列 1から始まる</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Range(ByVal RowIndex1 As Integer, ByVal ColumnIndex1 As Integer,
                              ByVal RowIndex2 As Integer, ByVal ColumnIndex2 As Integer) As ExcelRange
            Dim key As String = String.Format("{0}:{1}:{2}:{3}",
                                              New Object() {RowIndex1, ColumnIndex1, RowIndex2, ColumnIndex2})
            If _ManagedExcelRangeDictionary.ContainsKey(key) Then
                Return _ManagedExcelRangeDictionary(key)
            Else
                Dim r As New ExcelRange(_NativeExcelWorksheet.Range(
                                        _NativeExcelWorksheet.Cells(RowIndex1, ColumnIndex1),
                                        _NativeExcelWorksheet.Cells(RowIndex2, ColumnIndex2)))
                _ManagedExcelRangeDictionary.Add(key, r)
                Return r
            End If
        End Function

        Public Property [Name] As String
            Get
                Return _NativeExcelWorksheet.Name
            End Get
            Set(value As String)
                _NativeExcelWorksheet.Name = value
            End Set
        End Property

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
                If _ManagedExcelRangeDictionary IsNot Nothing Then
                    For Each r In _ManagedExcelRangeDictionary.Values
                        r.Dispose()
                        r = Nothing
                    Next
                End If

                If _NativeExcelWorksheet IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelWorksheet)
                    _NativeExcelWorksheet = Nothing
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
