Option Infer On
Option Strict Off

Namespace ExcelUtility
    Public Class ExcelWorksheets
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelWorksheetDictionary As Generic.Dictionary(Of String, ExcelWorksheet)

        'ネイティブリソース
        Private _NativeExcelWorksheets As Object

        Public Sub New(ByVal nativeExcelApplication As Object)
            _NativeExcelWorksheets = nativeExcelApplication.Worksheets
            _ManagedExcelWorksheetDictionary = New Generic.Dictionary(Of String, ExcelWorksheet)
        End Sub

        Default Public ReadOnly Property Item(ByVal index As Integer) As ExcelWorksheet
            Get
                Dim key As String = index.ToString
                If _ManagedExcelWorksheetDictionary.ContainsKey(key) Then
                    Return _ManagedExcelWorksheetDictionary(key)
                Else
                    Dim sheet = New ExcelWorksheet(_NativeExcelWorksheets.Item(index))
                    _ManagedExcelWorksheetDictionary.Add(key, sheet)
                    Return sheet
                End If
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As String) As ExcelWorksheet
            Get
                Dim key As String = index.ToString
                If _ManagedExcelWorksheetDictionary.ContainsKey(key) Then
                    Return _ManagedExcelWorksheetDictionary(key)
                Else
                    Dim sheet = New ExcelWorksheet(_NativeExcelWorksheets.Item(index))
                    _ManagedExcelWorksheetDictionary.Add(key, sheet)
                    Return sheet
                End If
            End Get
        End Property

        Public Sub Add()
            _NativeExcelWorksheets.Add()
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
                If _ManagedExcelWorksheetDictionary IsNot Nothing Then
                    For Each sheet In _ManagedExcelWorksheetDictionary.Values
                        sheet.Dispose()
                        sheet = Nothing
                    Next
                End If

                If _NativeExcelWorksheets IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelWorksheets)
                    _NativeExcelWorksheets = Nothing
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
