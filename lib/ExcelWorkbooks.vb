Option Infer On
Option Strict Off

Namespace ExcelUtility
    Public Class ExcelWorkbooks
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelWorkbookList As Generic.List(Of ExcelWorkbook)

        'ネイティブリソース
        Private _NativeExcelWorkbooks As Object

        'ネイティブリソース（引き渡されるのみ）
        Private _NativeRefonlyExcelApplication As Object


        Public Sub New(ByVal nativeExcelApplication As Object)
            _NativeRefonlyExcelApplication = nativeExcelApplication
            _NativeExcelWorkbooks = nativeExcelApplication.Workbooks
            _ManagedExcelWorkbookList = New Generic.List(Of ExcelWorkbook)
        End Sub

        Public Function Open(ByVal Filename As String) As ExcelWorkbook
            Dim book = New ExcelWorkbook(_NativeRefonlyExcelApplication, _NativeExcelWorkbooks.Open(Filename))
            _ManagedExcelWorkbookList.Add(book)
            Return book
        End Function

        Default Public ReadOnly Property Item(ByVal index As Integer) As ExcelWorkbook
            Get
                Dim book = New ExcelWorkbook(_NativeRefonlyExcelApplication, _NativeExcelWorkbooks.Item(index))
                _ManagedExcelWorkbookList.Add(book)
                Return book
            End Get
        End Property

        Public Function Add() As ExcelWorkbook
            Dim book = New ExcelWorkbook(_NativeRefonlyExcelApplication, _NativeExcelWorkbooks.Add())
            _ManagedExcelWorkbookList.Add(book)
            Return book
        End Function

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
                If _ManagedExcelWorkbookList IsNot Nothing Then
                    For Each book In _ManagedExcelWorkbookList
                        book.Dispose()
                        book = Nothing
                    Next
                End If

                If _NativeExcelWorkbooks IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelWorkbooks)
                    _NativeExcelWorkbooks = Nothing
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
