Option Infer On
Option Strict Off

Namespace ExcelUtility
    Public Class ExcelBorders
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelBorderDictionary As Generic.Dictionary(Of String, ExcelBorder)

        'ネイティブリソース
        Private _NativeExcelBorders As Object

        Public Enum XlBordersIndex
            xlEdgeTop = 8
            xlEdgeBottom = 9
            xlEdgeLeft = 7
            xlEdgeRight = 10
            xlInsideHorizontal = 12
            xlInsideVertical = 11
        End Enum

        Public Sub New(ByVal nativeBorders As Object)
            _NativeExcelBorders = nativeBorders
            _ManagedExcelBorderDictionary = New Generic.Dictionary(Of String, ExcelBorder)
        End Sub

        Default Public ReadOnly Property Item(ByVal index As XlBordersIndex) As ExcelBorder
            Get
                Dim key As String = index.ToString
                If _ManagedExcelBorderDictionary.ContainsKey(key) Then
                    Return _ManagedExcelBorderDictionary.Item(key)
                Else
                    Dim border = New ExcelBorder(_NativeExcelBorders.Item(index))
                    _ManagedExcelBorderDictionary.Add(key, border)
                    Return border
                End If
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As Integer) As ExcelBorder
            Get
                Dim key As String = index.ToString
                If _ManagedExcelBorderDictionary.ContainsKey(key) Then
                    Return _ManagedExcelBorderDictionary.Item(key)
                Else
                    Dim border = New ExcelBorder(_NativeExcelBorders.Item(index))
                    _ManagedExcelBorderDictionary.Add(key, border)
                    Return border
                End If
            End Get
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
                If _ManagedExcelBorderDictionary IsNot Nothing Then
                    For Each border In _ManagedExcelBorderDictionary.Values
                        border.Dispose()
                        border = Nothing
                    Next
                End If

                If _NativeExcelBorders IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelBorders)
                    _NativeExcelBorders = Nothing
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
