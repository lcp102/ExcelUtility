Option Infer On
Option Strict Off

Namespace ExcelUtility
    Public Class ExcelRange
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelFontList As Generic.List(Of ExcelFont)
        Private _ManagedExcelBordersList As Generic.List(Of ExcelBorders)

        'ネイティブリソース
        Private _NativeExcelRange As Object

        Public Sub New(ByVal nativeRange As Object)
            _NativeExcelRange = nativeRange
            _ManagedExcelFontList = New Generic.List(Of ExcelFont)
            _ManagedExcelBordersList = New Generic.List(Of ExcelBorders)
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲の値を表すバリアント型 (Variant) の値を取得または設定します。値の取得および設定が可能です。
        ''' </summary>
        Public Property Value As Object
            Get
                Return _NativeExcelRange.Value
            End Get
            Set(value As Object)
                _NativeExcelRange.Value = value
            End Set
        End Property

        ''' <summary>
        ''' オブジェクトの数式を、A1 参照形式で、マクロ言語で表すバリアント型 (Variant) の値を取得または設定します。
        ''' </summary>
        Public Property Formula As Object
            Get
                Return _NativeExcelRange.Formula
            End Get
            Set(value As Object)
                _NativeExcelRange.Formula = value
            End Set
        End Property

        ''' <summary>
        ''' オブジェクトの表示形式を表すバリアント型 (Variant) の値を取得または設定します。
        ''' </summary>
        Public Property NumberFormat As Object
            Get
                Return _NativeExcelRange.NumberFormat
            End Get
            Set(value As Object)
                _NativeExcelRange.NumberFormat = value
            End Set
        End Property

        Public ReadOnly Property Font As ExcelFont
            Get
                If _ManagedExcelFontList.Count > 0 Then
                    Return _ManagedExcelFontList(0)
                Else
                    Dim f As New ExcelFont(_NativeExcelRange.Font)
                    _ManagedExcelFontList.Add(f)
                    Return f
                End If
            End Get
        End Property

        Public ReadOnly Property Borders As ExcelBorders
            Get
                If _ManagedExcelBordersList.Count > 0 Then
                    Return _ManagedExcelBordersList(0)
                Else
                    Dim b As New ExcelBorders(_NativeExcelRange.Borders)
                    _ManagedExcelBordersList.Add(b)
                    Return b
                End If
            End Get
        End Property


        Public Sub [Select]()
            _NativeExcelRange.[Select]()
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
                If _ManagedExcelFontList IsNot Nothing Then
                    For Each f In _ManagedExcelFontList
                        f.Dispose()
                        f = Nothing
                    Next
                End If
                If _ManagedExcelBordersList IsNot Nothing Then
                    For Each b In _ManagedExcelBordersList
                        b.Dispose()
                        b = Nothing
                    Next
                End If

                If _NativeExcelRange IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelRange)
                    _NativeExcelRange = Nothing
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
