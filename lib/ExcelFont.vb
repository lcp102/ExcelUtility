Option Strict Off

Namespace ExcelUtility
    Public Class ExcelFont
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        'Private _ManagedYYYY

        'ネイティブリソース
        Private _NativeExcelFont As Object

        Public Sub New(ByVal nativeFont As Object)
            _NativeExcelFont = nativeFont
        End Sub

        ''' <summary>
        ''' フォントの名前を表すバリアント型 (Variant) の値を取得または設定します。
        ''' </summary>
        Public Property Name As Object
            Get
                Return _NativeExcelFont.Name
            End Get
            Set(value As Object)
                _NativeExcelFont.Name = value
            End Set
        End Property

        ''' <summary>
        ''' True の場合、フォントを太字にします。値の取得および設定が可能です。
        ''' </summary>
        Public Property Bold As Boolean
            Get
                Return _NativeExcelFont.Bold
            End Get
            Set(value As Boolean)
                _NativeExcelFont.Bold = value
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
                'If _ManagedYYYY IsNot Nothing Then
                '    _ManagedYYYY.Dispose()
                '    _ManagedYYYY = Nothing
                'End If

                If _NativeExcelFont IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelFont)
                    _NativeExcelFont = Nothing
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
