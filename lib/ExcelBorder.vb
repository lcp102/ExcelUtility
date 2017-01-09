Option Strict Off

Namespace ExcelUtility
    Public Class ExcelBorder
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        'Private _ManagedYYYY

        'ネイティブリソース
        Private _NativeExcelBorder As Object

        Public Enum XlLineStyle
            ''' <summary>実線(細)</summary>
            xlContinuous = 1
            ''' <summary>破線</summary>
            xlDash = -4115
            ''' <summary>一点鎖線</summary>
            xlDashDot = 4
            ''' <summary>二点鎖線</summary>
            xlDashDotDot = 5
            ''' <summary>点線</summary>
            xlDot = -4118
            ''' <summary>二重線</summary>
            xlDouble = -4119
            ''' <summary>無し</summary>
            xlLineStyleNone = -4142
            ''' <summary>斜め斜線</summary>
            xlSlantDashDot = 13
        End Enum

        Public Enum XlBorderWeight
            ''' <summary>極細</summary>
            xlHairline = 1
            ''' <summary>細</summary>
            xlThin = 2
            ''' <summary>中</summary>
            xlMedium = -4138
            ''' <summary>太</summary>
            xlThick = 4
        End Enum

        Public Sub New(ByVal nativeBorder As Object)
            _NativeExcelBorder = nativeBorder
        End Sub

        Public Property LineStyle As XlLineStyle
            Get
                Return _NativeExcelBorder.LineStyle
            End Get
            Set(value As XlLineStyle)
                _NativeExcelBorder.LineStyle = value
            End Set
        End Property

        Public Property Weight As XlBorderWeight
            Get
                Return _NativeExcelBorder.Weight
            End Get
            Set(value As XlBorderWeight)
                _NativeExcelBorder.Weight = value
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

                If _NativeExcelBorder IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelBorder)
                    _NativeExcelBorder = Nothing
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
