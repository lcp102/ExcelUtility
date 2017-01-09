Option Strict Off

Namespace ExcelUtility
    Public Class ExcelWorkbook
        Implements IDisposable

        'マネージドクラス
        'Private _ManagedXXXX

        'アンマネージドリソースを管理しているマネージドクラス（IDisposable）
        Private _ManagedExcelWorksheets As ExcelWorksheets

        'ネイティブリソース
        Private _NativeExcelWorkbook As Object

        Private _Closed As Boolean
        Private _ExcelApplicationVersion As Integer

        Public Enum XlFileFormat
            ''' <summary>ブックの標準</summary>
            xlWorkbookNormal = -4143
            ''' <summary>Excel 97 - 2003のバイナリファイル形式（BIFF8）</summary>
            xlExcel8 = 56
            ''' <summary>Excel2007ベースのファイル形式、VBAなし</summary>
            xlWorkbookDefault = 51
            ''' <summary>Excel2007ベースのマクロ有効ファイル形式、VBA有効</summary>
            xlOpenXMLWorkbookMacroEnabled = 52
        End Enum

        Public Sub New(ByVal nativeExcelApplication As Object, ByVal nativeWorkbook As Object)
            _ManagedExcelWorksheets = New ExcelWorksheets(nativeExcelApplication)
            _ExcelApplicationVersion = CInt(Val(nativeExcelApplication.Version))
            _NativeExcelWorkbook = nativeWorkbook
            _Closed = False
        End Sub

        Public Sub Close()
            _NativeExcelWorkbook.Close()
            _Closed = True
        End Sub

        Public Sub Close(ByVal SaveChanges As Boolean)
            _NativeExcelWorkbook.Close(SaveChanges)
            _Closed = True
        End Sub

        Public Sub Close(ByVal SaveChanges As Boolean, ByVal Filename As String)
            _NativeExcelWorkbook.Close(SaveChanges, Filename)
            _Closed = True
        End Sub


        ''' <summary>
        ''' ブックへの変更を別のファイルに保存します。
        ''' </summary>
        ''' <param name="Filename">保存するファイルの名前を表す文字列を指定します。完全パスを含めることもできます。完全パスを含めない場合は、ファイルは現在のフォルダーに保存されます。</param>
        ''' <remarks></remarks>
        Public Sub SaveAs(ByVal Filename As String)
            _NativeExcelWorkbook.SaveAs(Filename)
        End Sub

        ''' <summary>
        ''' ブックへの変更を別のファイルに保存します。
        ''' </summary>
        ''' <param name="Filename">保存するファイルの名前を表す文字列を指定します。完全パスを含めることもできます。完全パスを含めない場合は、ファイルは現在のフォルダーに保存されます。</param>
        ''' <param name="FileFormat">使っているEXCELのバージョンが97-2003の場合は、指定しても無視されます。ファイルを保存するときのファイル形式を指定します。指定できる形式については、XlFileFormat 列挙体の説明を参照してください。既存のファイルでは、指定された最後のファイル形式が既定のファイル形式です。新しいファイルでは、現在使用されている Excel のバージョンでのファイル形式が既定のファイル形式です。</param>
        ''' <remarks></remarks>
        Public Sub SaveAs(ByVal Filename As String, ByVal FileFormat As XlFileFormat)
            If _ExcelApplicationVersion < 12 Then
                '使っているEXCELのバージョンが97-2003
                Me.SaveAs(Filename)
            Else
                _NativeExcelWorkbook.SaveAs(Filename, FileFormat)
            End If
        End Sub

        ''' <summary>
        ''' 指定されたブックへの変更を保存します。
        ''' </summary>
        ''' <remarks>
        ''' ブックを開くときは、Openメソッドを使います。
        ''' 初めてブックを保存するときは、SaveAs メソッドを使ってファイル名を指定します。
        ''' </remarks>
        Public Sub Save()
            If _Closed Then Return
            _NativeExcelWorkbook.Save()
        End Sub

        ''' <summary>
        ''' ブックに関連付けられている最初のウィンドウをアクティブにします。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Activate()
            If _Closed Then Return
            _NativeExcelWorkbook.Activate()
        End Sub

        Public ReadOnly Property Worksheets() As ExcelWorksheets
            Get
                Return _ManagedExcelWorksheets
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
                If _ManagedExcelWorksheets IsNot Nothing Then
                    _ManagedExcelWorksheets.Dispose()
                    _ManagedExcelWorksheets = Nothing
                End If


                If _NativeExcelWorkbook IsNot Nothing Then
                    If Not _Closed Then
                        Try
                            _NativeExcelWorkbook.Close()
                        Catch ex As System.Runtime.InteropServices.COMException
                            If Not ex.Message.Contains("RPC_E_DISCONNECTED") Then
                                Throw New Exception("ExcelのWorkbookを閉じる際に想定外のエラーが発生しました。", ex)
                            End If
                        End Try
                    End If
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_NativeExcelWorkbook)
                    _NativeExcelWorkbook = Nothing
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
