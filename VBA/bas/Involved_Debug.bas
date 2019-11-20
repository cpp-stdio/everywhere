Attribute VB_Name = "Involved_Debug"
Option Explicit
'##############################################################################################################################
'
'   デバック関連
'   デバック時にのみ有効な関数が存在する。
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/14
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

Private Enum atDevelopmentSwitching
    modeDebug   'Debugだけだとエラーが表示されたため
    modeRelease 'リリースモードの場合はこっち
End Enum

'開発時はDebugにしておく
Private Const atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug
'------------------------------------------------------------------------------------------------------------------------------
'   デバック用のMsgBox。毎回書くのが面倒なので作った。
'   引数の説明も戻り値の説明も下記を参照。一部不要な箇所があったので、そこだけ省略
'
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function
'------------------------------------------------------------------------------------------------------------------------------
Public Function debugBox(ByRef prompt As String, Optional ByVal button As VbMsgBoxStyle = vbOKOnly, Optional ByRef title As String = "Microsoft Excel") As VbMsgBoxResult
    debugBox = vbOK
    'デバックモードでないと機能しない。
    If atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then
        debugBox = MsgBox(prompt, button, title)
    End If
End Function


'------------------------------------------------------------------------------------------------------------------------------
'   VBAをRANをした瞬間に上書き保存する機能がないので、セーブを手動で実装する。
'
'   book : 保存したいbook情報。設定しないとThisWorkbookが選択されます。
'------------------------------------------------------------------------------------------------------------------------------
Public Function debugBookSave(Optional ByVal book As Workbook = Nothing)
    
    Dim bookSave As Workbook
    If book Is Nothing Then
        Set bookSave = ThisWorkbook
    Else
        Set bookSave = book
    End If

    bookSave.Save
End Function
