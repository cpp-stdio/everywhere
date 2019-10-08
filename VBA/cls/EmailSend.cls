VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EMailSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================================
'   Outlookを使用し、メールを送信する。
'   最終チェックver : Outlook2010{14,0,7145,5000(32ビット)}
'==============================================================================================================================
Option Explicit
Dim OutlookApplication As Object
Dim OutlookMailItem As Object

Dim StartingFlag As Boolean
Const SilentMode = True '決してTrueに変えてはならぬ
'------------------------------------------------------------------------------------------------------------------------------
'   コンストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    StartingFlag = True
    Set OutlookMailItem = Nothing
    'Outlookの起動を確認
    On Error GoTo ErrorHandler '下記でエラーが発生するため
    Set OutlookApplication = GetObject(, "Outlook.Application")
ErrorHandler:
    If OutlookApplication Is Nothing Then
        StartingFlag = False
        '起動していないため起動させる。
        Dim Namespace As Object
        Set OutlookApplication = CreateObject("Outlook.Application")
        'Set Namespace = OutlookApplication.GetNamespace("MAPI")
        'Namespace.GetDefaultFolder(6).Display 'olFolderInbox
        Set Namespace = Nothing
    End If
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   デストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    '起動してなかったら終了
    If StartingFlag Then
        OutlookApplication.Quit
    End If
    Set OutlookApplication = Nothing
    Set OutlookMailItem = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   OutlookのMSGファイルオープン : 開けた(True)開けなかった(False)
'------------------------------------------------------------------------------------------------------------------------------
Public Function OpenFile(FileName As String) As Boolean
    'ファイルがない場合は...
    If Not Dir(FileName) <> "" Then
        OpenFile = False
        Exit Function
    End If
    '共有アイテムモードで開く
    Set OutlookMailItem = OutlookApplication.Session.OpenSharedItem(FileName)
    OpenFile = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
' メールアドレスかどうかの判定
'------------------------------------------------------------------------------------------------------------------------------
Private Function IsMailAddress(Address As String) As Boolean
    If Address Like "*@*" Then
        IsMailAddress = True
    Else
        IsMailAddress = False
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------
' 送信
'------------------------------------------------------------------------------------------------------------------------------
Public Function Send() As Boolean
    If Not SilentMode Then
        OutlookMailItem.Display
    Else
        Dim path As String: path = ActiveWorkbook.path + "\provisional.msg"
        OutlookMailItem.SaveAs path
        Application.Wait Now + TimeValue("0:00:03")  '念のため保存が完了するまで待つ
        'Outlookをすべて強制終了させる。
        Dim Explorer As Object
        For Each Explorer In OutlookApplication.Explorers
            Explorer.Application.Quit
        Next
        Application.Wait Now + TimeValue("0:00:03")  '念のため終了するまで待つ
        Set OutlookApplication = Nothing
        Set OutlookMailItem = Nothing
        'Outlookオブジェクトを再取得
        Set OutlookApplication = CreateObject("Outlook.Application")
        Set OutlookMailItem = OutlookApplication.Session.OpenSharedItem(path)
        OutlookMailItem.Send
        Application.Wait Now + TimeValue("0:00:03") '念のため送信する時間待つ
        StartingFlag = False                        '起動していた事実さえ消し去る
        '先ほど保存したメールファイルを削除
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Call FSO.DeleteFile(path, True)
        Set FSO = Nothing
        Application.Wait Now + TimeValue("0:00:03")  '念のため削除が完了するまで待つ
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------
'
' 本文を変更or拡張したいなら...
'
'------------------------------------------------------------------------------------------------------------------------------
Public Function GetBody() As String
    If OutlookMailItem.BodyFormat = 1 Then
        GetBody = OutlookMailItem.Body     '簡単モード
    Else
        GetBody = OutlookMailItem.HTMLBody '鬼畜モード
    End If
End Function
Public Function SetBody(Text As String) As String
    If OutlookMailItem.BodyFormat = 1 Then
        OutlookMailItem.Body = Text     '簡単モード
    Else
        OutlookMailItem.HTMLBody = Text '鬼畜モード
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------
'
' 件名を変更or拡張したいなら...
'
'------------------------------------------------------------------------------------------------------------------------------
Public Function GetSubject() As String
    GetSubject = OutlookMailItem.Subject
End Function
Public Function SetSubject(Subject As String) As String
    OutlookMailItem.Subject = Subject
End Function

'引数の説明は → http://outlook-navi.com/vbs/vbs11.htm
Public Function GetBodyFormat() As Long
    GetBodyFormat = OutlookMailItem.BodyFormat
End Function
