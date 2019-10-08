VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadHierarchy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================================
'
'   フォルダの階層読み込み
'
'==============================================================================================================================
Private Type Access
    name As String             '名前を格納する
    createdDay As String       '作成日を格納する
    lastAccessedDay As String  '最終アクセス日を格納する
    lastModifiedDay As String  '最終更新日を格納する
    size As Single             'データサイズ(byte単位)
End Type

Dim FileLength As Long   'ファイル用
Dim File() As Access     'ファイル用
Dim FolderLength As Long 'フォルダー用
Dim Folder()  As Access  'フォルダー用
'------------------------------------------------------------------------------------------------------------------------------
'   コンストラクター
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Call CloseData
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   デストラクター
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Call CloseData
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   階層読み込み(本体) : フォルダやファイルにロックがかかっているものは読み込めないので注意
'   戻り値(Boolean型) : 成功(True) , エラー(False)
'
'   FolderHierarchy   : フォルダーの階層を読み込むパス名
'   DeepFolder        : さらに深い階層をどこまで読むのか。負の数は無限回、0はその場のみ
'   FullPathFlag      : フルパスで取得するが。     はい(True) いいえ(False)：初期値(True)
'------------------------------------------------------------------------------------------------------------------------------
Public Function Search(ByVal FolderHierarchyPath As String, Optional ByVal DeepFolder As Long = -1, Optional ByVal FullPathFlag As Boolean = True) As Boolean
    Search = False 'まずは失敗する前提で話を進める。
    
    '何も書かれていない状態はエラーとする。
    If StrComp(FolderHierarchyPath, "", vbTextCompare) = 0 Then Exit Function
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
     
    'フォルダが存在するかどうか？
    If Not FSO.FolderExists(FolderHierarchyPath) Then
        Set FSO = Nothing
        Exit Function
    End If
    
    Dim obj As Object
    'そのパスに存在するすべてのフォルダー名を取得する。
    For Each obj In FSO.getFolder(FolderHierarchyPath).Files
        ReDim Preserve File(FileLength)
        If FullPathFlag Then
            File(FileLength).name = FolderHierarchyPath + "\" + obj.name
        Else
            File(FileLength).name = obj.name
        End If
        File(FileLength).createdDay = CStr(obj.DateCreated)
        File(FileLength).lastModifiedDay = CStr(obj.DateLastModified)
        File(FileLength).lastAccessedDay = CStr(obj.DateLastAccessed)
        File(FileLength).size = obj.size
        FileLength = FileLength + 1
    Next obj
    
     'そのパスに存在するすべてのフォルダー名を取得し、再起処理を行う
    For Each obj In FSO.getFolder(FolderHierarchyPath).SubFolders
        ReDim Preserve Folder(FolderLength)
        Folder(FolderLength).name = CStr(obj.Path)
        Folder(FolderLength).createdDay = CStr(obj.DateCreated)
        Folder(FolderLength).lastModifiedDay = CStr(obj.DateLastModified)
        Folder(FolderLength).lastAccessedDay = CStr(obj.DateLastAccessed)
        Folder(FolderLength).size = obj.size
        FolderLength = FolderLength + 1
        '再起処理開始
        If DeepFolder >= 0 Then
            If Not DeepFolder = 0 Then
                Call Search(CStr(obj.Path), DeepFolder - 1, FullPathFlag)
            End If
        Else
            Call Search(CStr(obj.Path), -1, FullPathFlag)
        End If
    Next obj
    
    Set obj = Nothing
    Set FSO = Nothing
    Search = True '最後まで来たので成功とみなす。
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   特定の拡張子に絞り込む ↑のSearchを先に実行してないと意味ない。(やる前にコピーすることをおすすめする)
'   ※この関数に成功した場合、フォルダーの情報は全て削除される。
'   戻り値(Boolean型) : 成功(True) , エラー(False)
'
'   ExtensionName : 例.テキスト(.txt)の場合"txt"と入力して。最初のドットは不要
'------------------------------------------------------------------------------------------------------------------------------
Public Function ExtensionSearch(ByVal ExtensionName As String) As Boolean
    ExtensionSearch = False
    'データがない場合は処理できない
    If FileLength = 0 Then Exit Function
    
    Dim i As Long
    Dim count As Long: count = 0
    Dim Temporary() As Access
    
    For i = 0 To FileLength - 1
        '一番最後が拡張子
        Dim extension As Variant
        extension = Split(File(i).name, ".")
        'OSの仕様上、大文字小文字は同じ拡張子
        If StrComp(UCase(ExtensionName), UCase(extension(UBound(extension))), vbBinaryCompare) = 0 Then
            ReDim Preserve Temporary(count)
            Temporary(count) = File(i)
            count = count + 1
        End If
    Next i
    'その拡張子がない場合は再構築はしない
    If count <= 0 Then
        Exit Function
    End If
    
    'データの再構築
    Call CloseData
    FileLength = count
    ReDim File(FileLength - 1)
    File = Temporary
    ExtensionSearch = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   クローズ
'------------------------------------------------------------------------------------------------------------------------------
Public Function CloseData()
    FileLength = 0
    FolderLength = 0
    ReDim File(FileLength)
    ReDim Folder(FolderLength)
End Function
'==============================================================================================================================
'   ファイルゲッター
'==============================================================================================================================
Public Function getFileLength() As Long
    getFileLength = FileLength - 1 'For文を止める為
End Function
'ファイル名を取得する
Public Function getFileName(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFileLength() Then
        getFileName = File(Index).name
    Else
        getFileName = "" '予備用
    End If
End Function
'作成日を取得する
Public Function getFileCreatedDay(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFileLength() Then
        getFileCreatedDay = File(Index).createdDay
    Else
        getFileCreatedDay = "" '予備用
    End If
End Function
'最終更新日を取得する
Public Function getFileLastModifiedDay(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFileLength() Then
        getFileLastModifiedDay = File(Index).lastModifiedDay
    Else
        getFileLastModifiedDay = "" '予備用
    End If
End Function
'最終アクセス日を取得する
Public Function getFileLastAccessedDay(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFileLength() Then
        getFileLastAccessedDay = File(Index).lastAccessedDay
    Else
        getFileLastAccessedDay = "" '予備用
    End If
End Function
'バイト数を取得する
Public Function getFileByteSize(ByVal Index As Long) As Single
    If Index >= 0 And Index <= getFileLength() Then
        getFileByteSize = File(Index).size
    Else
        getFileByteSize = 0
    End If
End Function
'==============================================================================================================================
'   フォルダゲッター
'==============================================================================================================================
Public Function getFolderLength() As Long
    getFolderLength = FolderLength - 1 'For文を止める為
End Function
'ファイル名を取得する
Public Function getFolderName(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFolderLength() Then
        getFolderName = Folder(Index).name
    Else
        getFolderName = "" '予備用
    End If
End Function
'作成日を取得する
Public Function getFolderCreatedDay(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFolderLength() Then
        getFolderCreatedDay = Folder(Index).createdDay
    Else
        getFolderCreatedDay = "" '予備用
    End If
End Function
'最終更新日を取得する
Public Function getFolderLastModifiedDay(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFolderLength() Then
        getFolderLastModifiedDay = Folder(Index).lastModifiedDay
    Else
        getFolderLastModifiedDay = "" '予備用
    End If
End Function
'最終アクセス日を取得する
Public Function getFolderLastAccessedDay(ByVal Index As Long) As String
    If Index >= 0 And Index <= getFolderLength() Then
        getFolderLastAccessedDay = Folder(Index).lastAccessedDay
    Else
        getFolderLastAccessedDay = "" '予備用
    End If
End Function
'バイト数を取得する
Public Function getFolderByteSize(ByVal Index As Long) As Single
    If Index >= 0 And Index <= getFolderLength() Then
        getFolderByteSize = Folder(Index).size
    Else
        getFolderByteSize = 0
    End If
End Function
