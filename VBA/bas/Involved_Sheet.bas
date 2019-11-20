Attribute VB_Name = "Involved_Sheet"
Option Explicit
'##############################################################################################################################
'
'   シート関連
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/20
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

'==============================================================================================================================
'   その名前がシート名に適切な名前であるか検査する
'
'   戻り値 : OK(True), NG(False)
'
'   sheetName : シート名
'==============================================================================================================================
Public Function checkSheetName(ByVal sheetName As String) As Boolean
    checkSheetName = False
    '条件その1 : 空の名前ではない。
    If StrComp(sheetName, "", vbBinaryCompare) = 0 Then Exit Function
    '条件その2 : 含んではいけない文字列がない。
    Dim textFor As Variant
    For Each textFor In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(sheetName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    '条件その3 : 名前は31文字以内である。
    If Len(sheetName) > 31 Then Exit Function
    '条件その4 : 同名のシートは存在出来ない。
    'aNewSheetにて不具合が発生したので分割する。
    checkSheetName = True
End Function

'==============================================================================================================================
'   等しい名前のシートを探す。
'
'   戻り値 : 等しい名前を持つシート。ない場合は、Nothingが返却される
'
'   sheetName : シート名
'   book : 対象のブック（任意）
'==============================================================================================================================
Public Function sheetToEqualsName(ByVal sheetName As String, Optional ByRef book As Workbook = Nothing) As Worksheet

    Dim searchBook As Workbook
    If book Is Nothing Then
        Set searchBook = ThisWorkbook
    Else
        Set searchBook = book
    End If

    Dim sheet As Worksheet
    For Each sheet In searchBook.sheets
        If StrComp(sheet.name, sheetName, vbBinaryCompare) = 0 Then
            Set sheetToEqualsName = sheet
            Exit Function
        End If
    Next
    Set sheetToEqualsName = Nothing
End Function

'==============================================================================================================================
'   新たなシートを作成。
'
'   戻り値 : 作成出来なかった場合はNothingが返却される
'
'   sheetName : シート名
'   book : 対象のブック（任意）
'==============================================================================================================================
Public Function aNewSheet(ByVal sheetName As String, Optional ByRef book As Workbook = Nothing) As Worksheet
    Set aNewSheet = Nothing
    If Not checkSheetName(sheetName) Then Exit Function

    Dim addBook As Workbook
    If book Is Nothing Then
        Set addBook = ThisWorkbook
    Else
        Set addBook = book
    End If

    Dim sheet As Worksheet
    Set sheet = sheetToEqualsName(sheetName, addBook)
    If Not sheet Is Nothing Then
        Set aNewSheet = sheet
        Exit Function
    End If

    Set sheet = addBook.sheets.add()
    sheet.name = sheetName
    sheet.Activate 'アクティブ化しておいた方が見た目は良い。
    Set aNewSheet = sheet
End Function

'==============================================================================================================================
'   シートを削除する。
'
'   戻り値 : 成功(True), 失敗(False)
'
'   sheet : 削除するシート。成功した場合、アクセス不可になるので注意が必要
'   book  : 対象のブック（任意）
'==============================================================================================================================
Public Function aDeletedSheet(ByRef sheet As Worksheet, Optional ByRef book As Workbook = Nothing) As Boolean
    aDeletedSheet = False
    
    If sheet Is Nothing Then
        'Nothingなので、既に削除済みと仮定する。
        aDeletedSheet = True
        Exit Function
    End If
    
    Dim deleteBook As Workbook
    If book Is Nothing Then
        Set deleteBook = ThisWorkbook
    Else
        Set deleteBook = book
    End If
    
    'メッセージが表示されるが、基本的に邪魔でしかない為、非表示にしておく
    Application.DisplayAlerts = False
    
    Dim deleteSheet As Worksheet
    For Each deleteSheet In deleteBook.sheets
        If StrComp(sheet.name, deleteSheet.name, vbBinaryCompare) = 0 Then
            Call deleteBook.sheets(sheet.name).delete
            Set sheet = Nothing  'シートを削除する
            aDeletedSheet = True '戻り値を変更
            Exit For
        End If
    Next
    
    'メッセージを表示
    Application.DisplayAlerts = True
End Function
'==============================================================================================================================
'   シートの情報を全て削除する。
'
'   sheet : 対象シート
'==============================================================================================================================
Public Function aInfoErasureSheet(ByRef sheet As Worksheet)
    Dim i As Long: i = 0
    'セルを全て削除
    sheet.Cells.Clear
    sheet.Columns.Clear
    sheet.Rows.Clear
    'テーブルの情報を削除
    For i = sheet.ListObjects.count To 1 Step -1
        Call sheet.ListObjects.item(i).delete
    Next i
    '埋め込みグラフを削除
    For i = sheet.ChartObjects.count To 1 Step -1
        Call sheet.ChartObjects(i).delete
    Next i
    '印刷時のページ区切りを削除
    'sheet.DisplayPageBreaks = False
    'ピボットテーブルを削除
    For i = sheet.PivotTables.count To 1 Step -1
        Call sheet.PivotTables(i).ClearTable
    Next i
    '図、クリップアート、図形、SmartArtの削除
    For i = sheet.Shapes.count To 1 Step -1
        Call sheet.Shapes.item(i).delete
    Next i
    'ヘッター、フッターは完全に削除することは不可能らしい
    With sheet.PageSetup
        For i = .Pages.count To 1 Step -1
            .Pages.item(i).CenterFooter = ""
            .Pages.item(i).CenterHeader = ""
            .Pages.item(i).LeftFooter = ""
            .Pages.item(i).LeftHeader = ""
            .Pages.item(i).RightFooter = ""
            .Pages.item(i).RightHeader = ""
        Next i
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .DifferentFirstPageHeaderFooter = True
    End With
    
End Function


