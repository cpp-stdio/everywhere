Option Explicit
'##############################################################################################################################
'
'   シート関連
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/10/28
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
Public Function sheetToEqualsName(ByVal sheetName As String, Optional ByVal book As Workbook = Nothing) As Worksheet

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
Public Function aNewSheet(ByVal sheetName As String, Optional ByVal book As Workbook = Nothing) As Worksheet
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

    Set sheet = addBook.sheets.Add()
    sheet.name = sheetName
    sheet.Activate 'アクティブ化しておいた方が見た目は良い。
    Set aNewSheet = sheet
End Function
