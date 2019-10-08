Attribute VB_Name = "Everywhere"
Option Explicit
'##############################################################################################################################
'
'   いつでも使えるがクラス化するほどでもない関数まとめ
'   VBAの仕様上、エディタのコンソールにファイルで分割する機能がない為、
'   一つのファイルにまとめておいた方が利便性が上がる。
'
'##############################################################################################################################

'==============================================================================================================================
'   現在時刻を返す
'
'   T_Flag :  0,年から秒までのすべての時刻 (例)2018.04.23.03.54.02
'             1,年から日までの日付         (例)2018.04.23
'             2,時から秒までの時間         (例)03.54.02
'             3,年のみ                     (例)2018
'             4,月のみ                     (例)04
'             5,日のみ                     (例)23
'             6,時のみ                     (例)03
'             7,分のみ                     (例)54
'             8,秒のみ                     (例)02
'    それ以外,全て"0"として処理する
'   ToBe   : 間に入れてほしい文字列「2018/04/23」「2018.04.23」等 T_Flagの値が0～2の時のみ有効
'==============================================================================================================================
Public Function CurrentTime(Optional T_Flag As Long = 0, Optional ToBe As String = ".") As String
    Dim NowYear() As String
    Dim NowTime() As String
    NowYear = Split(Format(Date, "yyyy:mm:dd"), ":")
    NowTime = Split(Format(time, "hh:mm:ss"), ":")
    If T_Flag = 1 Then     '年から日までの日付
        CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2)
    ElseIf T_Flag = 2 Then '時から秒までの時間
        CurrentTime = NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
    ElseIf T_Flag = 3 Then '年のみ
        CurrentTime = NowYear(0)
    ElseIf T_Flag = 4 Then '月のみ
        CurrentTime = NowYear(1)
    ElseIf T_Flag = 5 Then '日のみ
        CurrentTime = NowYear(2)
    ElseIf T_Flag = 6 Then '時のみ
        CurrentTime = NowTime(0)
    ElseIf T_Flag = 7 Then '分のみ
        CurrentTime = NowTime(1)
    ElseIf T_Flag = 8 Then '秒のみ
        CurrentTime = NowTime(2)
    Else                 '0を含めそれ以外
        CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2) + ToBe + NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
    End If
End Function
'==============================================================================================================================
'   文字の中にある、特定の文字列から特定の文字列までを取得する
'   戻り値 : 分割した文字列、Splitと名があるがspecificAとspecificBも挿入されるので注意。
'
'   text      : とある文字列
'   specificA : 1つ目の特定の文字列
'   specificB : 2つ目の特定の文字列
'==============================================================================================================================
Public Function BetweenSplit(ByVal text As String, ByVal specificA As String, ByVal specificB As String) As String()
    Dim returnLength As Long: returnLength = 0
    Dim returnString() As String
    ReDim returnString(returnLength)
    'エラー対応のため戻り値を初期化
    BetweenSplit = returnString
    '空白の挿入を確認
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificA, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificB, "", vbBinaryCompare) = 0 Then Exit Function
    '同じ文字列なら用途が違う為
    If StrComp(specificA, specificB, vbBinaryCompare) = 0 Then
        BetweenSplit = Split(text, specificA)
        Exit Function
    End If
    
    Dim textArray1() As String
    Dim textArray2() As String
    Dim count1 As Long: count1 = 0
    Dim count2 As Long: count2 = 0
    '------------------------------
    ' "specificA" の処理
    '------------------------------
    textArray1 = Split(text, specificA)
    '配列数が0以下の場合話が違う
    If UBound(textArray1) <= 0 Then
        BetweenSplit = textArray1
        Exit Function
    End If
    
    For count1 = 0 To UBound(textArray1)
        If Not StrComp(textArray1(count1), "", vbBinaryCompare) = 0 Then
            If count1 = 0 Then
                returnString(returnLength) = textArray1(0) '0番目は完成
                returnLength = returnLength + 1
            Else
                textArray1(count1) = specificA + textArray1(count1)
            End If
        End If
    Next count1
    '------------------------------
    ' "specificB" の処理
    '------------------------------
    Dim Body As String
    For count1 = 1 To UBound(textArray1)
        If Not StrComp(textArray1(count1), "", vbBinaryCompare) = 0 Then
            Body = ""
            textArray2 = Split(textArray1(count1), specificB)
            
            For count2 = 1 To UBound(textArray2)
                Body = Body + specificB + textArray2(count2)
            Next count2
                
            If Not StrComp(textArray2(0), "", vbBinaryCompare) = 0 Then
                ReDim Preserve returnString(returnLength)
                returnString(returnLength) = textArray2(0)
                returnLength = returnLength + 1
            End If
            
            If Not StrComp(Body, "", vbBinaryCompare) = 0 Then
                ReDim Preserve returnString(returnLength)
                returnString(returnLength) = Body
                returnLength = returnLength + 1
            End If
        End If
    Next count1
    
    BetweenSplit = returnString
End Function
'==============================================================================================================================
'   Split関数の複数版
'
'   delimiters : 本家とは違い、Optional型でない、配列でないとエラー出るので注意
'   min        : delimitersのどの位置から区切りするのか : 負の数、delimiters以上はエラー
'   max        : delimitersのどの位置まで区切りするのか : 負の数は全て区切る、delimiters以上でも全て区切る
'
'   その他、引数の説明は下記URLを参照
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/split-function
'==============================================================================================================================
Public Function Splits(ByVal expression As String, delimiters() As String, Optional ByVal limit As Long = -1, Optional ByVal compare As VbCompareMethod = vbBinaryCompare, Optional ByVal min As Long = 0, Optional ByVal max As Long = -1) As String()
    Dim returnString() As String
    ReDim returnString(0)
    Splits = returnString
    If UBound(delimiters) < 0 Then Exit Function
    If min < 0 Then Exit Function
    If max < 0 Or max > UBound(delimiters) Then max = UBound(delimiters)
    
    Dim returnLength As Long: returnLength = 0
    Dim textCount As Long, textArray() As String
    Dim bodyCount As Long, bodyArray() As String
    Dim limitCount As Long, limitString As String
    '-1の部分はlimitなのでこれで良い
    textArray = Split(expression, delimiters(min), -1, compare)
    Splits = textArray
    If min = max Then Exit Function
    If min >= UBound(delimiters) Then Exit Function
    
    For textCount = 0 To UBound(textArray)
        If Not StrComp(textArray(textCount), "", vbBinaryCompare) = 0 Then
            '-1の部分はlimitなのでこれで良い
            bodyArray = Splits(textArray(textCount), delimiters, -1, compare, min + 1, max)
            For bodyCount = 0 To UBound(bodyArray)
                If Not StrComp(bodyArray(bodyCount), "", vbBinaryCompare) = 0 Then
                    ReDim Preserve returnString(returnLength)
                    returnString(returnLength) = bodyArray(bodyCount)
                    returnLength = returnLength + 1
                    
                    If limit >= 0 And returnLength >= limit Then
                        limitString = ""
                        For limitCount = bodyCount To UBound(bodyArray)
                            limitString = limitString + bodyArray(limitCount)
                        Next limitCount
                        
                        returnString(returnLength - 1) = limitString
                        Splits = returnString
                        Exit Function
                    End If
                End If
            Next bodyCount
        End If
    Next
    Splits = returnString
End Function
'******************************************************************************************************************************
'
'   クリップボード関連
'   使用するには、「Microsoft Forms 2.0 Object Library」を参照設定します。
'
'******************************************************************************************************************************

' クリップボードに文字列を設定する。
Public Function SetClipboard_Text(ByVal text As String)
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    With New MSForms.DataObject
        .SetText text
        .PutInClipboard
    End With
End Function

' クリップボードから文字列を取得する。
Public Function GetClipboard_Text() As String
    Dim text As String: text = ""
    With New MSForms.DataObject
        .GetFromClipboard
        text = .GetText
    End With
    GetText = text
End Function
'******************************************************************************************************************************
'
'   シート関連関数
'
'******************************************************************************************************************************

'シートの存在確認、なければ(Nothing)
Public Function searchSheet(ByVal sheetName As String) As Worksheet
    Set searchSheet = Nothing
    Dim sheet As Worksheet
    For Each sheet In Worksheets
        If StrComp(sheet.name, sheetName, vbBinaryCompare) = 0 Then
            Set searchSheet = sheet
            Exit For
        End If
    Next
    Set sheet = Nothing
End Function

'戻り値、OK(True),NG(False)：2010の場合
Public Function checkSheetName(ByVal sheetName As String) As Boolean
    checkSheetName = False
    If StrComp(sheetName, "", vbBinaryCompare) = 0 Then Exit Function
    'シート名に含んではいけない文字列
    Dim textFor As Variant
    For Each textFor In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(sheetName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    '文字列は31文字いない。
    If Len(sheetName) > 31 Then Exit Function
    '同じ名前のシートは存在してはならない。
    Dim sheet As Worksheet
    Set sheet = searchSheet(sheetName)
    If Not sheet Is Nothing Then
        Set sheet = Nothing
        Exit Function
    End If
    checkSheetName = True
End Function

'新たなシートを作成。作成できなければ(Nothing)
Public Function aNewSheet(ByVal sheetName As String) As Worksheet
    Set aNewSheet = Nothing
    
    Dim sheet As Worksheet
    If Not checkSheetName(sheetName) Then
        Set sheet = searchSheet(sheetName)
        If Not sheet Is Nothing Then
            Set aNewSheet = sheet
        End If
        Exit Function
    End If
    
    Set sheet = Worksheets.Add()
    sheet.name = sheetName
    sheet.Activate 'アクティブ化しておいた方が見た目は良い。
    Set aNewSheet = sheet
End Function
'==============================================================================================================================
'
'   数値を判定
'   戻り値 : はい(true),いいえ(false)
'
'   text  : 判定用の数値
'   value : 数数値の入った数値型(Long,Double)のどちらか、エラーの場合はEmptyが入る
'           最終的には型の判定が要ります。↓参考URL：例→ If VarType(value) = vbLong Then
'           http://officetanaka.net/excel/vba/function/VarType.htm
'
'==============================================================================================================================
Public Function checkNumericalValue(ByVal text As String, ByRef value As Variant) As Boolean
    
    text = StrConv(text, vbNarrow)
    text = StrConv(text, vbLowerCase)
    text = LCase(text)
    If IsNumeric(text) Then
        value = Val(text)
        If StrComp(CStr(value), CStr(CLng(CStr(value))), vbBinaryCompare) = 0 Then
            value = CLng(CStr(value))
        End If
        checkNumericalValue = True
    Else
        value = Empty
        checkNumericalValue = False
    End If
End Function
'******************************************************************************************************************************
'
'   ファイル・フォルダー関連
'
'******************************************************************************************************************************

'==============================================================================================================================
'   ファイル読み込み、ある程度の文字コードに対応している。
'   戻り値 : その読み込んだファイルの文字列: エラーの場合は空白
'
'   fileName       : フルパス
'   characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS),(非推奨：_autodetect_all)
'==============================================================================================================================
Public Function readFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS") As String
    readFile = ""
    If Not Dir(fileName) <> "" Then Exit Function
    
    Dim Body As String

On Error GoTo ErrorHandler
    With CreateObject("ADODB.Stream")
        .Type = 2   'adTypeText
        .Charset = characterCord
        .Open
        .LoadFromFile (fileName)
        Body = .ReadText(-1)
        .Close
    End With
    
    readFile = Body '原文保持
    Exit Function
ErrorHandler:
    readFile = ""
    Exit Function
End Function
'==============================================================================================================================
'   ファイル書き込み、ある程度の文字コードに対応している。
'   戻り値 : 成功(True),失敗(False)
'
'   text           : 保存用の文字列
'   fileName       : フルパス
'   characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS)
'   addFlag        : ファイルがある場合、追加で書き込む , 初期値(書き込まない)
'==============================================================================================================================
Public Function writeFile(ByRef text As String, ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS", Optional ByVal addFlag As Boolean = False) As Boolean
    writeFile = False
    '書き込むデータが無い場合。
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    '追加で書き込むための確認事項
    If addFlag Then
        If Not Dir(fileName) <> "" Then
            addFlag = False
        End If
    End If
    
    Dim Body As String: Body = ""
On Error GoTo ErrorHandler
    With CreateObject("ADODB.Stream")
        .Type = 2   'adTypeText
        .Charset = characterCord
        .Open
        If addFlag Then
            .LoadFromFile (fileName)
            Body = .ReadText(-1)
        End If
        .WriteText Body + text
        .SaveToFile fileName, 2
        .Close
    End With
    
    writeFile = True
    Exit Function
ErrorHandler:
    writeFile = False
    Exit Function
End Function
