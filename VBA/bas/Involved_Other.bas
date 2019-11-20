Attribute VB_Name = "Involved_Other"
Option Explicit
'##############################################################################################################################
'
'   その他、よくジャンル分け不能な関数群
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/12
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
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

'==============================================================================================================================
'
'   配列が空なのかを判定する
'   参考URL : http://www.fingeneersblog.com/1612/
'
'   戻り値 : 空(true),空ではない(false)
'
'   text  : 判定用の数値
'   value : 数数値の入った数値型(Long,Double)のどちらか、エラーの場合はEmptyが入る
'           最終的には型の判定が要ります。↓参考URL：例→ If VarType(value) = vbLong Then
'           http://officetanaka.net/excel/vba/function/VarType.htm
'
'==============================================================================================================================
Public Function isEmptyArray(arrayVariant As Variant) As Boolean
    isEmptyArray = True '空だと仮定
On Error GoTo isEmptyArray_ErrorHandler
    'UBound関数を使用してエラーが発生するかどうかを確認
    If UBound(arrayVariant) > 0 Then
        isEmptyArray = False
    End If
    Exit Function
isEmptyArray_ErrorHandler:
    isEmptyArray = True
End Function
