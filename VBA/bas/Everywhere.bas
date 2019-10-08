Attribute VB_Name = "Everywhere"
Option Explicit
'##############################################################################################################################
'
'   ���ł��g���邪�N���X������قǂł��Ȃ��֐��܂Ƃ�
'   VBA�̎d�l��A�G�f�B�^�̃R���\�[���Ƀt�@�C���ŕ�������@�\���Ȃ��ׁA
'   ��̃t�@�C���ɂ܂Ƃ߂Ă������������֐����オ��B
'
'##############################################################################################################################

'==============================================================================================================================
'   ���ݎ�����Ԃ�
'
'   T_Flag :  0,�N����b�܂ł̂��ׂĂ̎��� (��)2018.04.23.03.54.02
'             1,�N������܂ł̓��t         (��)2018.04.23
'             2,������b�܂ł̎���         (��)03.54.02
'             3,�N�̂�                     (��)2018
'             4,���̂�                     (��)04
'             5,���̂�                     (��)23
'             6,���̂�                     (��)03
'             7,���̂�                     (��)54
'             8,�b�̂�                     (��)02
'    ����ȊO,�S��"0"�Ƃ��ď�������
'   ToBe   : �Ԃɓ���Ăق���������u2018/04/23�v�u2018.04.23�v�� T_Flag�̒l��0�`2�̎��̂ݗL��
'==============================================================================================================================
Public Function CurrentTime(Optional T_Flag As Long = 0, Optional ToBe As String = ".") As String
    Dim NowYear() As String
    Dim NowTime() As String
    NowYear = Split(Format(Date, "yyyy:mm:dd"), ":")
    NowTime = Split(Format(time, "hh:mm:ss"), ":")
    If T_Flag = 1 Then     '�N������܂ł̓��t
        CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2)
    ElseIf T_Flag = 2 Then '������b�܂ł̎���
        CurrentTime = NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
    ElseIf T_Flag = 3 Then '�N�̂�
        CurrentTime = NowYear(0)
    ElseIf T_Flag = 4 Then '���̂�
        CurrentTime = NowYear(1)
    ElseIf T_Flag = 5 Then '���̂�
        CurrentTime = NowYear(2)
    ElseIf T_Flag = 6 Then '���̂�
        CurrentTime = NowTime(0)
    ElseIf T_Flag = 7 Then '���̂�
        CurrentTime = NowTime(1)
    ElseIf T_Flag = 8 Then '�b�̂�
        CurrentTime = NowTime(2)
    Else                 '0���܂߂���ȊO
        CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2) + ToBe + NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
    End If
End Function
'==============================================================================================================================
'   �����̒��ɂ���A����̕����񂩂����̕�����܂ł��擾����
'   �߂�l : ��������������ASplit�Ɩ������邪specificA��specificB���}�������̂Œ��ӁB
'
'   text      : �Ƃ��镶����
'   specificA : 1�ڂ̓���̕�����
'   specificB : 2�ڂ̓���̕�����
'==============================================================================================================================
Public Function BetweenSplit(ByVal text As String, ByVal specificA As String, ByVal specificB As String) As String()
    Dim returnLength As Long: returnLength = 0
    Dim returnString() As String
    ReDim returnString(returnLength)
    '�G���[�Ή��̂��ߖ߂�l��������
    BetweenSplit = returnString
    '�󔒂̑}�����m�F
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificA, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificB, "", vbBinaryCompare) = 0 Then Exit Function
    '����������Ȃ�p�r���Ⴄ��
    If StrComp(specificA, specificB, vbBinaryCompare) = 0 Then
        BetweenSplit = Split(text, specificA)
        Exit Function
    End If
    
    Dim textArray1() As String
    Dim textArray2() As String
    Dim count1 As Long: count1 = 0
    Dim count2 As Long: count2 = 0
    '------------------------------
    ' "specificA" �̏���
    '------------------------------
    textArray1 = Split(text, specificA)
    '�z�񐔂�0�ȉ��̏ꍇ�b���Ⴄ
    If UBound(textArray1) <= 0 Then
        BetweenSplit = textArray1
        Exit Function
    End If
    
    For count1 = 0 To UBound(textArray1)
        If Not StrComp(textArray1(count1), "", vbBinaryCompare) = 0 Then
            If count1 = 0 Then
                returnString(returnLength) = textArray1(0) '0�Ԗڂ͊���
                returnLength = returnLength + 1
            Else
                textArray1(count1) = specificA + textArray1(count1)
            End If
        End If
    Next count1
    '------------------------------
    ' "specificB" �̏���
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
'   Split�֐��̕�����
'
'   delimiters : �{�ƂƂ͈Ⴂ�AOptional�^�łȂ��A�z��łȂ��ƃG���[�o��̂Œ���
'   min        : delimiters�̂ǂ̈ʒu�����؂肷��̂� : ���̐��Adelimiters�ȏ�̓G���[
'   max        : delimiters�̂ǂ̈ʒu�܂ŋ�؂肷��̂� : ���̐��͑S�ċ�؂�Adelimiters�ȏ�ł��S�ċ�؂�
'
'   ���̑��A�����̐����͉��LURL���Q��
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
    '-1�̕�����limit�Ȃ̂ł���ŗǂ�
    textArray = Split(expression, delimiters(min), -1, compare)
    Splits = textArray
    If min = max Then Exit Function
    If min >= UBound(delimiters) Then Exit Function
    
    For textCount = 0 To UBound(textArray)
        If Not StrComp(textArray(textCount), "", vbBinaryCompare) = 0 Then
            '-1�̕�����limit�Ȃ̂ł���ŗǂ�
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
'   �N���b�v�{�[�h�֘A
'   �g�p����ɂ́A�uMicrosoft Forms 2.0 Object Library�v���Q�Ɛݒ肵�܂��B
'
'******************************************************************************************************************************

' �N���b�v�{�[�h�ɕ������ݒ肷��B
Public Function SetClipboard_Text(ByVal text As String)
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    With New MSForms.DataObject
        .SetText text
        .PutInClipboard
    End With
End Function

' �N���b�v�{�[�h���當������擾����B
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
'   �V�[�g�֘A�֐�
'
'******************************************************************************************************************************

'�V�[�g�̑��݊m�F�A�Ȃ����(Nothing)
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

'�߂�l�AOK(True),NG(False)�F2010�̏ꍇ
Public Function checkSheetName(ByVal sheetName As String) As Boolean
    checkSheetName = False
    If StrComp(sheetName, "", vbBinaryCompare) = 0 Then Exit Function
    '�V�[�g���Ɋ܂�ł͂����Ȃ�������
    Dim textFor As Variant
    For Each textFor In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(sheetName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    '�������31�������Ȃ��B
    If Len(sheetName) > 31 Then Exit Function
    '�������O�̃V�[�g�͑��݂��Ă͂Ȃ�Ȃ��B
    Dim sheet As Worksheet
    Set sheet = searchSheet(sheetName)
    If Not sheet Is Nothing Then
        Set sheet = Nothing
        Exit Function
    End If
    checkSheetName = True
End Function

'�V���ȃV�[�g���쐬�B�쐬�ł��Ȃ����(Nothing)
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
    sheet.Activate '�A�N�e�B�u�����Ă��������������ڂ͗ǂ��B
    Set aNewSheet = sheet
End Function
'==============================================================================================================================
'
'   ���l�𔻒�
'   �߂�l : �͂�(true),������(false)
'
'   text  : ����p�̐��l
'   value : �����l�̓��������l�^(Long,Double)�̂ǂ��炩�A�G���[�̏ꍇ��Empty������
'           �ŏI�I�ɂ͌^�̔��肪�v��܂��B���Q�lURL�F�ၨ If VarType(value) = vbLong Then
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
'   �t�@�C���E�t�H���_�[�֘A
'
'******************************************************************************************************************************

'==============================================================================================================================
'   �t�@�C���ǂݍ��݁A������x�̕����R�[�h�ɑΉ����Ă���B
'   �߂�l : ���̓ǂݍ��񂾃t�@�C���̕�����: �G���[�̏ꍇ�͋�
'
'   fileName       : �t���p�X
'   characterCord  : �����R�[�h�w��(�C��) , �����l(Shift_JIS),(�񐄏��F_autodetect_all)
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
    
    readFile = Body '�����ێ�
    Exit Function
ErrorHandler:
    readFile = ""
    Exit Function
End Function
'==============================================================================================================================
'   �t�@�C���������݁A������x�̕����R�[�h�ɑΉ����Ă���B
'   �߂�l : ����(True),���s(False)
'
'   text           : �ۑ��p�̕�����
'   fileName       : �t���p�X
'   characterCord  : �����R�[�h�w��(�C��) , �����l(Shift_JIS)
'   addFlag        : �t�@�C��������ꍇ�A�ǉ��ŏ������� , �����l(�������܂Ȃ�)
'==============================================================================================================================
Public Function writeFile(ByRef text As String, ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS", Optional ByVal addFlag As Boolean = False) As Boolean
    writeFile = False
    '�������ރf�[�^�������ꍇ�B
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    '�ǉ��ŏ������ނ��߂̊m�F����
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
