Attribute VB_Name = "Involved_Array"
Option Explicit
'##############################################################################################################################
'
'   �z��֘A�֐�
'   VBA�̔z��ɂ�2��ނ���BVariant�ŕύX�\�ȃ^�C�v�������łȂ��^�C�v�B����ɂ��֐���2��ޕK�v�ɂȂ�B
'
'   �V�K�쐬�� : 2019/11/18
'   �ŏI�X�V�� : 2019/11/20
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'
'##############################################################################################################################

'==============================================================================================================================
'
'   �z�񂪋�Ȃ̂��𔻒肷��
'   ���̊֐���VBA�̎d�l��A�ǂ����֐������邱�Ƃ��o���Ȃ��ׁA�قړ����R�[�h��2�񏑂��K�v������B
'   �Q�lURL : http://www.fingeneersblog.com/1612/
'
'   �߂�l : ��(true),��ł͂Ȃ�(false)
'
'   arrayVariant : ����p�̔z��
'
'==============================================================================================================================
Public Function arrayIsEmpty(ByRef arrayVariant As Variant) As Boolean
    arrayIsEmpty = True '�󂾂Ɖ���
On Error GoTo isEmptyArray_ErrorHandler

    'UBound�֐����g�p���ăG���[���������邩�ǂ������m�F
    If UBound(arrayVariant) > 0 Then
        arrayIsEmpty = False
    End If
    Exit Function
    
isEmptyArray_ErrorHandler:
    arrayIsEmpty = True
End Function

Public Function arrayIsEmptyEx(ByRef arrayVariant() As Variant) As Boolean
    arrayIsEmptyEx = True '�󂾂Ɖ���
On Error GoTo isEmptyArrayEx_ErrorHandler

    'UBound�֐����g�p���ăG���[���������邩�ǂ������m�F
    If UBound(arrayVariant) > 0 Then
        arrayIsEmptyEx = False
    End If
    Exit Function

isEmptyArrayEx_ErrorHandler:
    arrayIsEmptyEx = True
End Function

'==============================================================================================================================
'
'   �z��̈ꕔ��؂�o���A�V�����z��Ƃ��ĕԋp����B
'
'   �߂�l : ����(True), ���s(False)
'
'   oldArray : �؂�o���p�̔z��
'   newArray : �ԋp�p�z��
'   min      : �ǂ�����
'   max      : �ǂ��܂�
'==============================================================================================================================
Public Function arraySplit(ByRef oldArray As Variant, ByRef newArray As Variant, Optional ByVal min As Long = -&HFF, Optional ByVal max As Long = -&HFF) As Boolean
    arraySplit = False '���s�Ɖ���
    If arrayIsEmpty(oldArray) Then Exit Function
    If errorSplit(min, max, LBound(oldArray), UBound(oldArray)) Then Exit Function
    'VBA�̎d�l�ケ�������͌ʂŏ����Ȃ���΂Ȃ�Ȃ��B
    Dim i As Long
    Dim length As Long: length = -1
    
    If VarType(newArray) = vbEmpty Then
        newArray = Array()
    End If
    
    For i = min To max
        length = length + 1
        ReDim Preserve newArray(length)
        newArray(length) = oldArray(i)
    Next i
    
    arraySplit = True
End Function

Public Function arraySplitEx(ByRef oldArray() As Variant, ByRef newArray() As Variant, Optional ByVal min As Long = -&HFF, Optional ByVal max As Long = -&HFF) As Boolean
    arraySplitEx = False '���s�Ɖ���
    If arrayIsEmptyEx(oldArray) Then Exit Function
    If errorSplit(min, max, LBound(oldArray), UBound(oldArray)) Then Exit Function
    'VBA�̎d�l�ケ�������͌ʂŏ����Ȃ���΂Ȃ�Ȃ��B
    Dim i As Long
    Dim length As Long: length = -1
    For i = min To max
        length = length + 1
        ReDim Preserve newArray(length)
        newArray(length) = oldArray(i)
    Next i
    
    arraySplitEx = True
End Function

Private Function errorSplit(ByRef min As Long, ByRef max As Long, ByVal minArray As Long, ByVal maxArray As Long) As Boolean
    errorSplit = True

    If min < minArray Then
        min = minArray
    End If
    
    If max > maxArray Then
        max = maxArray
    End If
    
    'VBA�̎d�l�œ��������ł�OK�Ƃ���B
    If min < max Then Exit Function
    
    errorSplit = False
End Function

'==============================================================================================================================
'
'   �z��̔��]
'   ���̊֐���VBA�̎d�l��A�ǂ����֐������邱�Ƃ��o���Ȃ��ׁA�قړ����R�[�h��2�񏑂��K�v������B
'
'   �߂�l : ����(True), ���s(False)
'
'   reversed : ���]����z��
'
'==============================================================================================================================
Public Function arrayReversed(ByRef oldArray As Variant, ByRef newArray As Variant) As Boolean
    arrayReversed = False
    If arrayIsEmpty(oldArray) Then Exit Function
    
    'oldArray��newArray���������ƃ�������j�󂵂Ă��܂���
    Dim old As Variant
    old = arrayCopy(oldArray)
    
    ReDim newArray(UBound(old))
    
    Dim i As Long
    For i = LBound(old) To UBound(old)
        newArray(UBound(old) - i) = old(i)
    Next i
    arrayReversed = True
    
End Function

Public Function arrayReversedEx(ByRef oldArray() As Variant, ByRef newArray() As Variant) As Boolean
    arrayReversedEx = False
    If arrayIsEmptyEx(oldArray) Then Exit Function
    
    'oldArray��newArray���������ƃ�������j�󂵂Ă��܂���
    Dim old() As Variant
    old = arrayCopyEx(oldArray)
    
    ReDim newArray(UBound(old))
    
    Dim i As Long
    For i = LBound(old) To UBound(old)
        newArray(UBound(old) - i) = old(i)
    Next i
    arrayReversedEx = True
End Function

'==============================================================================================================================
'
'   �z��̃R�s�[
'   ���̊֐���VBA�̎d�l��A�ǂ����֐������邱�Ƃ��o���Ȃ��ׁA�قړ����R�[�h��2�񏑂��K�v������B
'
'   �߂�l : �R�s�[�����z��
'
'   copy : ���]����z��
'
'==============================================================================================================================
Public Function arrayCopy(ByRef copy As Variant) As Variant
    arrayCopy = Empty
    If arrayIsEmpty(copy) Then Exit Function

    Dim c As Variant
    ReDim c(UBound(copy))
    
    Dim i As Long
    For i = LBound(copy) To UBound(copy)
        c(i) = copy(i)
    Next i
    arrayCopy = c
End Function

Public Function arrayCopyEx(ByRef copy() As Variant) As Variant()
    Dim c() As Variant
    arrayCopyEx = c
    
    If arrayIsEmptyEx(copy) Then Exit Function

    ReDim c(UBound(copy))
    
    Dim i As Long
    For i = LBound(copy) To UBound(copy)
        c(i) = copy(i)
    Next i
    arrayCopyEx = c
End Function
