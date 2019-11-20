Attribute VB_Name = "Involved_Other"
Option Explicit
'##############################################################################################################################
'
'   ���̑��A�悭�W�����������s�\�Ȋ֐��Q
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/11/12
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
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

'==============================================================================================================================
'
'   �z�񂪋�Ȃ̂��𔻒肷��
'   �Q�lURL : http://www.fingeneersblog.com/1612/
'
'   �߂�l : ��(true),��ł͂Ȃ�(false)
'
'   text  : ����p�̐��l
'   value : �����l�̓��������l�^(Long,Double)�̂ǂ��炩�A�G���[�̏ꍇ��Empty������
'           �ŏI�I�ɂ͌^�̔��肪�v��܂��B���Q�lURL�F�ၨ If VarType(value) = vbLong Then
'           http://officetanaka.net/excel/vba/function/VarType.htm
'
'==============================================================================================================================
Public Function isEmptyArray(arrayVariant As Variant) As Boolean
    isEmptyArray = True '�󂾂Ɖ���
On Error GoTo isEmptyArray_ErrorHandler
    'UBound�֐����g�p���ăG���[���������邩�ǂ������m�F
    If UBound(arrayVariant) > 0 Then
        isEmptyArray = False
    End If
    Exit Function
isEmptyArray_ErrorHandler:
    isEmptyArray = True
End Function
