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
Public Function CurrentTime(Optional ByVal T_Flag As Long = 0, Optional ByVal ToBe As String = ".") As String
    Dim NowYear() As String
    Dim NowTime() As String
    NowYear = Split(Format(Date, "yyyy:mm:dd"), ":")
    NowTime = Split(Format(Time, "hh:mm:ss"), ":")
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
Public Function checkNumericalValue(ByVal Text As String, ByRef value As Variant) As Boolean

    Text = StrConv(Text, vbNarrow)
    Text = StrConv(Text, vbLowerCase)
    Text = LCase(Text)
    If IsNumeric(Text) Then
        value = Val(Text)
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
'   ������̒�����A�����݂̂𔲂��o���B�Q�lURL��
'   https://vbabeginner.net/vba%E3%81%A7%E6%96%87%E5%AD%97%E5%88%97%E3%81%8B%E3%82%89%E6%95%B0%E5%AD%97%E3%81%AE%E3%81%BF%E3%82%92%E6%8A%BD%E5%87%BA%E3%81%99%E3%82%8B/
'
'   �߂�l : �����o���������A�G���[�̏ꍇ�͋�̔z�񂪕ԋp����܂��B
'
'   text  : �������܂܂�镶����
'
'==============================================================================================================================
Public Function findNumber(ByVal Text As String) As Variant()
    Dim reg As Object     '���K�\���N���X�I�u�W�F�N�g
    Dim matches As Object 'RegExp.Execute����
    Dim match As Object   '�������ʃI�u�W�F�N�g
    Dim i As Long         '���[�v�J�E���^
    
    Dim returnVariant() As Variant
    ReDim returnVariant(0)
    findNumber = returnVariant
    
    Set reg = CreateObject("VBScript.RegExp")
    
    '�����͈́�������̍Ō�܂Ō���
    reg.Global = True
    '��������������������
    reg.Pattern = "[0-9]"
    '�������s
    Set matches = reg.Execute(Text)
    '������v�����������[�v
    For i = 0 To matches.count - 1
        '�R���N�V�����̌����[�v�I�u�W�F�N�g���擾
        Set match = matches.item(i)
        '������v������
        ReDim Preserve returnVariant(i)
        returnVariant(i) = match.value
    Next
    findNumber = returnVariant
End Function
