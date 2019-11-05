Attribute VB_Name = "Involved_Process"
'##############################################################################################################################
'
'   �t�H�C���֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/11/04
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'
'##############################################################################################################################

'==============================================================================================================================
'   �t�@�C�������m�F����B
'
'   �߂�l : OK(True), NG(False)
'
'   fileName : �t�@�C����
'==============================================================================================================================
Public Function checkFileName(ByVal fileName As String) As Boolean
    checkFileName = False
    '��������1 : ��̖��O�ł͂Ȃ��B
    If StrComp(fileName, "", vbBinaryCompare) = 0 Then Exit Function
    '��������2 : �܂�ł͂����Ȃ������񂪂Ȃ��B
    Dim textFor As Variant
    For Each textFor In Array("��", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(fileName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    checkFileName = True
End Function

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

On Error GoTo readFile_ErrorHandler
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
readFile_ErrorHandler:
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
' Public Function writeFile(ByRef text As String, ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS", Optional ByVal addFlag As Boolean = False) As Boolean
'     writeFile = False
'     '�������ރf�[�^�������ꍇ�B
'     If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
'     '�ǉ��ŏ������ނ��߂̊m�F����
'     If addFlag Then
'         If Not Dir(fileName) <> "" Then
'             addFlag = False
'         End If
'     End If
'
'     Dim Body As String: Body = ""
' On Error GoTo writeFile_ErrorHandler
'     With CreateObject("ADODB.Stream")
'         .Type = 2   'adTypeText
'         .Charset = characterCord
'         .Open
'         If addFlag Then
'             .LoadFromFile (fileName)
'             Body = .ReadText(-1)
'         End If
'         .WriteText Body + text
'         .SaveToFile fileName, 2
'         .Close
'     End With
'
'     writeFile = True
'     Exit Function
' writeFile_ErrorHandler:
'     writeFile = False
'     Exit Function
' End Function
