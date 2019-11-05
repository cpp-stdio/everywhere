Option Explicit
'##############################################################################################################################
'
'   �V�[�g�֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/10/28
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'
'##############################################################################################################################

'==============================================================================================================================
'   ���̖��O���V�[�g���ɓK�؂Ȗ��O�ł��邩��������
'
'   �߂�l : OK(True), NG(False)
'
'   sheetName : �V�[�g��
'==============================================================================================================================
Public Function checkSheetName(ByVal sheetName As String) As Boolean
    checkSheetName = False
    '��������1 : ��̖��O�ł͂Ȃ��B
    If StrComp(sheetName, "", vbBinaryCompare) = 0 Then Exit Function
    '��������2 : �܂�ł͂����Ȃ������񂪂Ȃ��B
    Dim textFor As Variant
    For Each textFor In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(sheetName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    '��������3 : ���O��31�����ȓ��ł���B
    If Len(sheetName) > 31 Then Exit Function
    '��������4 : �����̃V�[�g�͑��ݏo���Ȃ��B
    'aNewSheet�ɂĕs������������̂ŕ�������B
    checkSheetName = True
End Function
'==============================================================================================================================
'   ���������O�̃V�[�g��T���B
'
'   �߂�l : ���������O�����V�[�g�B�Ȃ��ꍇ�́ANothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�Ӂj
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
'   �V���ȃV�[�g���쐬�B
'
'   �߂�l : �쐬�o���Ȃ������ꍇ��Nothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�Ӂj
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
    sheet.Activate '�A�N�e�B�u�����Ă��������������ڂ͗ǂ��B
    Set aNewSheet = sheet
End Function
