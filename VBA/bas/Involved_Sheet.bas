Attribute VB_Name = "Involved_Sheet"
Option Explicit
'##############################################################################################################################
'
'   �V�[�g�֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/11/20
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
'   �V���ȃV�[�g���쐬�B
'
'   �߂�l : �쐬�o���Ȃ������ꍇ��Nothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�Ӂj
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
    sheet.Activate '�A�N�e�B�u�����Ă��������������ڂ͗ǂ��B
    Set aNewSheet = sheet
End Function

'==============================================================================================================================
'   �V�[�g���폜����B
'
'   �߂�l : ����(True), ���s(False)
'
'   sheet : �폜����V�[�g�B���������ꍇ�A�A�N�Z�X�s�ɂȂ�̂Œ��ӂ��K�v
'   book  : �Ώۂ̃u�b�N�i�C�Ӂj
'==============================================================================================================================
Public Function aDeletedSheet(ByRef sheet As Worksheet, Optional ByRef book As Workbook = Nothing) As Boolean
    aDeletedSheet = False
    
    If sheet Is Nothing Then
        'Nothing�Ȃ̂ŁA���ɍ폜�ς݂Ɖ��肷��B
        aDeletedSheet = True
        Exit Function
    End If
    
    Dim deleteBook As Workbook
    If book Is Nothing Then
        Set deleteBook = ThisWorkbook
    Else
        Set deleteBook = book
    End If
    
    '���b�Z�[�W���\������邪�A��{�I�Ɏז��ł����Ȃ��ׁA��\���ɂ��Ă���
    Application.DisplayAlerts = False
    
    Dim deleteSheet As Worksheet
    For Each deleteSheet In deleteBook.sheets
        If StrComp(sheet.name, deleteSheet.name, vbBinaryCompare) = 0 Then
            Call deleteBook.sheets(sheet.name).delete
            Set sheet = Nothing  '�V�[�g���폜����
            aDeletedSheet = True '�߂�l��ύX
            Exit For
        End If
    Next
    
    '���b�Z�[�W��\��
    Application.DisplayAlerts = True
End Function
'==============================================================================================================================
'   �V�[�g�̏���S�č폜����B
'
'   sheet : �ΏۃV�[�g
'==============================================================================================================================
Public Function aInfoErasureSheet(ByRef sheet As Worksheet)
    Dim i As Long: i = 0
    '�Z����S�č폜
    sheet.Cells.Clear
    sheet.Columns.Clear
    sheet.Rows.Clear
    '�e�[�u���̏����폜
    For i = sheet.ListObjects.count To 1 Step -1
        Call sheet.ListObjects.item(i).delete
    Next i
    '���ߍ��݃O���t���폜
    For i = sheet.ChartObjects.count To 1 Step -1
        Call sheet.ChartObjects(i).delete
    Next i
    '������̃y�[�W��؂���폜
    'sheet.DisplayPageBreaks = False
    '�s�{�b�g�e�[�u�����폜
    For i = sheet.PivotTables.count To 1 Step -1
        Call sheet.PivotTables(i).ClearTable
    Next i
    '�}�A�N���b�v�A�[�g�A�}�`�ASmartArt�̍폜
    For i = sheet.Shapes.count To 1 Step -1
        Call sheet.Shapes.item(i).delete
    Next i
    '�w�b�^�[�A�t�b�^�[�͊��S�ɍ폜���邱�Ƃ͕s�\�炵��
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


