Attribute VB_Name = "Involved_Book"
Option Explicit
'##############################################################################################################################
'
'   �u�b�N�֘A�̃}�N��
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/10/28
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'
'##############################################################################################################################


'==============================================================================================================================
'   ���������O�̃V�[�g��T���B
'
'   �߂�l : ���������O�����V�[�g�B�Ȃ��ꍇ�́ANothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�Ӂj
'==============================================================================================================================
Public Function BookToEqualsName(ByVal bookName As String) As Workbook
    Set BookToEqualsName = Nothing

    Dim book As Workbook
    For Each book In Workbooks
        If StrComp(book.name, bookName, vbBinaryCompare) = 0 Then
            Set BookToEqualsName = book
            Exit Function
        End If
    Next
End Function
