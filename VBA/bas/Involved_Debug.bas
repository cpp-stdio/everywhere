Attribute VB_Name = "Involved_Debug"
Option Explicit
'##############################################################################################################################
'
'   �f�o�b�N�֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/11/05
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'
'##############################################################################################################################

Private Enum atDevelopmentSwitching
    modeDebug   'Debug�������ƃG���[���\�����ꂽ����
    modeRelease '�����[�X���[�h�̏ꍇ�͂�����
End Enum

'�J������Debug�ɂ��Ă���
Private Const atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug
'------------------------------------------------------------------------------------------------------------------------------
'   �f�o�b�N�p��MsgBox�B���񏑂��̂��ʓ|�Ȃ̂ō�����B
'   �����̐������߂�l�̐��������L���Q�ƁB�ꕔ�s�v�ȉӏ����������̂ŁA���������ȗ�
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function
'------------------------------------------------------------------------------------------------------------------------------
Public Function debugBox(ByRef prompt As String, Optional ByVal button As VbMsgBoxStyle = vbOKOnly, Optional ByRef title As String = "Microsoft Excel") As VbMsgBoxResult
    debugBox = vbOK
    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then
        debugBox = MsgBox(prompt, button, title)
    End If
End Function
