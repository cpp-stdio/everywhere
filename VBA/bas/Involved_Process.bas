Attribute VB_Name = "Involved_Process"
'##############################################################################################################################
'
'   �����֘A
'   �uInvolved_Debug�v�̊֐��𗘗p���Ă���̂ŁA�����ɓǂݍ���ł�������
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2019/11/04
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'
'##############################################################################################################################

'�v���O�������s���Ԃ��J�n�v���p�̕ϐ�
Private beginTime As Date

'==============================================================
'
'   �v���O�����������Ԃ��v������B
'
'==============================================================

'���s���Ԍv���J�n
Public Function processMeasure_ToBegin()
    beginTime = Time
End Function

'���s���Ԍv���I��
Public Function performanceMeasure_ToEnd(Optional ByRef message As String = "")
    Call debugBox(message + "���s���Ԃ� " + Format(Time - beginTime, "nn��ss�b") + " �ł���", vbInformation + vbOKOnly)
End Function

'==============================================================
'
'   VBA�̃V�[�g�X�V���̏������y��������B
'   ���̊֐����ĂԂ����ŁA�������Ԃ�8.5�{�قǌ��シ��B
'
'   �Q�lURL
'   https://tonari-it.com/vba-processing-speed/
'
'==============================================================
Public Function reduceProcess_ToBegin()
    Application.Calculation = xlCalculationManual '�v�Z���[�h���}�j���A���ɂ���
    Application.EnableEvents = False              '�C�x���g���~������
    Application.ScreenUpdating = False            '��ʕ\���X�V���~������
End Function

Public Function reduceProcess_ToEnd()
    Application.Calculation = xlCalculationAutomatic '�v�Z���[�h�������ɂ���
    Application.EnableEvents = True                  '�C�x���g���J�n������
    Application.ScreenUpdating = True                '��ʕ\���X�V���J�n������
End Function
