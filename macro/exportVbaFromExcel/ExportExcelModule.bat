@echo off

rem ---------------------------------------------------------------
rem �@�\�FVBA�̃\�[�X��Export����
rem ---------------------------------------------------------------
rem �g�����F
rem     ���o�b�`�t�@�C���Ɖ��L�Ăяo���Ă���ExportExcelModule.vbs��
rem     VBA�\�[�X���L�q����Excel�t�@�C���̕ۑ��t�H���_�̐e�t�H���_�Ɋi�[����
rem ---------------------------------------------------------------

rem VBA�̃\�[�X�̕ۑ��ӏ���ݒ�
set EXPORT_PATH="E:\xxxxxxxxxxx"

rem ���̃o�b�`�����݂���t�H���_���J�����g�Ɉړ�
pushd %0\..

cls


rem --------------------------------------------------------------------
rem �J�����g�t�H���_�ƃT�u�t�H���_�Ɋ܂߂Ă���S�Ă�EXCEL�ixlsm���ΏۂɂȂ�j�����[�v���A
rem ExportExcelModule.vbs�Ń\�[�X���G�N�X�|�[�g����
rem (���[�v�͂����ł��Ă��邽�߁A�኱���\��������...VBS�ōċA�����������Ȃ��čςށB)
rem --------------------------------------------------------------------
for /F "usebackq" %%i in (`dir /s /b *.xls `) do ( 
    echo %%i 
    CScript ExportExcelModule.vbs %%i %EXPORT_PATH%
    rem pause
)
pause
exit

rem --------------------------------------------------------------------
rem �׋�������
rem �����P�F�J�����g�f�B���N�g���̊g���q��xls�̃t�@�C�����o��
rem    for %%i in (*.xls) do ( echo %%i )
rem �����Q(���K�\���������Ȃ�����)�F
rem for /F "usebackq" %%i in (`dir /s /b *.xls ^| findstr /V ".*\.xls$" `) do ()
rem --------------------------------------------------------------------