@echo off
rem �������Łu���O��t���ĕۑ��v����ۂɁA�����R�[�h���uANSI�v�ɂ��Ă��������B

echo =======================================================
echo  Excel�摜�ϊ��c�[�� �����v���O����
echo =======================================================
echo.
echo �c�[���ɕK�v�ȃt�H���_�ƃ��C�u�������������܂�...
echo.

echo --- 1. �t�H���_���쐬��...
if not exist "excels" mkdir excels
if not exist "images" mkdir images
echo    -> ����
echo.

echo --- 2. PDF�ϊ��c�[��(Poppler)���_�E�����[�h��...
echo    (����ɂ͏������Ԃ�������ꍇ������܂�)
curl -L "https://github.com/oschwartz10612/poppler-windows/releases/download/v25.07.0-0/Release-25.07.0-0.zip" -o "poppler.zip"
echo    -> ����
echo.

echo --- 3. Poppler��W�J��...
tar -xf poppler.zip
echo    -> ����
echo.

echo --- 4. �t�H���_���� 'poppler' �ɕύX��...
rem ���������� �����̃t�H���_�����C�����܂��� ����������
if exist "poppler-25.07.0" (
    rename "poppler-25.07.0" "poppler"
    echo    -> ����
) else (
    echo    -> �t�H���_��������Ȃ����߁A�X�L�b�v���܂����B
)
echo.

echo --- 5. Poppler��ZIP�t�@�C�����폜��...
del poppler.zip
echo    -> ����
echo.

echo --- 6. Python���C�u�������C���X�g�[����...
python -m venv venv
call venv\Scripts\activate
pip install pdf2image pywin32
echo    -> ����
echo.
echo =======================================================
echo  �S�Ă̏������������܂����I
echo =======================================================
echo.
pause