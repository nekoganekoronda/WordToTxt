'***** Word�t�@�C�����R�s�[���A�e�L�X�g�t�@�C�����쐬����VBScript *****

Option Explicit

'�h���b�O�A���h�h���b�v�����t�@�C���̐�΃p�X���i�[
Dim GetPathArray
Set GetPathArray = WScript.Arguments
'�t�@�C���V�X�e���I�u�W�F�N�g
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'�C�e���[�^
Dim pt

'�쐬����e�L�X�g�t�@�C���̐�΃p�X
Dim strTextFilePath
'Word�̃I�u�W�F�N�g
Dim objWordApp
Dim objWordDoc

'�t�@�C���̐��������[�v����
For Each pt in GetPathArray

    'Word�̊g���q���������Atxt����������
    strTextFilePath = Left(pt, Len(pt) - 3) & "txt"

    '���[�h�̃I�u�W�F�N�g���쐬
    Set objWordApp = WScript.CreateObject("Word.Application")

    '�G���[���������Ȃ������ꍇ
    If Err.Number = 0 Then
        '���[�h�h�L�������g���J��
        Set objWordDoc = objWordApp.Documents.Open(pt)

        '�G���[���������Ȃ������ꍇ
        If Err.Number = 0 Then
            '�e�L�X�g�`���ŕۑ�
            objWordDoc.SaveAs strTextFilePath, 2

            objWordDoc.Close
            objWordApp.Quit
        Else
            WScript.Echo "�G���[�F" & Err.Descripticon
        End If
    Else
        WScript.Echo "�G���[�F" & Err.Descripticon
    End If
Next

'�I�u�W�F�N�g�ϐ����N���A
Set objFSO = Nothing
Set objWordApp = Nothing
Set objWordDoc = Nothing