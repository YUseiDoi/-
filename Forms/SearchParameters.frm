VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchParameters 
   Caption         =   "�����t�H�[��"
   ClientHeight    =   3708
   ClientLeft      =   -168
   ClientTop       =   -696
   ClientWidth     =   12048
   OleObjectBlob   =   "SearchParameters.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SearchParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sOrigin As String       '�o���n�_
Dim sDest As String         '�ڕW�n�_
Dim sCityName As String     '�ڕW�n�_�̊X

'�����{�^��
Private Sub SearchButton_Click()
    '�e�L�X�g�{�b�N�X����e�L�X�g�擾
    sOrigin = OriginBox.Value
    sDest = DestBox.Value
    sCityName = DestCityBox.Value
    If sOrigin <> "" And sDest <> "" And sCityName <> "" Then
        '�����X�^�[�g
        Call GetRouteFromGoogleMapsAPI(sOrigin, sDest, sCityName)
    End If
End Sub

'�t�H�[���̏�����
Private Sub UserForm_Initialize()
    '�e�L�X�g�{�b�N�X�ɕ����s���s�ł���悤�ɂ���
    ResultBox.MultiLine = True
    ResultBox.EnterKeyBehavior = True
End Sub

'�I���{�^��
Private Sub FinishButton_Click()
    Me.Hide
    End
End Sub
