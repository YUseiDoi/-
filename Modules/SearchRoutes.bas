Attribute VB_Name = "SearchRoutes"
Option Explicit

'���[�U�[�t�H�[���ϐ�
Public formSearchParam As SearchParameters

'�����J�n�{�^���N���b�N
Public Function StartSearch()
    '�����t�H�[���̃C���X�^���X���쐬
    Set formSearchParam = New SearchParameters
    '�t�H�[���\��
    formSearchParam.Show
End Function

'API���g���ă��[�g����
Public Function GetRouteFromGoogleMapsAPI(ByVal sOrigin As String, ByVal sDest As String, ByVal sCityName As String)
    Dim tblAPIkey As ListObject 'API�L�[�e�[�u��
    Dim sAPIkey As String   'API�L�[
    Dim oHttpReq As XMLHTTP60   'API���X�|���X
    Dim dictResponse As Scripting.Dictionary 'API���X�|���X���`��̃f�[�^
    Dim sResponse As String     '���X�|���X�f�[�^
    Dim oReg As RegExp      '���K�\��
    Dim oMatch As Object        '��������
    Dim oEachMatch As Variant       '�e��������
    Dim oKey As Variant         'Dictionary��Key
    
    '�I�u�W�F�N�g�̏�����
    Set oHttpReq = New XMLHTTP60
    Set dictResponse = New Scripting.Dictionary
    Set oReg = New RegExp
    
    'API�L�[���擾
    Set tblAPIkey = ThisWorkbook.Sheets("API_KEY").ListObjects("APIKEY�e�[�u��")
    sAPIkey = tblAPIkey.ListColumns("API�L�[").DataBodyRange(1).Value
    
    'API���N�G�X�g�𑗂�
    oHttpReq.Open "GET", "https://maps.googleapis.com/maps/api/directions/json?origin=" & sOrigin & "&destination=" & sDest & "+" & sCityName & "&key=" & sAPIkey
    oHttpReq.send
    'API�R�[�����s�҂�
    Do While oHttpReq.readyState < 4
        DoEvents
    Loop
    
    'API���X�|���X�m�F
    sResponse = oHttpReq.responseText
    
    '���K�\���ݒ�
    oReg.Pattern = "[""html_instructions""].+,\n"        '���K�\���p�^�[��
    oReg.IgnoreCase = False     '�啶���Ə������̋��
    oReg.Global = True     '���͑S�̂���������
    
    '���͌���
    Set oMatch = oReg.Execute(sResponse)
    
    '�������ʂ������Ɋi�[
    For Each oEachMatch In oMatch
        dictResponse.Add oEachMatch, 1
    Next
    
    '�������ʂ���]�v�ȕ�������폜
    For Each oKey In dictResponse.Keys
        If InStr(1, oKey, "html_instructions", vbTextCompare) > 0 Then
            oKey = Left(oKey, InStr(1, oKey, """,", vbTextCompare) - 1)
            oKey = Right(oKey, Len(oKey) - InStr(1, oKey, ": """, vbTextCompare) - 2)
            oKey = ReplaceAllTargetStrings(oKey, "\u003cb\u003e", "")
            oKey = ReplaceAllTargetStrings(oKey, "\u003c/b\u003e", "")
            oKey = ReplaceAllTargetStrings(oKey, "/\u003cwbr/\u003e", "")
            oKey = ReplaceAllTargetStrings(oKey, "\u003cdiv style=\""font-size:0.9em\""\u003", "")
            oKey = ReplaceAllTargetStrings(oKey, "\u003c/div\u003e", "")
            '���[�U�[�t�H�[���Ɍ������ʂ�\��
            If formSearchParam.ResultBox.Value <> "" Then
                formSearchParam.ResultBox.Value = formSearchParam.ResultBox.Value & vbCrLf & oKey
            Else
                formSearchParam.ResultBox.Value = oKey
            End If
        Else
            dictResponse.Remove oKey
        End If
    Next
    
End Function

'�������1�s���擾
Public Function GetEachLine(ByVal sInput As String) As Scripting.Dictionary
    Dim sline As String
    Dim s As String
    Dim i As Long
    Dim lr As Long
    
    lr = 12
    For i = 1 To Len(sInput)
        s = Mid(sInput, i, 1)
        If s = vbLf Then
            Cells(lr, 2) = sline
            sline = ""
            lr = lr + 1
        ElseIf s <> vbCr Then
            sline = sline & s
        End If
    Next
End Function

'��������ꊇ�u������
Public Function ReplaceAllTargetStrings(ByVal sInput As String, ByVal sFindText As String, ByVal sReplaceText As String) As String
    Do While True
        If InStr(1, sInput, sFindText) > 0 Then
            sInput = Replace(sInput, sFindText, sReplaceText)
        Else
            Exit Do
        End If
    Loop
    ReplaceAllTargetStrings = sInput
End Function
