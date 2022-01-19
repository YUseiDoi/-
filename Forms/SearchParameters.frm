VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchParameters 
   Caption         =   "検索フォーム"
   ClientHeight    =   3708
   ClientLeft      =   -168
   ClientTop       =   -696
   ClientWidth     =   12048
   OleObjectBlob   =   "SearchParameters.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SearchParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sOrigin As String       '出発地点
Dim sDest As String         '目標地点
Dim sCityName As String     '目標地点の街

'検索ボタン
Private Sub SearchButton_Click()
    'テキストボックスからテキスト取得
    sOrigin = OriginBox.Value
    sDest = DestBox.Value
    sCityName = DestCityBox.Value
    If sOrigin <> "" And sDest <> "" And sCityName <> "" Then
        '検索スタート
        Call GetRouteFromGoogleMapsAPI(sOrigin, sDest, sCityName)
    End If
End Sub

'フォームの初期化
Private Sub UserForm_Initialize()
    'テキストボックスに複数行改行できるようにする
    ResultBox.MultiLine = True
    ResultBox.EnterKeyBehavior = True
End Sub

'終了ボタン
Private Sub FinishButton_Click()
    Me.Hide
    End
End Sub
