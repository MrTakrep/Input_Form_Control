VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hover_control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Ini adalah script di class module bernama hover_control
Option Explicit

Public WithEvents TextboxMouseUp As MSForms.TextBox
Attribute TextboxMouseUp.VB_VarHelpID = -1
Public WithEvents ComboboxMouseUp As MSForms.ComboBox
Attribute ComboboxMouseUp.VB_VarHelpID = -1
Public WithEvents TextboxEnter As MSForms.TextBox
Attribute TextboxEnter.VB_VarHelpID = -1
Public WithEvents ComboboxEnter As MSForms.ComboBox
Attribute ComboboxEnter.VB_VarHelpID = -1
Dim frm_ctrl As String

Private Sub TextboxEnter_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        Call Delete_Form_Control(TextboxEnter.Parent)
        Call Input_Form_Control(TextboxEnter.Parent)
        Call Check_Form_Control(TextboxEnter.Parent)
        If InStr(1, TextboxEnter.tag, "warna", vbTextCompare) > 0 Then
                    frm_ctrl = Right(TextboxEnter.tag, Len(TextboxEnter.tag) - InStr(TextboxEnter.tag, ":"))
                    frm_ctrl = Left(frm_ctrl, InStr(frm_ctrl, ":") - 1)
                    'MsgBox "hover_" & frm_ctrl
                    With TextboxEnter.Parent.Controls("hover_" & frm_ctrl & TextboxEnter.ControlTipText)
                        .Height = GetHeightFromInputLabel("hover_" & frm_ctrl)
                        .Left = TextboxEnter.Left - 12
                        .Picture = GetPictureFromInputLabel("hover_" & frm_ctrl)
                        .PicturePosition = fmPicturePositionCenter
                        .Width = GetWidthFromInputLabel("hover_" & frm_ctrl)
                        .Top = TextboxEnter.Top - 6
                        .ZOrder (0)
                    End With
                    'MsgBox TextboxEnter.Parent.Controls("hover_" & frm_ctrl & TextboxEnter.ControlTipText).Name
                    'MsgBox GetWidthFromInputLabel("hover_" & frm_ctrl)
        End If
        TextboxEnter.ZOrder (0)
    End If
End Sub

Private Sub ComboboxEnter_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        Call Delete_Form_Control(ComboboxEnter.Parent)
        Call Input_Form_Control(ComboboxEnter.Parent)
        Call Check_Form_Control(ComboboxEnter.Parent)
        If InStr(1, ComboboxEnter.tag, "warna", vbTextCompare) > 0 Then
                    frm_ctrl = Right(ComboboxEnter.tag, Len(ComboboxEnter.tag) - InStr(ComboboxEnter.tag, ":"))
                    frm_ctrl = Left(frm_ctrl, InStr(frm_ctrl, ":") - 1)
                    'MsgBox "hover_" & frm_ctrl
                    With ComboboxEnter.Parent.Controls("hover_" & frm_ctrl & ComboboxEnter.ControlTipText)
                        .Height = GetHeightFromInputLabel("hover_" & frm_ctrl)
                        .Left = ComboboxEnter.Left - 12
                        .Picture = GetPictureFromInputLabel("hover_" & frm_ctrl)
                        .PicturePosition = fmPicturePositionCenter
                        .Width = GetWidthFromInputLabel("hover_" & frm_ctrl)
                        .Top = ComboboxEnter.Top - 6
                        .ZOrder (0)
                    End With
                    'MsgBox ComboboxEnter.Parent.Controls("hover_" & frm_ctrl & ComboboxEnter.ControlTipText).Name
                    'MsgBox GetWidthFromInputLabel("hover_" & frm_ctrl)
        End If
        ComboboxEnter.ZOrder (0)
    End If
End Sub

Private Sub TextboxMouseUp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If InStr(1, TextboxMouseUp.tag, "warna", vbTextCompare) > 0 Then
        Call Delete_Form_Control(TextboxMouseUp.Parent)
        Call Input_Form_Control(TextboxMouseUp.Parent)
        Call Check_Form_Control(TextboxMouseUp.Parent)
        If InStr(1, TextboxMouseUp.tag, "warna", vbTextCompare) > 0 Then
                    frm_ctrl = Right(TextboxMouseUp.tag, Len(TextboxMouseUp.tag) - InStr(TextboxMouseUp.tag, ":"))
                    frm_ctrl = Left(frm_ctrl, InStr(frm_ctrl, ":") - 1)
                    'MsgBox "hover_" & frm_ctrl
                    With TextboxMouseUp.Parent.Controls("hover_" & frm_ctrl & TextboxMouseUp.ControlTipText)
                        .Height = GetHeightFromInputLabel("hover_" & frm_ctrl)
                        .Left = TextboxMouseUp.Left - 12
                        .Picture = GetPictureFromInputLabel("hover_" & frm_ctrl)
                        .PicturePosition = fmPicturePositionCenter
                        .Width = GetWidthFromInputLabel("hover_" & frm_ctrl)
                        .Top = TextboxMouseUp.Top - 6
                        .ZOrder (0)
                    End With
                    'MsgBox TextboxMouseUp.Parent.Controls("hover_" & frm_ctrl & TextboxMouseUp.ControlTipText).Name
                    'MsgBox GetWidthFromInputLabel("hover_" & frm_ctrl)
        End If
        TextboxMouseUp.ZOrder (0)
        End If
End Sub

Private Sub ComboboxMouseUp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If InStr(1, ComboboxMouseUp.tag, "warna", vbTextCompare) > 0 Then
        Call Delete_Form_Control(ComboboxMouseUp.Parent)
        Call Input_Form_Control(ComboboxMouseUp.Parent)
        Call Check_Form_Control(ComboboxMouseUp.Parent)
        If InStr(1, ComboboxMouseUp.tag, "warna", vbTextCompare) > 0 Then
                    frm_ctrl = Right(ComboboxMouseUp.tag, Len(ComboboxMouseUp.tag) - InStr(ComboboxMouseUp.tag, ":"))
                    frm_ctrl = Left(frm_ctrl, InStr(frm_ctrl, ":") - 1)
                    'MsgBox "hover_" & frm_ctrl
                    With ComboboxMouseUp.Parent.Controls("hover_" & frm_ctrl & ComboboxMouseUp.ControlTipText)
                        .Height = GetHeightFromInputLabel("hover_" & frm_ctrl)
                        .Left = ComboboxMouseUp.Left - 12
                        .Picture = GetPictureFromInputLabel("hover_" & frm_ctrl)
                        .PicturePosition = fmPicturePositionCenter
                        .Width = GetWidthFromInputLabel("hover_" & frm_ctrl)
                        .Top = ComboboxMouseUp.Top - 6
                        .ZOrder (0)
                    End With
                    'MsgBox ComboboxMouseUp.Parent.Controls("hover_" & frm_ctrl & ComboboxMouseUp.ControlTipText).Name
                    'MsgBox GetWidthFromInputLabel("hover_" & frm_ctrl)
        End If
        ComboboxMouseUp.ZOrder (0)
        End If
End Sub

Private Function GetPictureFromInputLabel(tag As String) As IPictureDisp
    Dim ctrl As Control
    
    ' Iterasi melalui semua kontrol di Input_Form
    For Each ctrl In Input_Form.Controls
        ' Periksa apakah tag kontrol sesuai dengan yang diinginkan
        If ctrl.tag = tag Then
            ' Jika tag sesuai, dan kontrol adalah Label
            ' Ambil gambar dari kontrol tersebut dan kembalikan
            If TypeOf ctrl Is MSForms.Label Then
                Set GetPictureFromInputLabel = ctrl.Picture
                Exit Function
            End If
        End If
    Next ctrl
    
    ' Jika tidak ditemukan kontrol dengan tag yang sesuai, kembalikan Nothing
    Set GetPictureFromInputLabel = Nothing
End Function

Private Function GetWidthFromInputLabel(tag As String) As Single
    Dim ctrl As Control
    
    ' Iterasi melalui semua kontrol di Input_Form
    For Each ctrl In Input_Form.Controls
        ' Periksa apakah tag kontrol sesuai dengan yang diinginkan
        If ctrl.tag = tag Then
            ' Jika tag sesuai, dan kontrol adalah Label
            ' Ambil lebar dari kontrol tersebut dan kembalikan
            If TypeOf ctrl Is MSForms.Label Then
                GetWidthFromInputLabel = ctrl.Width
                Exit Function
            End If
        End If
    Next ctrl
    
    ' Jika tidak ditemukan kontrol dengan tag yang sesuai, kembalikan 0
    GetWidthFromInputLabel = 0
End Function

Private Function GetHeightFromInputLabel(tag As String) As Single
    Dim ctrl As Control
    
    ' Iterasi melalui semua kontrol di Input_Form
    For Each ctrl In Input_Form.Controls
        ' Periksa apakah tag kontrol sesuai dengan yang diinginkan
        If ctrl.tag = tag Then
            ' Jika tag sesuai, dan kontrol adalah Label
            ' Ambil lebar dari kontrol tersebut dan kembalikan
            If TypeOf ctrl Is MSForms.Label Then
                GetHeightFromInputLabel = ctrl.Height
                Exit Function
            End If
        End If
    Next ctrl
    
    ' Jika tidak ditemukan kontrol dengan tag yang sesuai, kembalikan 0
    GetHeightFromInputLabel = 0
End Function

