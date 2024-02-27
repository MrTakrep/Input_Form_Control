Attribute VB_Name = "bootstrap_form_control"
Option Explicit
Dim box_label As MSForms.Label
Dim ctrl, ctrl2 As Control
Dim frm_ctrl As String
Sub Delete_Form_Control(frm As UserForm)
    ' Memeriksa apakah kontrol adalah label dan namanya mengandung "warna"__DL
    For Each ctrl In frm.Controls
        If (TypeName(ctrl) = "Label" Or TypeName(ctrl) = "ComboBox") And InStr(ctrl.Name, "warna") > 0 Then
            frm.Controls.Remove ctrl.Name
        End If
    Next ctrl
End Sub

Sub Input_Form_Control(frm As UserForm)
    For Each ctrl In frm.Controls
        'Mengidentifikasi seluruh Textbox / ComboBox yang tag nya mengandung kata warna___TC
        If (TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox") And InStr(ctrl.tag, "warna") Then
            'Membuat label format berdasarkan tag yang mengandung kata warna___LC
            With frm
                'Mengambil format warna di tag berdasarkan pemisah : : ___W
                frm_ctrl = Right(ctrl.tag, Len(ctrl.tag) - InStr(ctrl.tag, ":"))
                frm_ctrl = Left(frm_ctrl, InStr(frm_ctrl, ":") - 1)
                '___W
                Set box_label = .Controls.Add("Forms.Label.1", "hover_" & frm_ctrl & ctrl.ControlTipText, True)
                With box_label
                    'Membuat format textbox kategori default___D
                    For Each ctrl2 In Input_Form.Controls
                        If ctrl2.tag = frm_ctrl And InStr(ctrl2.tag, "default") Then
                            .ControlTipText = "hover_" & frm_ctrl
                            .Left = ctrl.Left - 7
                            .Height = ctrl2.Height
                            .Picture = ctrl2.Picture
                            .PicturePosition = fmPicturePositionCenter
                            .Width = ctrl2.Width
                            .Top = ctrl.Top
                            .ZOrder (0)
                            Exit For
                        End If
                    Next ctrl2
                    '___D
                    'Membuat format textbox kategori big___B
                    For Each ctrl2 In Input_Form.Controls
                        If ctrl2.tag = frm_ctrl And InStr(ctrl2.tag, "big") Then
                            .ControlTipText = "hover_" & frm_ctrl
                            .Left = ctrl.Left - 7
                            .Height = ctrl2.Height
                            .Picture = ctrl2.Picture
                            .PicturePosition = fmPicturePositionCenter
                            .Width = ctrl2.Width
                            .Top = ctrl.Top
                            .ZOrder (0)
                            Exit For
                        End If
                    Next ctrl2
                    '___B
                    'Membuat format textbox kategori small___S
                    For Each ctrl2 In Input_Form.Controls
                        If ctrl2.tag = frm_ctrl And InStr(ctrl2.tag, "small") Then
                            .ControlTipText = "hover_" & frm_ctrl
                            .Left = ctrl.Left - 7
                            .Height = ctrl2.Height
                            .Picture = ctrl2.Picture
                            .PicturePosition = fmPicturePositionCenter
                            .Width = ctrl2.Width
                            .Top = ctrl.Top
                            .ZOrder (0)
                            Exit For
                        End If
                    Next ctrl2
                    '___S
                End With
            End With
                '___LC
        End If
        '___TC
        ctrl.ZOrder (0)
    Next ctrl
End Sub

Sub Check_Form_Control(frm As UserForm)
    For Each ctrl In frm.Controls
        'Mengidentifikasi seluruh Textbox / ComboBox yang tag nya mengandung kata warna___TC
        If (TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox") And InStr(ctrl.tag, "warna") Then
            'Mengubah label format berdasarkan tag yang mengandung kata warna___LC
            With frm
                'Mengambil format warna di tag berdasarkan pemisah : : ___W
                frm_ctrl = Right(ctrl.tag, Len(ctrl.tag) - InStr(ctrl.tag, ":"))
                frm_ctrl = Left(frm_ctrl, InStr(frm_ctrl, ":") - 1)
                '___W
                With box_label
                    'Mengubah format textbox kategori default___D
                    For Each ctrl2 In Input_Form.Controls
                        If ctrl2.tag = frm_ctrl And InStr(ctrl2.tag, "default") Then
                            .ControlTipText = "hover_" & frm_ctrl
                            .Left = ctrl.Left - 7
                            .Height = ctrl2.Height
                            .Picture = ctrl2.Picture
                            .PicturePosition = fmPicturePositionCenter
                            .Width = ctrl2.Width
                            .Top = ctrl.Top
                            .ZOrder (1)
                            Exit For
                        End If
                    Next ctrl2
                    '___D
                    'Mengubah format textbox kategori big___B
                    For Each ctrl2 In Input_Form.Controls
                        If ctrl2.tag = frm_ctrl And InStr(ctrl2.tag, "big") Then
                            .ControlTipText = "hover_" & frm_ctrl
                            .Left = ctrl.Left - 7
                            .Height = ctrl2.Height
                            .Picture = ctrl2.Picture
                            .PicturePosition = fmPicturePositionCenter
                            .Width = ctrl2.Width
                            .Top = ctrl.Top
                            .ZOrder (1)
                            Exit For
                        End If
                    Next ctrl2
                    '___B
                    'Mengubah format textbox kategori small___S
                    For Each ctrl2 In Input_Form.Controls
                        If ctrl2.tag = frm_ctrl And InStr(ctrl2.tag, "small") Then
                            .ControlTipText = "hover_" & frm_ctrl
                            .Left = ctrl.Left - 7
                            .Height = ctrl2.Height
                            .Picture = ctrl2.Picture
                            .PicturePosition = fmPicturePositionCenter
                            .Width = ctrl2.Width
                            .Top = ctrl.Top
                            .ZOrder (1)
                            Exit For
                        End If
                    Next ctrl2
                    '___S
                End With
            End With
                '___LC
        End If
        '___TC
        'ctrl.ZOrder (0)
    Next ctrl
End Sub
