Attribute VB_Name = "bootsrtap_hover_form_control"
' Ini adalah script di module bernama hover_formating
Option Explicit

Public TextboxMouseUp As New Collection
Public ComboboxMouseUp As New Collection
Public TextboxEnter As New Collection
Public ComboboxEnter As New Collection

Sub Hover_Format(frm As MSForms.UserForm)
    Dim ctrl As Control
    Dim tb As hover_control
    Dim cb As hover_control

    Set TextboxMouseUp = New Collection
    Set ComboboxMouseUp = New Collection
    Set TextboxEnter = New Collection
    Set ComboboxEnter = New Collection

    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "TextBox" And InStr(ctrl.tag, "warna") > 0 Then
            Set tb = New hover_control
            Set tb.TextboxMouseUp = ctrl
            TextboxMouseUp.Add tb
        ElseIf TypeName(ctrl) = "ComboBox" And InStr(ctrl.tag, "warna") > 0 Then
            Set cb = New hover_control
            Set cb.ComboboxMouseUp = ctrl
            ComboboxMouseUp.Add cb
        End If
    Next ctrl

    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "TextBox" And InStr(ctrl.tag, "warna") > 0 Then
            Set tb = New hover_control
            Set tb.TextboxEnter = ctrl
            TextboxEnter.Add tb
        ElseIf TypeName(ctrl) = "ComboBox" And InStr(ctrl.tag, "warna") > 0 Then
            Set cb = New hover_control
            Set cb.ComboboxEnter = ctrl
            ComboboxEnter.Add cb
        End If
    Next ctrl
End Sub

