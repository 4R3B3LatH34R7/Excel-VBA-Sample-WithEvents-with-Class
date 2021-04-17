VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   972
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3192
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'for class
Dim LabelsAsButtons As Collection

Sub buttonClick_Handler(btnID As String, dowhat As doWhat_Enum) 'can be bytes
    If Not Me.ToggleButton1.Value Then
        If btnID <> "" Then
            If Right(btnID, 6) = "Button" Then
                Me.Controls(btnID).SpecialEffect = IIf(dowhat = mDn, fmSpecialEffectSunken, fmSpecialEffectFlat)
                Me.Controls(btnID).BorderStyle = IIf(dowhat = mUp, fmBorderStyleSingle, fmBorderStyleNone)
                Me.Label4.Caption = btnID
            End If
        End If
    End If
End Sub

Private Sub AddButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call buttonClick_Handler(btnID:="AddButton", dowhat:=mDn)
End Sub
Private Sub AddButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call buttonClick_Handler(btnID:="AddButton", dowhat:=mUp)
End Sub

Private Sub CreateButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call buttonClick_Handler(btnID:="CreateButton", dowhat:=mDn)
End Sub
Private Sub CreateButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call buttonClick_Handler(btnID:="CreateButton", dowhat:=mUp)
End Sub

Private Sub CancelButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call buttonClick_Handler(btnID:="CancelButton", dowhat:=mDn)
End Sub
Private Sub CancelButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call buttonClick_Handler(btnID:="CancelButton", dowhat:=mUp)
End Sub

Private Sub exitButton_Click()
    Me.Hide
End Sub

'____________________Class section___________________________
Private Sub ToggleButton1_Change() 'this is not really required, just to show that class event handling can be turned on/off
    If Me.ToggleButton1.Value Then
        Call instantiateClass 'could be called in the userform initialze event
        Me.ToggleButton1.Caption = "Class ON"
    Else
        Call de_instantiateClass 'not really necessary
        Me.ToggleButton1.Caption = "Class OFF"
    End If
    Me.ToggleButton1.BackColor = IIf(Me.ToggleButton1.Value, vbRed, vbGreen)
    Me.Label4.Caption = ""
End Sub
Private Sub UserForm_Initialize()
    Call instantiateClass 'this is important
    Me.ToggleButton1.BackColor = IIf(Me.ToggleButton1.Value, vbRed, vbGreen)
End Sub
Sub instantiateClass()
Dim oneLabelAsButton As clsLabel 'declare a variable as label class
Dim oneControl As Control
    Set LabelsAsButtons = New Collection 'collection to save selected labels to be handled by class event handlers

    For Each oneControl In Me.Controls
        If TypeName(oneControl) = "Label" Then
            If Right(oneControl.Name, 6) = "Button" Then 'just to add selected labels,if all labels to be used, no need for this
                Set oneLabelAsButton = New clsLabel 'setting the declared variable as label class
                Set oneLabelAsButton.clsLabel = oneControl 'acutally set selected label control as class label
                LabelsAsButtons.Add oneLabelAsButton 'adding only selected labels to class collection so that their events will be handled by class event
            End If
        End If
    Next oneControl
End Sub
Sub de_instantiateClass()
    Set LabelsAsButtons = Nothing
End Sub

'barebone code
'Dim LabelsAsButtons As Collection
'Private Sub UserForm_Initialize()
'Dim oneLabelAsButton As clsLabel
'Dim oneControl As Control
'    Set LabelsAsButtons = New Collection
'    For Each oneControl In Me.Controls
'        If TypeName(oneControl) = "Label" Then
'            If Right(oneControl.Name, 6) = "Button" Then
'                Set oneLabelAsButton = New clsLabel
'                Set oneLabelAsButton.clsLabel = oneControl
'                LabelsAsButtons.Add oneLabelAsButton
'            End If
'        End If
'    Next oneControl
'End Sub

