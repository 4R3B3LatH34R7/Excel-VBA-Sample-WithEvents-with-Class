VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum doWhat_Enum 'not necessary, it could be byte, put here just to help coder typing
    mDn = 1
    mUp = 0
End Enum

Public WithEvents clsLabel As MSForms.Label 'this is important, after adding this, label events will be handled
Attribute clsLabel.VB_VarHelpID = -1

Private Sub clsLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call clsLabel_Click_Handler(dowhat:=mDn) 'handler code could be directly put here
End Sub

Private Sub clsLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call clsLabel_Click_Handler(dowhat:=mUp) 'handler code could be directly put here
End Sub

Private Sub clsLabel_Click_Handler(dowhat As doWhat_Enum) 'this is not really necessary, this could be inside each event
    UserForm1.Controls(clsLabel.Name).SpecialEffect = IIf(dowhat = mDn, fmSpecialEffectSunken, fmSpecialEffectFlat)
    UserForm1.Controls(clsLabel.Name).BorderStyle = IIf(dowhat = mUp, fmBorderStyleSingle, fmBorderStyleNone)
    UserForm1.Label4.Caption = clsLabel.Name + "_cls" 'should not be in the classmodule as it should be handled in the userform label mousedown event
End Sub

'barebone code
'Option Explicit
'Public WithEvents clsLabel As MSForms.Label
'Private Sub clsLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Call clsLabel_Click_Handler(dowhat:=1)
'End Sub
'Private Sub clsLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Call clsLabel_Click_Handler(dowhat:=0)
'End Sub
'Private Sub clsLabel_Click_Handler(dowhat As Byte)
'    UserForm1.Controls(clsLabel.Name).SpecialEffect = IIf(dowhat = 1, fmSpecialEffectSunken, fmSpecialEffectFlat)
'    UserForm1.Controls(clsLabel.Name).BorderStyle = IIf(dowhat = 0, fmBorderStyleSingle, fmBorderStyleNone)
'End Sub
