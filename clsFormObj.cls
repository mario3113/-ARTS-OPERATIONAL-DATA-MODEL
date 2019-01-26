VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents xFormItem As MSForms.OptionButton
Attribute xFormItem.VB_VarHelpID = -1
Dim UF  'if we declare this as a userform the callback won't work!?
Public ClickCallback As String, MouseMoveCallback As String

    Function findUF(X As MSForms.Control)
        'This finds the userform containing X.  We cannot directly test _
         typename(...)="Userform" because a userform's typename is the name _
         of the userform, e.g., frmSelectFont or the default of Userform1, _
         Userform2, etc. _
         We cannot test Typeof Y is MSForms.Userform because it incorrectly _
         returns True for a frame! _
         So, we indirectly infer that an object is a userform if it is not _
         any of the others that could contain an embedded control.
        Dim Y
        Set Y = X.Parent
        Do While TypeName(Y) = "Frame" Or TypeName(Y) = "Page" _
                Or TypeName(Y) = "MultiPage"
            Set Y = Y.Parent
            Loop
        Set findUF = Y
        End Function
Property Set FormItem(uFormItem As MSForms.Control)
    Set xFormItem = uFormItem
    Set UF = findUF(uFormItem)
    End Property
Property Get FormItem() As MSForms.Control
    Set FormItem = xFormItem
    End Property
Private Sub xFormItem_Click()
    If ClickCallback <> "" Then _
        CallByName UF, ClickCallback, VbMethod, xFormItem
    End Sub

Private Sub xFormItem_MouseMove(ByVal Button As Integer, _
        ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If MouseMoveCallback <> "" Then _
        CallByName UF, MouseMoveCallback, VbMethod, _
            xFormItem, Button, Shift, X, Y
    End Sub
