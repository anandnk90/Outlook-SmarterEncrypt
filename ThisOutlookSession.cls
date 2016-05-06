VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Anand Nandakumar - 06-May-2016

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
   'SensitiveString is contains a list of keywords separated by commas
   Dim SensitiveString As Variant
   Dim Element As Variant
   Dim sFlag As Boolean
   
   sFlag = False
   
   SensitiveString = Array("National ID", "Emplid", "SSN", "Payroll ID", "GID", "Compensation", "Comp Rate")
   For Each Element In SensitiveString

    If InStr(1, Replace(Item.Body, " ", ""), Replace(Element, " ", ""), vbTextCompare) > 0 Then
        sFlag = True
    End If
    
   Next Element
   
   
   'If attachment exists and is greater than 1KB
   If Item.Attachments.Count > 0 Or sFlag = True Then
   
   'Taken from MS Support page for encrypting emails
   'Pops up box asking if you want to encrypt email
    If MsgBox("This email may have sensitive information. Do you want to encrypt before sending?", vbYesNo) = vbYes Then
      Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
      Dim oProp As Long
      oProp = CLng(Item.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS))
      Debug.Print "Original flag value is: " & oProp
      uFlags = 0
      ulFlags = ulFlags Or &H1 ' Add the encrypted flag
      ulFlags = ulFlags Or &H2 ' Add the signed flag
      Item.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
      Debug.Print "Updated flag value is: " & ulFlags
    End If
   
   End If
End Sub
