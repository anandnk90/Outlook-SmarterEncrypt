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
   Dim SensitiveString As Variant
   Dim Element As Variant
   Dim sFlag, aFlag As Boolean
   Dim oProp As Long
   Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
   
   'SensitiveString is contains a list of keywords separated by commas
   SensitiveString = Array("National ID", "Emplid", "SSN", "Payroll ID", "GID", "Compensation", "Comp Rate")
   
   oProp = CLng(Item.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS))
   Debug.Print "Original flag value is: " & oProp
   
   uFlags = 0
   sFlag = False
   aFlag = False
   
   For Each Element In SensitiveString
    If (InStr(1, Replace(Item.Body, " ", ""), Replace(Element, " ", ""), vbTextCompare) > 0) Or (InStr(1, Replace(Item.Subject, " ", ""), Replace(Element, " ", ""), vbTextCompare) > 0) Then
        sFlag = True
    End If
   Next Element
   Debug.Print "sFlag value is: " & sFlag
   
   For Each Attachment In Item.Attachments
            If Attachment.Size > 512 Then
                aFlag = True
            End If
   Next Attachment
   Debug.Print "aFlag value is: " & aFlag
   
   If ((aFlag = True Or sFlag = True) And oProp <> 3) Then
    'Pops up box asking if you want to encrypt email
        If MsgBox("This email may have sensitive information. Do you want to encrypt before sending?", vbYesNo) = vbYes Then
            ulFlags = ulFlags Or &H1 ' Add the encrypted flag
            ulFlags = ulFlags Or &H2 ' Add the signed flag
            Item.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
            Debug.Print "Updated flag value is: " & ulFlags
        End If
   End If

End Sub
