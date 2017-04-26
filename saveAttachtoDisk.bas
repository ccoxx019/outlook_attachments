Attribute VB_Name = "saveAttachtoDisk"
Public Sub saveAttachtoDisk(itm As Outlook.MailItem)
Dim objAtt As Outlook.Attachment
Dim saveFolder As String
saveFolder = "C:\Users\You\Documents\OneNote" ' put desired file location here
     For Each objAtt In itm.Attachments
          objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName
          Set objAtt = Nothing
     Next
End Sub

