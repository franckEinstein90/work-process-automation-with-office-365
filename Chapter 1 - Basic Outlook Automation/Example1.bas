Attribute VB_Name = "Example1"
Option Explicit

Public Type TemailInfo
    Sender As String
    DateSent As Date
    DateReceived As Date
    Subject As String
    AttachementCount As Integer
End Type

Private Sub displayMsgInfo(ByRef emailInfo As TemailInfo)
    With emailInfo
        MsgBox ("Message from " & .Sender & vbCrLf & _
            "- sent on " & .DateSent & vbCrLf & _
            "- received on " & .DateReceived & vbCrLf & _
            "Subject: " & .Subject & vbCrLf & _
            .AttachementCount & " attachements")
    End With
End Sub


Private Function getEmailInfo(omail As Outlook.MailItem) As TemailInfo
    With omail
        getEmailInfo.Sender = .SenderName
        getEmailInfo.DateSent = .SentOn
        getEmailInfo.DateReceived = .ReceivedTime
        getEmailInfo.Subject = .Subject
        getEmailInfo.AttachementCount = .Attachments.Count
    End With
End Function

Public Sub msgInfo()
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim omail As Outlook.MailItem
    Dim selObjectCtr As Long
    Dim emailInfo As TemailInfo
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection

    For selObjectCtr = 1 To myOlSel.Count
        If myOlSel.Item(selObjectCtr).Class = OlObjectClass.olMail Then
            Set omail = myOlSel.Item(selObjectCtr)
            emailInfo = getEmailInfo(omail)
            Call displayMsgInfo(emailInfo)
        End If
    Next selObjectCtr
End Sub
