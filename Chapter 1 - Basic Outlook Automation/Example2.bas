Attribute VB_Name = "Module2"
Option Explicit

Public Type TemailInfo
    Sender As String
    DateSent As Date
    DateReceived As Date
    subject As String
    AttachementCount As Integer
End Type

Private Sub displayMsgInfo(ByRef emailInfo As TemailInfo)
    With emailInfo
        MsgBox ("Message from " & .Sender & vbCrLf & _
            "- sent on " & .DateSent & vbCrLf & _
            "- received on " & .DateReceived & vbCrLf & _
            "Subject: " & .subject & vbCrLf & _
            .AttachementCount & " attachements")
    End With
End Sub


Private Function getEmailInfo(omail As Outlook.MailItem) As TemailInfo
    With omail
        getEmailInfo.Sender = .SenderName
        getEmailInfo.DateSent = .SentOn
        getEmailInfo.DateReceived = .ReceivedTime
        getEmailInfo.subject = .subject
        getEmailInfo.AttachementCount = .Attachments.Count
    End With
End Function

Private Function isOrderSlip(objAtt As Outlook.Attachment) As Boolean
    Dim attName As String: attName = objAtt.fileName
    Dim strPattern As String: strPattern = "new order\s*\d+\.csv"
    Dim regEx As New RegExp
        
    With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
    End With
    
    isOrderSlip = regEx.Test(attName)
End Function

Private Function getOrderNumber(objAtt As Outlook.Attachment) As Long
    Dim attName As String: attName = objAtt.fileName
    Dim strPattern As String: strPattern = "new order\s*(\d+)\.csv"
    Dim regEx As New RegExp
    Dim mcolResults As MatchCollection
    
    
    With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
    End With
        
    If regEx.Test(attName) Then
         Set mcolResults = regEx.Execute(attName)
         isOrderList = mcolResults(1)
    Else
        isOrderList = -1
    End If
End Function

Public Sub msgInfo()
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim omail As Outlook.MailItem
    Dim selObjectCtr As Long
    
    
    Dim attName As String
    Dim objAtt As Outlook.Attachment
    
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection

    For selObjectCtr = 1 To myOlSel.Count
        If myOlSel.Item(selObjectCtr).Class = OlObjectClass.olMail Then
            Set omail = myOlSel.Item(selObjectCtr)
                For Each objAtt In omail.Attachments
                    If isOrderList(objAtt) Then
                        Stop
                    End If
                Next objAtt
        End If
    Next selObjectCtr
End Sub

