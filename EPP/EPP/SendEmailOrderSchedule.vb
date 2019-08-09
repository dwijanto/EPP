Imports System.Text
Imports EPP.PublicClass

Public Class SendEmailOrderSchedule
    Inherits Email
    Public Property CutOffDate As DateTime
    Public Property CollectionDate As Date
    Public Property EmailNotification As Date
    Public Property ChequeSubmit As Date


    Public Sub New(ByVal Cutoffdate As DateTime, ByVal CollectionDate As Date, ByVal EmailNotification As Date, ByVal chequesubmit As Date, ByVal sender As String,
                   ByVal sendto As String)
        Me.CutOffDate = Cutoffdate
        Me.CollectionDate = CollectionDate
        Me.ChequeSubmit = chequesubmit
        Me.EmailNotification = EmailNotification
        Me.sendto = sendto
        Me.sender = sender
        'Me.cc = "liedwi@gmail.com;dwijanto@yahoo.com"
        initmessage()

    End Sub

   
    Private Sub initmessage()
        Dim sb As New StringBuilder
        Me.subject = String.Format("{0:MMMM} Staff Purchasing Schedule", CutOffDate)
        sb.Append("<!DOCTYPE html>")
        sb.Append("<html>")
        sb.Append("<head>")
        sb.Append("<meta charset=us-ascii />")
        sb.Append("</head>")
        sb.Append("<style>")
        sb.Append("td {padding-left:10px;padding-right:10px;}")
        sb.Append("th {padding-left:10px;padding-right:10px;}")
        sb.Append(".right-align{text-align:right;}")
        sb.Append(".center-align{text-align:center;}")
        sb.Append(".defaultfont{font-size:11.0pt;")
        sb.Append("font-family:""Calibri"",""sans-serif"";}")
        sb.Append("</style>")
        sb.Append("<body class=""defaultfont"">")
        sb.Append("<p>Dear All,</p>")
        sb.Append(String.Format("<p>Good Morning!<br>Following is the {0:MMMM} {0:yyyy} staff purchasing schedule.</p>", CutOffDate))
        sb.Append("<table class=""defaultfont"">")
        sb.Append(String.Format("<tr><td>Application cutoff date</td><td>: {0:MMM} {0:dd (dddd)}.</td></tr><tr><td>Order Notification mail</td><td>: {1:MMM dd (dddd)}.</td></tr><tr><td>Cheque submit deadline</td><td>: {2:MMM dd (dddd)}.</td></tr><tr><td>Goods Collection</td><td>: {3:MMM dd (dddd)} <u>{4}</u> .</td></tr><tr>", CutOffDate, EmailNotification, ChequeSubmit, CollectionDate, DbAdapter1.GetLocation))
        sb.Append("</table>")
        sb.Append("<p>As we have limited space to store the products, please respect the schedule to pick up your order.</p>")
        sb.Append("<p>Have a nice shopping.</p>")
        sb.Append("<p>Best Regards,</p>")
        sb.Append("<p>e-staff Purchase Administrator.</p>")
        sb.Append("</body>")
        sb.Append("</html>")

        body = sb.ToString
    End Sub

    Private Sub initmessageold()
        Dim sb As New StringBuilder
        Me.subject = String.Format("{0:MMMM} Staff Purchasing Schedule", CutOffDate)
        sb.Append("Dear All," & vbCrLf & vbCrLf)
        sb.Append(String.Format("Good Morning! Following is the {0:MMMM} {0:yyyy} staff purchasing schedule.", CutOffDate) & vbCrLf & vbCrLf)
        sb.Append(String.Format("Application cutoff date : {0:MMM} {0:dd} -> Order Notification mail : {1:MMM dd (dddd)} -> Cheque submit deadline {2:MMM dd (dddd)} -> Order collection: {3:MMM dd (dddd)}", CutOffDate, EmailNotification, ChequeSubmit, CollectionDate) & vbCrLf & vbCrLf)
        sb.Append("As we have limit space to store the products, please well schedule your pick up plan with appreciate." & vbCrLf & vbCrLf)
        sb.Append("Have a nice shopping." & vbCrLf & vbCrLf)
        sb.Append("Best Regards," & vbCrLf & vbCrLf)
        sb.Append("e-staff Purchase Administrator" & vbCrLf)
        body = sb.ToString
    End Sub
End Class
