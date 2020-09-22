Imports System.Text
Imports EPP.PublicClass
Class AcceptedOrder
    Implements iOrder
    Public Property Subject As String Implements iOrder.Subject
    Dim _dr As DataRow
    Dim cutoffdr As DataRow

    Sub New(ByVal dr As DataRow, ByVal cutoffdr As DataRow)
        ' TODO: Complete member initialization 
        _dr = dr
        Me.cutoffdr = cutoffdr
        Subject = "e-Staff Purchase System Notification - Accepted Order."
    End Sub
    Public Function BodyMessageOld() As String
        Dim sb As New StringBuilder
        Dim arrRows() As DataRow

        sb.Append(String.Format("Dear {0},{1}{1}", _dr.Item("employeename"), vbCrLf))
        sb.Append(String.Format("Order Status    : ***** Accepted *****{0}", vbCrLf))
        sb.Append(String.Format("Billing To      : {0}{1}", _dr.Item("billingto"), vbCrLf))
        sb.Append(String.Format("Billing To Name : {0}{1}{1}", _dr.Item("billingtoname"), vbCrLf))
        '------------------------123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                 1         2         3         4         5         6         7         8         9        10        11        12      
        sb.Append(String.Format("No. Product Id    Description                              Inquiry Qty  Confirmed Qty              Staff Price       Amount" & vbCrLf))
        sb.Append(vbCrLf)
        arrRows = _dr.GetChildRows("relHDDTL")
        Dim mytotal As Double = 0
        For i = 0 To arrRows.GetUpperBound(0)
            Dim row As DataRow = arrRows(i)
            sb.Append(String.Format("{7,2}. {0,-13} {1,-50} {2,-13}  {3,-13} {4,12} {5,12} {6}", row.Item("refno"), fixlength(row.Item("descriptionname"), 50), row.Item("qty"), row.Item("confirmedqty"), String.Format("{0:#,##0.00}", row.Item("staffprice")), String.Format("{0:#,##0.00}", row.Item("staffprice") * row.Item("confirmedqty")), vbCrLf, i + 1))
            mytotal = mytotal + (row.Item("staffprice") * row.Item("confirmedqty"))
        Next
        sb.Append(vbCrLf)
        sb.Append(String.Format("Total Amount : {0:#,##0.00}{1}{1}", mytotal, vbCrLf))

        sb.Append(String.Format("Delivery to office on coming Tuesday dated {1}.{0}", vbCrLf, String.Format("{0:MMM dd}", cutoffdr.Item("collectiondate"))))
        sb.Append(String.Format("Please submit your cheque on or before {1}.{0}{0}", vbCrLf, String.Format("{0:MMM dd}", cutoffdr.Item("chequesubmitdate"))))
        sb.Append(String.Format("Terms and Conditions :{0}", vbCrLf))
        sb.Append(String.Format("1.  Bounced cheque will be handled by Finance Department and charged for administration fee.{0}", vbCrLf))
        sb.Append(String.Format("2.  Unpaid item will not be delivered until payment settlement.{0}", vbCrLf))
        sb.Append(String.Format("3.  No cancellation of staff order is allowed.{0}{0}", vbCrLf))
        'sb.Append(String.Format("Have a nice shopping, {0}{0}", vbCrLf))
        sb.Append(String.Format("Thank you for shopping, {0}{0}", vbCrLf))
        sb.Append("e-Staff Purchase System Administrator.")

        Return sb.ToString
    End Function

    Public Function BodyMessage() As String Implements iOrder.BodyMessage
        Dim sb As New StringBuilder
        Dim arrRows() As DataRow
        sb.Append("<!DOCTYPE html><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head>")
        sb.Append("<style>")
        sb.Append("td {padding-left:10px;padding-right:10px;}")
        sb.Append("th {padding-left:10px;padding-right:10px;}")
        sb.Append(".right-align{text-align:right;}")
        sb.Append(".center-align{text-align:center;}")
        sb.Append(".defaultfont{font-size:11.0pt;")
        sb.Append("font-family:""Calibri"",""sans-serif"";}")
        sb.Append("</style>")
        sb.Append("<body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p>", _dr.Item("employeename")))
        sb.Append(String.Format("<table class=""defaultfont""><tr><td>Order Status</td><td>: <b>***** Accepted *****</b></td></tr><tr>"))
        sb.Append(String.Format("<td>Billing To</td><td>: {0}</td></tr><tr>", _dr.Item("billingto")))
        sb.Append(String.Format("<td>Billing To Name</td><td>: {0}</td></tr></table><br>", _dr.Item("billingtoname"), vbCrLf))
        '------------------------123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                 1         2         3         4         5         6         7         8         9        10        11        12      
        sb.Append(String.Format("<table class=""defaultfont"" border=0 cellspacing=0><tr><th>No.</th><th>Product Id</th><th>Description</th><th>Inquiry Qty</th><th>Confirmed Qty</th><th>Staff Price</th><th>Amount</th>"))
        arrRows = _dr.GetChildRows("relHDDTL")
        Dim mytotal As Double = 0
        For i = 0 To arrRows.GetUpperBound(0)
            Dim row As DataRow = arrRows(i)
            sb.Append(String.Format("<tr><td class=""right-align"">{7}.</td><td>{0}</td><td>{1}</td><td class=""right-align"">{2}</td><td class=""right-align"">{3}</td><td class=""right-align"">{4}</td><td class=""right-align"">{5}</td><td class=""right-align"">{6}</td></tr>", row.Item("refno"), row.Item("descriptionname"), row.Item("qty"), row.Item("confirmedqty"), String.Format("{0:#,##0.00}", row.Item("staffprice")), String.Format("{0:#,##0.00}", row.Item("staffprice") * row.Item("confirmedqty")), vbCrLf, i + 1))
            mytotal = mytotal + (row.Item("staffprice") * row.Item("confirmedqty"))
        Next
        sb.Append(vbCrLf)
        sb.Append(String.Format("<tr><td></td><td></td><td></td><td></td><td></td><td>Total Amount</td><td class=""right-align""> {0:#,##0.00}</td></tr></table>", mytotal, vbCrLf))

        'sb.Append(String.Format("<p>Delivery to office on coming Tuesday dated {1}.{0}", vbCrLf, String.Format("{0:MMM dd}", cutoffdr.Item("collectiondate"))))
        sb.Append(String.Format("<p>Delivery to office on coming {2} dated {1}.{0}", vbCrLf, String.Format("{0:MMM dd}", cutoffdr.Item("collectiondate")), String.Format("{0:dddd}", cutoffdr.Item("collectiondate"))))
        sb.Append(String.Format("<br>Please submit your check on or before {1}.{0}{0}</p>", vbCrLf, String.Format("{0:MMM dd}", cutoffdr.Item("chequesubmitdate"))))
        sb.Append(String.Format("<p>Terms and Conditions :<br>"))
        sb.Append(String.Format("1.  Bounced cheque will be handled by Finance Department and charged for administration fee.<br>"))
        sb.Append(String.Format("2.  Unpaid item will not be delivered until payment settlement.<br>"))
        sb.Append(String.Format("3.  No cancellation of staff order is allowed.<br>"))
        'sb.Append(String.Format("Have a nice shopping, {0}{0}", vbCrLf))
        sb.Append(String.Format("<p>Thank you for shopping,</p>"))
        sb.Append("<p>e-Staff Purchase System Administrator.</p>")

        Return sb.ToString

    End Function

    Private Function fixlength(ByVal p1 As String, ByVal p2 As Integer) As Object
        If p1.Length > p2 Then
            p1 = p1.Substring(0, p2)
        End If
        Return p1
    End Function

End Class

Class RejectedOrder
    Implements iOrder
    Public Property Subject As String Implements iOrder.Subject

    Dim _dr As DataRow
    Sub New(ByVal dr As DataRow)
        ' TODO: Complete member initialization 
        _dr = dr
        Subject = "e-Staff Purchase System Notification - Rejected Order."
    End Sub

    Public Function BodyMessageOld() As String
        Dim sb As New StringBuilder
        Dim arrRows() As DataRow

        sb.Append(String.Format("Dear {0},{1}{1}", _dr.Item("employeename"), vbCrLf))
        sb.Append(String.Format("Please be informed that your selected items is out of stock (or phase out), {0}{0}", vbCrLf))
        sb.Append(String.Format("Order Status    : ***** Rejected *****{0}", vbCrLf))
        sb.Append(String.Format("Billing To      : {0}{1}", _dr.Item("billingto"), vbCrLf))
        sb.Append(String.Format("Billing To Name : {0}{1}{1}", _dr.Item("billingtoname"), vbCrLf))
        '------------------------123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                 1         2         3         4         5         6         7         8         9        10        11        12      
        sb.Append(String.Format("No. Product Id    Description                              Inquiry Qty  Confirmed Qty              Staff Price       Amount" & vbCrLf))
        sb.Append(vbCrLf)
        arrRows = _dr.GetChildRows("relHDDTL")
        Dim mytotal As Double = 0
        For i = 0 To arrRows.GetUpperBound(0)
            Dim row As DataRow = arrRows(i)
            sb.Append(String.Format("{7,2}. {0,-13} {1,-50} {2,-13}  {3,-13} {4,12} {5,12} {6}", row.Item("refno"), fixlength(row.Item("descriptionname"), 50), row.Item("qty"), row.Item("confirmedqty"), String.Format("{0:#,##0.00}", row.Item("staffprice")), String.Format("{0:#,##0.00}", row.Item("staffprice") * row.Item("confirmedqty")), vbCrLf, i + 1))
            mytotal = mytotal + (row.Item("staffprice") * row.Item("confirmedqty"))
        Next
        sb.Append(vbCrLf)
        sb.Append(String.Format("Total Amount : {0:#,##0.00}{1}{1}", mytotal, vbCrLf))

        sb.Append(String.Format("Pleases kindly consider to re-order on next staff purchase.{0}{0}", vbCrLf))
        sb.Append(String.Format("Have a nice shopping,{0}{0}", vbCrLf))
        sb.Append("e-Staff Purchase System Administrator.")

        Return sb.ToString
    End Function

    Public Function BodyMessage() As String Implements iOrder.BodyMessage
        'Dim sb As New StringBuilder
        'Dim myret As String = String.Empty
        'Dim arrRows() As DataRow
        'arrRows = _dr.GetChildRows("relHDDTL")
        'For i = 0 To arrRows.GetUpperBound(0)
        '    Dim row As DataRow = arrRows(i)

        'Next
        'Return myret
        Dim sb As New StringBuilder
        Dim arrRows() As DataRow
        sb.Append("<!DOCTYPE html><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head>")
        sb.Append("<style>")
        sb.Append("td {padding-left:10px;padding-right:10px;}")
        sb.Append("th {padding-left:10px;padding-right:10px;}")
        sb.Append(".right-align{text-align:right;}")
        sb.Append(".center-align{text-align:center;}")
        sb.Append(".defaultfont{font-size:11.0pt;")
        sb.Append("font-family:""Calibri"",""sans-serif"";}")
        sb.Append("</style>")
        sb.Append("<body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p>", _dr.Item("employeename")))
        sb.Append(String.Format("<p>Please be informed that your selected items is out of stock (or phase out),</p>", vbCrLf))
        sb.Append(String.Format("<table class=""defaultfont""><tr><td>Order Status</td><td>: <b>***** Rejected *****</b></td></tr>"))
        sb.Append(String.Format("<tr><td>Billing To</td><td>: {0}</td></tr>", _dr.Item("billingto"), vbCrLf))
        sb.Append(String.Format("<tr><td>Billing To Name</td><td>: {0}</td></tr></table><br>", _dr.Item("billingtoname")))
        '------------------------123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                 1         2         3         4         5         6         7         8         9        10        11        12      
        sb.Append(String.Format("<table class=""defaultfont"" border=0 cellspacing=0><tr><th>No.</th><th>Product Id</th><th>Description</th><th>Inquiry Qty</th><th>Confirmed Qty</th><th>Staff Price</th><th>Amount</th></tr>"))
        arrRows = _dr.GetChildRows("relHDDTL")
        Dim mytotal As Double = 0
        For i = 0 To arrRows.GetUpperBound(0)
            Dim row As DataRow = arrRows(i)
            sb.Append(String.Format("<tr><td class=""right-align"">{7}.</td><td>{0}</td><td>{1}</td><td class=""right-align"">{2}</td><td class=""right-align"">{3}</td><td class=""right-align"">{4}</td><td class=""right-align"">{5}</td><td class=""right-align"">{6}</td></tr>", row.Item("refno"), row.Item("descriptionname"), row.Item("qty"), row.Item("confirmedqty"), String.Format("{0:#,##0.00}", row.Item("staffprice")), String.Format("{0:#,##0.00}", row.Item("staffprice") * row.Item("confirmedqty")), vbCrLf, i + 1))
            mytotal = mytotal + (row.Item("staffprice") * row.Item("confirmedqty"))
        Next
        sb.Append(vbCrLf)
        sb.Append(String.Format("<tr><td></td><td></td><td></td><td></td><td></td><td>Total Amount</td><td class=""right-align"">{0:#,##0.00}</td></tr></table>", mytotal))

        sb.Append(String.Format("<p>Pleases kindly consider to re-order on next staff purchase.</p>"))
        sb.Append(String.Format("<p>Have a nice shopping,</p>"))
        sb.Append("<p>e-Staff Purchase System Administrator.</p>")

        Return sb.ToString
    End Function
    Private Function fixlength(ByVal p1 As String, ByVal p2 As Integer) As Object
        If p1.Length > p2 Then
            p1 = p1.Substring(0, p2)
        End If
        Return p1
    End Function


        
End Class

Class ChequeReminder
    Implements iOrder

    Dim _dr As DataRow

    Sub New(ByVal dr As DataRow)
        ' TODO: Complete member initialization 
        _dr = dr
        Subject = "e-Staff Purchase System Notification - Submit Cheque Reminder."
    End Sub

    Public Function bodymessageold() As String
        Dim sb As New StringBuilder
        Dim arrRows() As DataRow

        sb.Append(String.Format("Dear {0},{1}{1}", _dr.Item("employeename"), vbCrLf))
        sb.Append(String.Format("According to our record, You have not submit your cheque. Please ignore this email if you have done it.{0}{0}", vbCrLf))

        sb.Append(String.Format("Billing To      : {0}{1}", _dr.Item("billingto"), vbCrLf))
        sb.Append(String.Format("Billing To Name : {0}{1}{1}", _dr.Item("billingtoname"), vbCrLf))
        '------------------------123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                 1         2         3         4         5         6         7         8         9        10        11        12      
        sb.Append(String.Format("No. Product Id    Description                              Inquiry Qty  Confirmed Qty              Staff Price       Amount" & vbCrLf))
        sb.Append(vbCrLf)
        arrRows = _dr.GetChildRows("relHDDTL")
        Dim mytotal As Double = 0
        For i = 0 To arrRows.GetUpperBound(0)
            Dim row As DataRow = arrRows(i)
            sb.Append(String.Format("{7,2}. {0,-13} {1,-50} {2,-13}  {3,-13} {4,12} {5,12} {6}", row.Item("refno"), fixlength(row.Item("descriptionname"), 50), row.Item("qty"), row.Item("confirmedqty"), String.Format("{0:#,##0.00}", row.Item("staffprice")), String.Format("{0:#,##0.00}", row.Item("staffprice") * row.Item("confirmedqty")), vbCrLf, i + 1))
            mytotal = mytotal + (row.Item("staffprice") * row.Item("confirmedqty"))
        Next
        sb.Append(vbCrLf)
        sb.Append(String.Format("Total Amount : {0:#,##0.00}{1}{1}", mytotal, vbCrLf))

        sb.Append(String.Format("Terms and Conditions :{0}", vbCrLf))
        sb.Append(String.Format("1.  Bounced cheque will be handled by Finance Department and charged for administration fee.{0}", vbCrLf))
        sb.Append(String.Format("2.  Unpaid item will not be delivered until payment settlement.{0}", vbCrLf))
        sb.Append(String.Format("3.  No cancellation of staff order is allowed.{0}{0}", vbCrLf))
        'sb.Append(String.Format("Have a nice shopping, {0}{0}", vbCrLf))
        sb.Append(String.Format("Thank you for shopping, {0}{0}", vbCrLf))
        sb.Append("e-Staff Purchase System Administrator.")

        Return sb.ToString
    End Function
    Public Function BodyMessage() As String Implements iOrder.BodyMessage
        Dim sb As New StringBuilder
        Dim arrRows() As DataRow
        sb.Append("<!DOCTYPE html><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head>")
        sb.Append("<style>")
        sb.Append("td {padding-left:10px;padding-right:10px;}")
        sb.Append("th {padding-left:10px;padding-right:10px;}")
        sb.Append(".right-align{text-align:right;}")
        sb.Append(".center-align{text-align:center;}")
        sb.Append(".defaultfont{font-size:11.0pt;")
        sb.Append("font-family:""Calibri"",""sans-serif"";}")
        sb.Append("</style>")
        sb.Append("<body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p>", _dr.Item("employeename")))
        sb.Append(String.Format("<p>According to our record, You have not submit your cheque. Please ignore this email if you have done it.</p>"))

        sb.Append(String.Format("<table class=""defaultfont""><tr><td>Billing To</td><td>: {0}</td></tr>", _dr.Item("billingto")))
        sb.Append(String.Format("<tr><td>Billing To Name</td><td>: {0}</td></tr></table><br>", _dr.Item("billingtoname")))
        '------------------------123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                 1         2         3         4         5         6         7         8         9        10        11        12      
        sb.Append(String.Format("<table class=""defaultfont"" border=0 cellspacing=0><tr><th>No.</th><th>Product Id</th><th>Description</th><th>Inquiry Qty</th><th>Confirmed Qty</th><th>Staff Price</th><th>Amount</th><th>" & vbCrLf))
        arrRows = _dr.GetChildRows("relHDDTL")
        Dim mytotal As Double = 0
        For i = 0 To arrRows.GetUpperBound(0)
            Dim row As DataRow = arrRows(i)
            sb.Append(String.Format("<tr><td class=""right-align"">{7}.</td><td>{0}</td><td>{1}</td><td class=""right-align"">{2}</td><td class=""right-align"">{3}</td><td class=""right-align"">{4}</td><td class=""right-align"">{5}</td><td class=""right-align"">{6}</td></tr>", row.Item("refno"), row.Item("descriptionname"), row.Item("qty"), row.Item("confirmedqty"), String.Format("{0:#,##0.00}", row.Item("staffprice")), String.Format("{0:#,##0.00}", row.Item("staffprice") * row.Item("confirmedqty")), vbCrLf, i + 1))
            mytotal = mytotal + (row.Item("staffprice") * row.Item("confirmedqty"))
        Next    
        sb.Append(String.Format("<tr><td></td><td></td><td></td><td></td><td></td><td>Total Amount</td><td class=""right-align"">{0:#,##0.00}<td></tr></table>", mytotal, vbCrLf))

        sb.Append(String.Format("<p>Terms and Conditions :<br>"))
        sb.Append(String.Format("1.  Bounced cheque will be handled by Finance Department and charged for administration fee.<br>"))
        sb.Append(String.Format("2.  Unpaid item will not be delivered until payment settlement.<br>"))
        sb.Append(String.Format("3.  No cancellation of staff order is allowed.</p>"))
        'sb.Append(String.Format("Have a nice shopping, {0}{0}", vbCrLf))
        sb.Append(String.Format("<p>Thank you for shopping, </p>"))
        sb.Append("<p>e-Staff Purchase System Administrator.</p>")

        Return sb.ToString
    End Function

    Public Property Subject As String Implements iOrder.Subject
    Private Function fixlength(ByVal p1 As String, ByVal p2 As Integer) As Object
        If p1.Length > p2 Then
            p1 = p1.Substring(0, p2)
        End If
        Return p1
    End Function
End Class


Class PickupItem
    Implements iOrder

    Dim _dr As DataRow
    Dim cutoffdr As DataRow
    Sub New(ByVal dr As DataRow, ByVal cutoffdr As DataRow)
        ' TODO: Complete member initialization 
        _dr = dr
        Me.cutoffdr = cutoffdr
        Subject = "e-Staff Purchase System Notification - Order Collection."
    End Sub
    'Public Function bodymessageold() As String
    '    Dim sb As New StringBuilder

    '    sb.Append(String.Format("Dear {0},{1}{1}", _dr.Item("employeename"), vbCrLf))
    '    sb.Append(String.Format("Your order is ready and please collect your order on following time slots.{0}{0}", vbCrLf))
    '    'sb.Append(String.Format("[{0:MMM dd}] 1. 12:00 - 13:00{1}", cutoffdr.Item("collectiondate"), vbCrLf))
    '    'sb.Append(String.Format("         2. 16:30 - 18.00{0}{0}", vbCrLf))
    '    'sb.Append(String.Format("[{0:MMM dd}] 1. 11:30 - 12:30{1}", cutoffdr.Item("collectiondate"), vbCrLf))
    '    'sb.Append(String.Format("         2. 16:00 - 17.30{0}{0}", vbCrLf))
    '    sb.Append(String.Format("[{0:MMM dd}] 1. 10:45 - 11:45{1}", cutoffdr.Item("collectiondate"), vbCrLf))
    '    sb.Append(String.Format("         2. 15:30 - 17.30{0}{0}", vbCrLf))
    '    'sb.Append(String.Format("Have a nice shopping, {0}{0}", vbCrLf))
    '    sb.Append(String.Format("Thank you for shopping, {0}{0}", vbCrLf))
    '    sb.Append("e-Staff Purchase System Administrator.")

    '    Return sb.ToString
    'End Function
    Public Function BodyMessage() As String Implements iOrder.BodyMessage
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head>")
        sb.Append("<style>")
        sb.Append("td {padding-left:10px;padding-right:10px;}")
        sb.Append("th {padding-left:10px;padding-right:10px;}")
        sb.Append(".right-align{text-align:right;}")
        sb.Append(".center-align{text-align:center;}")
        sb.Append(".defaultfont{font-size:11.0pt;")
        sb.Append("font-family:""Calibri"",""sans-serif"";}")
        sb.Append("</style>")
        sb.Append("<body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p>", _dr.Item("employeename")))
        sb.Append(String.Format("<p>Your order is ready and please collect your order on following time slots.</p>"))
        'sb.Append(String.Format("<p>Location: {0}.</p>", "Bangkok"))
        sb.Append(String.Format("<p>Location: {0}.</p>", DbAdapter1.GetLocation))
        'sb.Append(String.Format("<table class=""defaultfont""><tr><td>[{0:MMM dd}]</td><td>1. 12:00 - 13:00</td></tr>", cutoffdr.Item("collectiondate")))
        sb.Append(String.Format("<table class=""defaultfont""><tr><td>[{0:MMM dd}]</td><td>1. 11:30 - 12:30</td></tr>", cutoffdr.Item("collectiondate")))
        sb.Append(String.Format("<tr><td></td><td>2. 15:30 - 17.30</td></tr></table>", vbCrLf))
        'sb.Append(String.Format("Have a nice shopping, {0}{0}", vbCrLf))
        sb.Append(String.Format("<p>Thank you for shopping,</p>"))
        sb.Append("<p>e-Staff Purchase System Administrator.</p>")

        Return sb.ToString
    End Function

    Public Property Subject As String Implements iOrder.Subject
    Private Function fixlength(ByVal p1 As String, ByVal p2 As Integer) As Object
        If p1.Length > p2 Then
            p1 = p1.Substring(0, p2)
        End If
        Return p1
    End Function
End Class

Interface iOrder
    Property Subject As String
    Function BodyMessage() As String
End Interface