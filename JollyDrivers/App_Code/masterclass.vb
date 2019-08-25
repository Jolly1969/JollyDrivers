Imports ReadyPOP3
Imports System.Net.Mail
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Threading

Imports System.Drawing
Imports System.Drawing.Imaging


Public Class masterclass
    Inherits System.Web.UI.Page
    Public domain_name As String = System.Configuration.ConfigurationManager.AppSettings("domain_name").ToString()

    Public str_domain_customfield As String = "tbl_customfields_" & replacedotsforunderscore(System.Configuration.ConfigurationManager.AppSettings("domain_name").ToString())

    Public commencementdate As Date = Convert.ToDateTime(System.Configuration.ConfigurationManager.AppSettings("commencementdate"))
    Public googleAPI_AccessKey As String = System.Configuration.ConfigurationManager.AppSettings("googleAPI_AccessKey").ToString() '"AIzaSyDT1kfD4Miaf-nA_BRJDxI2SJaWEHsWhOI"
    Public timezoneid_preference As String = retrieve_timezoneid()
    Public business_name As String = System.Configuration.ConfigurationManager.AppSettings("business_name").ToString()

    Public uploadlocker As String = System.Configuration.ConfigurationManager.AppSettings("uploadlocker").ToString()
    Public invoicelocker As String = "C:\Inetpub\vhosts\jollygoodsolutions.com\httpdocs\locker\invoices\"
    Public smsglobal_userid As String = System.Configuration.ConfigurationManager.AppSettings("smsglobal_userid").ToString()
    Public smsglobal_pwd As String = System.Configuration.ConfigurationManager.AppSettings("smsglobal_pwd").ToString()

    Public smsglobal_master_id As String = "jgs0000"
    Public smsglobal_master_pwd As String = "jgs0000"

    Public Pop3_emailaddress As String = System.Configuration.ConfigurationManager.AppSettings("Pop3_emailaddress").ToString()
    Public Pop3_emailpassword As String = System.Configuration.ConfigurationManager.AppSettings("Pop3_emailpassword").ToString()
    Public Pop3_host As String = System.Configuration.ConfigurationManager.AppSettings("Pop3_host").ToString()
    Public default_countrycode As String = "61"
    Public str_gradient As String = "<div class='bottom-gradient add-top add-bottom'><span class='left'></span><span class='center'></span><span class='right'></span></div>"

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Dim myCookies As New System.Net.CookieContainer

    Function replacedotsforunderscore(tempstr As String) As String
        Dim newstr As String = Replace(tempstr, ".", "_")
        Return newstr

    End Function
    Function domainrecordexists(table_name As String) As Boolean
        Dim SQL As String = "SELECT COUNT(indexid) AS countofindexid FROM " & table_name & " WHERE (domain_name = @domain_name)"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim x As Integer = command.ExecuteScalar
        command = Nothing
        Return x
    End Function
    Function retrieve_timezoneid() As String
        Dim SQL As String = "SELECT timezoneid FROM tbl_domain_timezoneid WHERE domain_name=@domain_name"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim x As String = Convert.ToString(command.ExecuteScalar)
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x
    End Function
    Function convert_UTC_to_preftimezone(UTC_datetime As Date) As Date
        Dim cstZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(timezoneid_preference)
        Return TimeZoneInfo.ConvertTimeFromUtc(UTC_datetime, cstZone)
        'Return TimeZoneInfo.ConvertTimeBySystemTimeZoneId(UTC_datetime, "UTC", timezoneid_preference)
    End Function

    Function convert_servertime_to_UTC(server_datetime As Date) As Date
        Return TimeZoneInfo.ConvertTimeToUtc(server_datetime, TimeZoneInfo.Local)
        'Return TimeZoneInfo.ConvertTimeBySystemTimeZoneId(server_datetime, TimeZoneInfo.Local.Id, "UTC")
    End Function
    Sub Delay(ByVal dblSecs As Double)
        Dim dblWaitTil As Date
        dblWaitTil = DateAdd(DateInterval.Second, dblSecs, DateTime.Now)
        Do Until DateTime.Now > dblWaitTil
        Loop
    End Sub
    Public ReadOnly Property ConnString_production() As String
        Get
            Return ConfigurationManager.ConnectionStrings("jollygoodsolutions_production").ConnectionString
        End Get
    End Property
    Public ReadOnly Property ConnString() As String
        Get
            Return ConfigurationManager.ConnectionStrings("LocalSQLServer").ConnectionString
        End Get
    End Property
    Public ReadOnly Property ConnString_admin() As String
        Get
            Return ConfigurationManager.ConnectionStrings("jollygoodsolutions_admin").ConnectionString
        End Get
    End Property

    Public ReadOnly Property ConnString_business() As String
        Get
            Return ConfigurationManager.ConnectionStrings("jollygoodsolutions_business").ConnectionString
        End Get
    End Property

    Public ReadOnly Property ConnString_runtime() As String
        Get
            Return ConfigurationManager.ConnectionStrings("jollygoodsolutions_runtime").ConnectionString
        End Get
    End Property


   
    Function IsAlphaNumeric(ByVal sChr As String) As Boolean
        IsAlphaNumeric = sChr Like "[0-9A-Za-z]"
    End Function

    Function removelastvhar_ifnotalphanumeric(ByVal aStr As String) As String

        Dim lastchar As String = Right(aStr, 1)
        If Not IsAlphaNumeric(lastchar) Then
            aStr = Left(aStr, Len(aStr) - 1)
        End If
        Return aStr
    End Function
    Function getDefault_Pop3_emailaddress() As String
        Dim SQL As String = "SELECT pop3_emailaddress FROM tbl_emailaddresses WHERE domain_name=@domain_name AND isprimary=1"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim x As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x
    End Function
    Function getDefault_Pop3_emailpassword() As String
        Dim SQL As String = "SELECT pop3_password FROM tbl_emailaddresses WHERE domain_name=@domain_name AND isprimary=1"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim x As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x
    End Function
    Protected Sub sendEmail(ByVal recipient_address As String, ByVal cc_address As String, ByVal bcc_address As String, ByVal sender_address As String, ByVal subject As String, ByVal message_body As String, ByVal IsBodyHtml As Boolean)
        Dim mail As New MailMessage()
        mail.From = New MailAddress(sender_address)

        recipient_address = removelastvhar_ifnotalphanumeric(recipient_address)
        Dim to_array As String() = recipient_address.Split(",")
        For acounter As Integer = LBound(to_array) To UBound(to_array)
            mail.To.Add(LTrim(RTrim(to_array(acounter).ToString)))
        Next

        If cc_address <> "" Then
            cc_address = removelastvhar_ifnotalphanumeric(cc_address)
            Dim cc_array As String() = cc_address.Split(",")
            For acounter As Integer = LBound(cc_array) To UBound(cc_array)
                mail.CC.Add(LTrim(RTrim(cc_array(acounter).ToString)))
            Next
        End If

        If bcc_address <> "" Then
            bcc_address = removelastvhar_ifnotalphanumeric(bcc_address)
            Dim bcc_array As String() = bcc_address.Split(",")
            For bcounter As Integer = LBound(bcc_array) To UBound(bcc_array)
                mail.Bcc.Add(LTrim(RTrim(bcc_array(bcounter).ToString)))
            Next
        End If

        mail.Subject = subject
        mail.Body = message_body
        mail.IsBodyHtml = IsBodyHtml
        'send the message
        Dim smtp As New SmtpClient()
        Dim basicAuthenticationInfo As New System.Net.NetworkCredential(Pop3_emailaddress, Pop3_emailpassword)
        smtp.UseDefaultCredentials = False
        smtp.Credentials = basicAuthenticationInfo

        smtp.Host = "localhost"
        smtp.Port = 25
        smtp.Send(mail)



    End Sub


    Protected Sub set_customfield_byliteralid(ByVal literal_id As String, ByVal html_text As String)
        Dim SQL As String = "UPDATE " & str_domain_customfield & " SET html_text=@html_text WHERE literal_id=@literal_id"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@literal_id", SqlDbType.VarChar).Value = literal_id
        command.Parameters.Add("@html_text", SqlDbType.VarChar).Value = html_text
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub
    Function get_customfield_byliteralid(ByVal literal_id As String) As String
        Dim SQL As String = "SELECT html_text FROM " & str_domain_customfield & " WHERE literal_id=@literal_id"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@literal_id", SqlDbType.VarChar).Value = literal_id
        conn.Open()
        Dim x As String = Convert.ToString(command.ExecuteScalar)
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x
    End Function
    Function get_customfield_byliteralid_restorepoint(ByVal literal_id As String) As String
        Dim SQL As String = "SELECT html_restorepoint FROM " & str_domain_customfield & " WHERE literal_id=@literal_id"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@literal_id", SqlDbType.VarChar).Value = literal_id
        conn.Open()
        Dim x As String = Convert.ToString(command.ExecuteScalar)
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x
    End Function
    Protected Sub set_customfield_byliteralid_usingrestorepoint(ByVal literal_id As String)
        Dim SQL As String = "UPDATE " & str_domain_customfield & " SET html_text=html_restorepoint WHERE literal_id=@literal_id"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@literal_id", SqlDbType.VarChar).Value = literal_id
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Protected Sub set_customfield_byliteralid_restorepoint(ByVal literal_id As String, ByVal html_restorepoint As String)
        Dim SQL As String = "UPDATE " & str_domain_customfield & " SET html_restorepoint=@html_restorepoint WHERE literal_id=@literal_id"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@literal_id", SqlDbType.VarChar).Value = literal_id
        command.Parameters.Add("@html_restorepoint", SqlDbType.VarChar).Value = html_restorepoint
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub
    Function get_customfield_byliteralid_inc_prefix_suffix(ByVal literal_id As String) As String
        Dim SQL As String = "SELECT html_text,standard_prefix,standard_suffix FROM " & str_domain_customfield & " WHERE literal_id=@literal_id"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@literal_id", SqlDbType.VarChar).Value = literal_id
        conn.Open()
        Dim drR As SqlDataReader = command.ExecuteReader
        command = Nothing
        drR.Read()
        Dim html_text As String = Convert.ToString(drR.Item("html_text"))
        Dim standard_prefix As String = Convert.ToString(drR.Item("standard_prefix"))
        Dim standard_suffix As String = Convert.ToString(drR.Item("standard_suffix"))
        Dim tempstr As String = standard_prefix & html_text & standard_suffix
        drR.Close()

        conn.Close()
        conn = Nothing
        Return tempstr
    End Function

 

    Function generate_sparkheadermenu_account() As String
        Dim headerstr As New StringBuilder
        headerstr.AppendLine("<div class='row box dark'>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar1' runat='server' style='width:90%; background-color:#0066FF; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<center>")
        'headerstr.AppendLine("<label id='focusbar2' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        'headerstr.AppendLine("</center>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar3' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar4' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar5' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar6' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='account_details.aspx' class='button call-to-action'>Account Manager</a>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<a href='admin_underconstruction.aspx' class='button call-to-action'>System Administrator</a>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='content_logo.aspx'  class='button call-to-action'>Content Management</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='product_admin.aspx'  class='button call-to-action'>Products & <br />Orders</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='comms_sms_compose.aspx'  class='button call-to-action'>Email<br /> & SMS</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<a href='order_profile.aspx'  class='button call-to-action'>Order<br />Manager</a>")
        headerstr.AppendLine("</div>")


        headerstr.AppendLine("</div>")
        headerstr.AppendLine("</div>")
        Return headerstr.ToString
    End Function
    Function generate_sparkheadermenu_ordertracking() As String

        Dim headerstr As New StringBuilder
        headerstr.AppendLine("<div class='row box dark'>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar1' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<center>")
        'headerstr.AppendLine("<label id='focusbar2' runat='server' style='width:90%; background-color:#0066FF; height:5px'>&nbsp;</label>")
        'headerstr.AppendLine("</center>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar3' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar4' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar5' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar6' runat='server' style='width:90%; background-color:#0066FF; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='account_details.aspx' class='button call-to-action'>Account Manager</a>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<a href='admin_underconstruction.aspx' class='button call-to-action'>System Administrator</a>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='content_logo.aspx'  class='button call-to-action'>Content Management</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='product_admin.aspx'  class='button call-to-action'>Products & <br />Orders</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='comms_sms_compose.aspx'  class='button call-to-action'>Email<br /> & SMS</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<a href='order_profile.aspx'  class='button call-to-action'>Order<br />Manager</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")
        headerstr.AppendLine("</div>")
        Return headerstr.ToString
    End Function
    Function generate_sparkheadermenu_content() As String
        Dim headerstr As New StringBuilder
        headerstr.AppendLine("<div class='row box dark'>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar1' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<center>")
        'headerstr.AppendLine("<label id='focusbar2' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        'headerstr.AppendLine("</center>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar3' runat='server' style='width:90%; background-color:#0066FF; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar4' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar5' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar6' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='account_details.aspx' class='button call-to-action'>Account Manager</a>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<a href='admin_underconstruction.aspx' class='button call-to-action'>System Administrator</a>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='content_logo.aspx'  class='button call-to-action'>Content Management</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='product_admin.aspx'  class='button call-to-action'>Products & <br />Orders</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='comms_sms_compose.aspx'  class='button call-to-action'>Email<br /> & SMS</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<a href='order_profile.aspx'  class='button call-to-action'>Order<br />Manager</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")
        headerstr.AppendLine("</div>")
        Return headerstr.ToString
    End Function
    Function generate_sparkheadermenu_products() As String
        Dim headerstr As New StringBuilder
        headerstr.AppendLine("<div class='row box dark'>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar1' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<center>")
        'headerstr.AppendLine("<label id='focusbar2' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        'headerstr.AppendLine("</center>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar3' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar4' runat='server' style='width:90%; background-color:#0066FF; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar5' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar6' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='row remove-bottom'>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='account_details.aspx' class='button call-to-action'>Account Manager</a>")
        headerstr.AppendLine("</div>")

        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<a href='admin_underconstruction.aspx' class='button call-to-action'>System Administrator</a>")
        'headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='content_logo.aspx'  class='button call-to-action'>Content Management</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='product_admin.aspx'  class='button call-to-action'>Products & <br />Orders</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='comms_sms_compose.aspx'  class='button call-to-action'>Email<br /> & SMS</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<a href='order_profile.aspx'  class='button call-to-action'>Order<br />Manager</a>")
        headerstr.AppendLine("</div>")


        headerstr.AppendLine("</div>")
        headerstr.AppendLine("</div>")
        Return headerstr.ToString
    End Function
    Function generate_sparkheadermenu_comms() As String
        Dim headerstr As New StringBuilder
        ' headerstr.AppendLine("<img src='../images/Extras/Icons/black/wrench_plus_2_icon&24.png' />&nbsp;<span style='font-size:larger; font-weight:bolder'>SITE ADMININISTRATOR</span><br /><br />")
        headerstr.AppendLine("<div class='row box dark'>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        'headerstr.AppendLine("<div class='one column'>")
        'headerstr.AppendLine("<img src='../images/Extras/Icons/white/arrow_right_icon&16.png' />")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar1' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<center>")
        'headerstr.AppendLine("<label id='focusbar2' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        'headerstr.AppendLine("</center>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar3' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar4' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar5' runat='server' style='width:90%; background-color:#0066FF; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<center>")
        headerstr.AppendLine("<label id='focusbar6' runat='server' style='width:90%; background-color:#333333; height:5px'>&nbsp;</label>")
        headerstr.AppendLine("</center>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='row remove-bottom'>")
        'headerstr.AppendLine("<div class='one column'>")
        'headerstr.AppendLine("&nbsp;")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='account_details.aspx' class='button call-to-action'>Account Manager</a>")
        headerstr.AppendLine("</div>")
        'headerstr.AppendLine("<div class='three columns'>")
        'headerstr.AppendLine("<a href='admin_underconstruction.aspx' class='button call-to-action'>System Administrator</a>")
        'headerstr.AppendLine("</div>")
        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='content_logo.aspx'  class='button call-to-action'>Content Management</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='product_admin.aspx'  class='button call-to-action'>Products & <br />Orders</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='three columns'>")
        headerstr.AppendLine("<a href='comms_sms_compose.aspx'  class='button call-to-action'>Email<br /> & SMS</a>")
        headerstr.AppendLine("</div>")

        headerstr.AppendLine("<div class='four columns'>")
        headerstr.AppendLine("<a href='order_profile.aspx'  class='button call-to-action'>Order<br />Manager</a>")
        headerstr.AppendLine("</div>")


        headerstr.AppendLine("</div>")
        headerstr.AppendLine("</div>")
        Return headerstr.ToString
    End Function
    Function get_sparkmenu(ByVal selecteditem As String) As String

        Session("current_adminpage") = selecteditem


        Dim menustr As New StringBuilder
        Dim isFeatured As Boolean = False
        Select Case selecteditem

            Case "account_details.aspx", "admin_defaulttimezone.aspx", "account_invoices.aspx", "account_legal_termsandconditions.aspx", "account_legal_privacypolicy.aspx", "account_legal_legaldisclaimer.aspx", _
                "admin_underconstruction.aspx", "admin_membership.aspx", "admin_pagetitle.aspx", "admin_metadescription.aspx", "admin_metakeywords.aspx", _
                "admin_search_googleanalytics.aspx", "admin_search_xmlsitemap.aspx", "admin_search_robots.aspx", "admin_logo_favicon.aspx", _
                "support.aspx", "jollygoodsolutions_ourad.aspx"

                menustr.Append(generateSparkmenu_account_settings(selecteditem))
                menustr.Append(generateSparkmenu_sysadmin_general(selecteditem))
                menustr.Append(generateSparkmenu_sysadmin_pagetitle(selecteditem))
                menustr.Append(generateSparkmenu_sysadmin_support(selecteditem))

            Case "content_logo.aspx", "content_logo_small.aspx", "content_menutext.aspx", "content_newsticker.aspx", _
                 "content_section1_introduction.aspx", "content_section1_tagline.aspx", "content_section1_imageslider.aspx", "content_section1_thingstodo.aspx", _
                 "content_section1_about.aspx", "content_integration_opengraph.aspx", "content_integration_sharethis.aspx", "content_integration_googlemaps.aspx", _
                 "content_contactus.aspx", "content_features.aspx", "content_featureicons.aspx", "document_termsandconditions.aspx", _
                 "document_privacypolicy.aspx", "document_disclaimer.aspx", "document_frequentlyaskedquestions.aspx", "document_copyright.aspx"
                menustr.Append(generateSparkmenu_content_header(selecteditem))
                menustr.Append(generateSparkmenu_content_section1(selecteditem))
                menustr.Append(generateSparkmenu_content_integration(selecteditem))
                menustr.Append(generateSparkmenu_content_other(selecteditem))


            Case "product_images.aspx", "product_admin.aspx", "product_archives.aspx", "payment_bankdetails.aspx", "payment_paypaldetails.aspx", "payment_other.aspx", "product_orderformheader.aspx", _
                "order_confirmation.aspx", "order_statuses.aspx", "ordersummary_header.aspx", "document_termsofpurchase.aspx", "orderform_customfields.aspx"
                'menustr.Append(generateSparkmenu_content_services(selecteditem))
                menustr.Append(generateSparkmenu_content_products(selecteditem))
                menustr.Append(generateSparkmenu_content_orderprocess(selecteditem))
                menustr.Append(generateSparkmenu_content_payments(selecteditem))
                menustr.Append(generateSparkmenu_content_ordersummary(selecteditem))


            Case "comms_addressbook.aspx", "comms_sms_senderids.aspx", "comms_sms_alerts.aspx", "comms_sms_compose.aspx", "admin_defaultemailaddress.aspx", "comms_email_compose.aspx", "comms_email_draft.aspx", "comms_email_outbox.aspx", _
                "comms_sms_history.aspx", "comms_emailtemplates_createnew.aspx", "comms_emailtemplates_edit.aspx", "comms_emailtemplates_autofill.aspx"
                menustr.Append(generateSparkmenu_comms_addressbook(selecteditem))
                menustr.Append(generateSparkmenu_comms_sms(selecteditem))
                menustr.Append(generateSparkmenu_comms_email(selecteditem))
                menustr.Append(generateSparkmenu_comms_emailtemplates(selecteditem))

            Case "order_status_manager.aspx", "order_profile.aspx", "payment_notifications_ipn.aspx", "payment_notifications_bank.aspx", "payment_notifications_other.aspx"
                menustr.Append(generateSparkmenu_content_orders(selecteditem))
                menustr.Append(generateSparkmenu_content_paymentnotification(selecteditem))

        End Select

        Return menustr.ToString

    End Function
    Function generateSparkmenu_account_settings(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "account_details.aspx", "account_invoices.aspx", "account_legal_termsandconditions.aspx", "account_legal_privacypolicy.aspx", "account_legal_legaldisclaimer.aspx"
                menustr.AppendLine("<div class='four columns box light featured alpha selectedtop'>")
            Case Else
                menustr.AppendLine("<div class='four columns alpha'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/book_side_icon&48.png' />")
        menustr.AppendLine("<h4><a href='account_details.aspx' style='text-decoration: none'>Your Account</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage your account settings with Jolly Good Solutions!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "account_details.aspx", "Account Particulars", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "account_invoices.aspx", "Your Invoices", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "account_legal_termsandconditions.aspx", "Terms & Conditions", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "account_legal_privacypolicy.aspx", "Privacy Policy", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "account_legal_legaldisclaimer.aspx", "Legal Disclaimer", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    'Function generateSparkmenu_account_invoices(ByVal selecteditem As String) As String
    '    Dim menustr As New StringBuilder
    '    Select Case selecteditem
    '        Case "account_invoices.aspx"
    '            menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
    '        Case Else
    '            menustr.AppendLine("<div class='four columns omega'>")
    '    End Select
    '    menustr.AppendLine("<div class='headset price clearfix'>")
    '    menustr.AppendLine("<img src='../images/Extras/Icons/black/wallet_icon&48.png' />")
    '    menustr.AppendLine("<h4><a href='account_invoices.aspx' style='text-decoration: none'>Your Invoices</a></h4>")
    '    menustr.AppendLine("</div>")
    '    menustr.AppendLine(str_gradient)
    '    menustr.AppendLine("<p>View your Jolly Good Solutions Invoices here!</p>")
    '    menustr.AppendLine(createmenuitem(selecteditem, "account_invoices.aspx", "Your Invoices", ""))
    '    menustr.AppendLine("</div>")
    '    Return menustr.ToString
    'End Function
    'Function generateSparkmenu_account_legaldocuments(ByVal selecteditem As String) As String
    '    Dim menustr As New StringBuilder
    '    Select Case selecteditem
    '        Case "account_legal_termsandconditions.aspx", "account_legal_privacypolicy.aspx", "account_legal_legaldisclaimer.aspx"
    '            menustr.AppendLine("<div class='four columns box light featured selectedtop omega '>")
    '        Case Else
    '            menustr.AppendLine("<div class='four columns omega'>")
    '    End Select
    '    menustr.AppendLine("<div class='headset price clearfix'>")
    '    menustr.AppendLine("<img src='../images/Extras/Icons/black/doc_lines_stright_icon&48.png' />")
    '    menustr.AppendLine("<h4><a href='account_legal_termsandconditions.aspx' style='text-decoration: none'>Legal Documents</a></h4>")
    '    menustr.AppendLine("</div>")
    '    menustr.AppendLine(str_gradient)
    '    menustr.AppendLine("<p>View your Jolly Good Solutions Invoices here!</p>")
    '    menustr.AppendLine(createmenuitem(selecteditem, "account_legal_termsandconditions.aspx", "Terms & Conditions", ""))
    '    menustr.AppendLine(createmenuitem(selecteditem, "account_legal_privacypolicy.aspx", "Privacy Policy", ""))
    '    menustr.AppendLine(createmenuitem(selecteditem, "account_legal_legaldisclaimer.aspx", "Legal Disclaimer", ""))
    '    menustr.AppendLine("</div>")
    '    Return menustr.ToString
    'End Function

    Function generateSparkmenu_sysadmin_general(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "admin_underconstruction.aspx", "admin_defaulttimezone.aspx", "admin_membership.aspx", "admin_logo_favicon.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/wrench_plus_2_icon&48.png' />")
        menustr.AppendLine("<h4><a href='admin_defaulttimezone.aspx' style='text-decoration: none'>General Admin</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage general site settings here!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "admin_defaulttimezone.aspx", "Default Time Zone", ""))
        'menustr.AppendLine(createmenuitem(selecteditem, "admin_defaultemailaddress.aspx", "Default Email Address", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "admin_underconstruction.aspx", "Under Construction Page", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "admin_membership.aspx", "Manage Users", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "admin_logo_favicon.aspx", "Favicon Image", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_sysadmin_pagetitle(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "admin_pagetitle.aspx", "admin_metadescription.aspx", "admin_metadescription.aspx", "admin_metakeywords.aspx", _
            "admin_search_googleanalytics.aspx", "admin_search_xmlsitemap.aspx", "admin_search_robots.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/zoom_icon&48.png' />")
        menustr.AppendLine("<h4><a href='admin_pagetitle.aspx' style='text-decoration: none'>Search</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage your META Data here!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "admin_pagetitle.aspx", "HTML Page Title", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "admin_metadescription.aspx", "META Description Tag", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "admin_metakeywords.aspx", "META Keywords", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "admin_search_googleanalytics.aspx", "Google Analytics", ""))
        'menustr.AppendLine(createmenuitem(selecteditem, "admin_search_xmlsitemap.aspx", "XML Sitemap", ""))
        'menustr.AppendLine(createmenuitem(selecteditem, "admin_search_robots.aspx", "robots.txt", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_sysadmin_support(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "support.aspx", "jollygoodsolutions_ourad.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop'>")
            Case Else
                menustr.AppendLine("<div class='four columns'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../_layout/images/orb_48px.png' />")
        menustr.AppendLine("<h4><a href='admin_search_googleanalytics.aspx' style='text-decoration: none'>Support</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Need help - contact us anytime!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "support.aspx", "Support Tickets", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "jollygoodsolutions_ourad.aspx", "Our Ad Footer", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function

    Function generateSparkmenu_content_header(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "content_logo.aspx", "content_logo_small.aspx", "content_menutext.aspx", "content_newsticker.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop alpha'>")
            Case Else
                menustr.AppendLine("<div class='four columns alpha'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/home_icon&48.png' />")
        menustr.AppendLine("<h4><a href='content_logo.aspx' style='text-decoration: none'>Page Header</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage the Logo images, Menu & Section Headers and News Ticker!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "content_logo.aspx", "Header Logo Image", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_logo_small.aspx", "Mobile Logo Image", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_menutext.aspx", "Menu/Section Headers", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_newsticker.aspx", "News Ticker", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_content_section1(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "content_section1_tagline.aspx", "content_section1_introduction.aspx", "content_section1_imageslider.aspx", "content_section1_about.aspx", "content_section1_thingstodo.aspx", _
            "content_features.aspx", "content_featureicons.aspx", "content_contactus.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/layers_1_icon&48.png' />")
        menustr.AppendLine("<h4><a href='content_section1_introduction.aspx' style='text-decoration: none'>Main Content</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>This is the first section under the page header!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "content_section1_tagline.aspx", "Tag Line", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_section1_introduction.aspx", "Introduction", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_section1_imageslider.aspx", "Image Slider", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_section1_about.aspx", "About Us", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_featureicons.aspx", "Feature Icons", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_features.aspx", "Feature List", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_contactus.aspx", "Contact Us Options", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_content_integration(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "content_integration_googlemaps.aspx", "content_integration_opengraph.aspx", "content_integration_sharethis.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/share_icon&48.png' />")
        menustr.AppendLine("<h4><a href='content_integration_opengraph.aspx' style='text-decoration: none'>Integration</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage connections to external parties like Facebook, Google Maps, Twitter etc!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "content_integration_opengraph.aspx", "Open Graph Protocol", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_integration_sharethis.aspx", "Social Media Sharing", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "content_integration_googlemaps.aspx", "Google Maps", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_content_other(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "content_testimonials.aspx", "document_termsandconditions.aspx", "document_privacypolicy.aspx", "document_disclaimer.aspx", "document_frequentlyaskedquestions.aspx", "document_copyright.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop'>")
            Case Else
                menustr.AppendLine("<div class='four columns'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/balance_icon&48.png' />")
        menustr.AppendLine("<h4><a href='content_features.aspx' style='text-decoration: none'>Legal Documents</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Don't forget these!</p>")

        menustr.AppendLine(createmenuitem(selecteditem, "document_termsandconditions.aspx", "Terms & Conditions", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "document_privacypolicy.aspx", "Privacy Policy", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "document_disclaimer.aspx", "Legal Disclaimer", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "document_frequentlyaskedquestions.aspx", "FAQ", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "document_copyright.aspx", "Copyright", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function


    Function generateSparkmenu_content_orderprocess(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "product_orderformheader.aspx", "order_confirmation.aspx", "order_statuses.aspx", "orderform_customfields.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/doc_lines_icon&48.png' />")
        menustr.AppendLine("<h4><a href='product_orderformheader.aspx' style='text-decoration: none'>Order Submission</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Order Workflow & Form!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "order_statuses.aspx", "Order Statuses", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "product_orderformheader.aspx", "Order Form Header", ""))
        'menustr.AppendLine(createmenuitem(selecteditem, "orderform_customfields.aspx", "Order Form Custom Fields", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "order_confirmation.aspx", "Submit Confirmation Page", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function

    Function generateSparkmenu_content_orders(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "order_status_manager.aspx", "order_profile.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/calc_icon&48.png' />")
        menustr.AppendLine("<h4><a href='order_profile.aspx' style='text-decoration: none'>Review Orders</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Keeping track of your orders!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "order_profile.aspx", "Review an Order", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "order_status_manager.aspx", "Order Lists", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function

    Function generateSparkmenu_content_paymentnotification(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "payment_notifications_ipn.aspx", "payment_notifications_bank.aspx", "payment_notifications_other.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/attention_icon&48.png' />")
        menustr.AppendLine("<h4><a href='order_profile.aspx' style='text-decoration: none'>Payment<br>Notifications</a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Keeping track of payments!</p>")
        'menustr.AppendLine(createmenuitem(selecteditem, "payment_sms_alerts.aspx", "Payment Alerts (SMS)", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "payment_notifications_ipn.aspx", "Paypal IPN", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "payment_notifications_bank.aspx", "Bank Deposits", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "payment_notifications_other.aspx", "Other Payments", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function


    Function generateSparkmenu_content_payments(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "payment_paypaldetails.aspx", "payment_bankdetails.aspx", "payment_other.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/wallet_icon&48.png' />")
        menustr.AppendLine("<h4><a href='payment_bankdetails.aspx' style='text-decoration: none'>Payment Options<br /></a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage the Payment Options!</p>")

        menustr.AppendLine(createmenuitem(selecteditem, "payment_bankdetails.aspx", "Bank Account", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "payment_paypaldetails.aspx", "Paypal", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "payment_other.aspx", "Other Payment Options", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function

    Function generateSparkmenu_content_ordersummary(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "ordersummary_header.aspx", "document_termsofpurchase.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/doc_new_icon&48.png' />")
        menustr.AppendLine("<h4><a href='payment_bankdetails.aspx' style='text-decoration: none'>Order Summary<br /></a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Configure the Order Summary Page!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "ordersummary_header.aspx", "Page Header", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "document_termsofpurchase.aspx", "Terms of Purchase", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function

    Function generateSparkmenu_content_products(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "product_admin.aspx", "product_images.aspx", "product_archives.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/shop_cart_icon&48.png' />")
        menustr.AppendLine("<h4><a href='product_admin.aspx' style='text-decoration: none'>Products<br /></a></h4>")
        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Maintain your Product Portfolio!</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "product_images.aspx", "Product Images", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "product_admin.aspx", "Product List", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "product_archives.aspx", "Archived Products", ""))
        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function

    Function generateSparkmenu_comms_addressbook(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "comms_addressbook.aspx"
                menustr.AppendLine("<div class='four columns box light featured alpha selectedtop'>")
            Case Else
                menustr.AppendLine("<div class='four columns alpha'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/notepad_2_icon&48.png' />")
        menustr.AppendLine("<h4><a href='comms_addressbook.aspx' style='text-decoration: none'>Contacts</a></h4>")

        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Keeping track of important contacts for sms and email.</p>")

        menustr.AppendLine(createmenuitem(selecteditem, "comms_addressbook.aspx", "Address Book", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function


    Function generateSparkmenu_comms_sms(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "comms_sms_senderids.aspx", "comms_sms_alerts.aspx", "comms_sms_compose.aspx", "comms_sms_history.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop'>")
            Case Else
                menustr.AppendLine("<div class='four columns price'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/phone_icon&48.png' />")
        menustr.AppendLine("<h4><a href='comms_sms_senderids.aspx' style='text-decoration: none'>SMS</a></h4>")

        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage SMS here.</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "comms_sms_senderids.aspx", "Sender IDs", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_sms_alerts.aspx", "SMS Alerts", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_sms_compose.aspx", "Compose SMS", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_sms_history.aspx", "SMS Sent Items", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_comms_email(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "admin_defaultemailaddress.aspx", "comms_email_compose.aspx", "comms_email_draft.aspx", "comms_email_outbox.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop omega'>")
            Case Else
                menustr.AppendLine("<div class='four columns price omega'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/mail_icon&48.png' />")
        menustr.AppendLine("<h4><a href='comms_sms_senderids.aspx' style='text-decoration: none'>Email</a></h4>")

        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Your email settings.</p>")
        menustr.AppendLine(createmenuitem(selecteditem, "admin_defaultemailaddress.aspx", "Default Email", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_email_compose.aspx", "Compose Email", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_email_draft.aspx", "Draft Items", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_email_outbox.aspx", "Sent Items", ""))


        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function generateSparkmenu_comms_emailtemplates(ByVal selecteditem As String) As String
        Dim menustr As New StringBuilder
        Select Case selecteditem
            Case "comms_emailtemplates_autofill.aspx", "comms_emailtemplates_createnew.aspx"
                menustr.AppendLine("<div class='four columns box light featured selectedtop'>")
            Case Else
                menustr.AppendLine("<div class='four columns price'>")
        End Select
        menustr.AppendLine("<div class='headset price clearfix'>")
        menustr.AppendLine("<img src='../images/Extras/Icons/black/doc_new_icon&48.png' />")
        menustr.AppendLine("<h4><a href='comms_email_inbox.aspx' style='text-decoration: none'>Email Templates</a></h4>")

        menustr.AppendLine("</div>")
        menustr.AppendLine(str_gradient)
        menustr.AppendLine("<p>Manage your email templates here.</p>")

        menustr.AppendLine(createmenuitem(selecteditem, "comms_emailtemplates_autofill.aspx", "Autofill Tags", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_emailtemplates_createnew.aspx", "Create New Template", ""))
        menustr.AppendLine(createmenuitem(selecteditem, "comms_emailtemplates_edit.aspx", "Edit Email Templates", ""))

        menustr.AppendLine("</div>")
        Return menustr.ToString
    End Function
    Function createmenuitem(ByVal selecteditem As String, ByVal menuitem As String, ByVal menutext As String, ByVal alternativeitem As String) As String
        Dim menuitemstr As String = ""
        If (selecteditem = menuitem) Or (selecteditem = alternativeitem) Then
            menuitemstr = "<div class='row box dark remove-bottom'>&nbsp;<img src='../images/Extras/Icons/white/checkmark_icon&16.png' />&nbsp;&nbsp;<a href='#feature' style='text-decoration: none'><strong>" & menutext & "</strong><img src='../images/Extras/Icons/white/round_and_down_icon&16.png' style='float:right; margin-right:5px'/></a></div>"
        Else
            menuitemstr = "<div class='row box light remove-bottom'>&nbsp;<img src='../images/Extras/Icons/black/round_arrow_right_icon&16.png' />&nbsp;&nbsp;<a href='" & menuitem & "#feature' style='text-decoration: none'>" & menutext & "</a></div>"
            'menuitemstr = "<div class='row box light remove-bottom'>&nbsp;<img src='../images/Extras/Icons/black/round_arrow_right_icon&16.png' />&nbsp;&nbsp;<a onclick='location.href='" & menuitem & "'#feature' style='text-decoration: none'>" & menutext & "</a></div>"
        End If
        Return menuitemstr
    End Function



    Function evaluateInt(ByVal theInt As Integer) As Boolean
        If theInt = 1 Then Return True Else Return False
    End Function
    Function evaluateBln(ByVal thebln As Boolean) As Integer
        If thebln = True Then Return 1 Else Return 0
    End Function

    Function StripHTMLTags(ByVal HTMLToStrip As String) As String
        Dim stripped As String
        If HTMLToStrip <> "" Then
            stripped = Regex.Replace(HTMLToStrip, "<(.|\n)+?>", String.Empty)
            Return stripped
        Else
            Return ""
        End If
    End Function


    Public Function CheckURL(ByVal URL As String) As Boolean
        Dim result As Boolean = False
        Try
            Dim Response As Net.WebResponse = Nothing
            Dim WebReq As Net.HttpWebRequest = Net.HttpWebRequest.Create(URL)
            Response = WebReq.GetResponse
            Response.Close()
            result = True
        Catch ex As Exception
        End Try
        Return result
    End Function

    Public Function isUnderConstruction() As Boolean
        Dim SQL As String = "SELECT is_underconstruction FROM tbl_underconstruction WHERE domain_name=@domain_name"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim x As Integer = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return evaluateInt(x)

    End Function

    Function getlabel(ByVal fullstr As String) As String
        Dim startPos As Integer = InStr(fullstr, "(")
        Dim endPos As Integer = InStr(fullstr, ")")
        Dim lenStr As Integer = endPos - startPos - 1
        Dim newstr As String = Mid(fullstr, startPos + 1, lenStr)
        Return newstr
    End Function

    Function getnumber(ByVal fullstr As String) As String
        Dim stopPos As Integer = InStr(fullstr, "(")
        Dim newstr As String = Left(fullstr, stopPos - 1)
        Return newstr
    End Function

    Function calcSMSunits(ByRef WhatLen As Integer) As Integer
        Dim numunits As Integer
        Select Case WhatLen
            Case Is <= 160
                numunits = 1
            Case Is <= 306
                numunits = 2
            Case Is <= 459
                numunits = 3
            Case Is <= 612
                numunits = 4
            Case Is <= 765
                numunits = 5
            Case Is <= 918
                numunits = 6
            Case Is <= 1071
                numunits = 7
            Case Is <= 1224
                numunits = 8
            Case Is <= 1377
                numunits = 9
            Case Is <= 1530
                numunits = 10
        End Select
        Return numunits

    End Function
    Function truncateAt(ByRef NumUnits As Integer) As Integer
        Dim maxchars As Integer
        Select Case NumUnits
            Case 1
                maxchars = 160
            Case 2
                maxchars = 306
            Case 3
                maxchars = 459
            Case 4
                maxchars = 612
            Case 5
                maxchars = 765
            Case 6
                maxchars = 918
            Case 7
                maxchars = 1071
            Case 8
                maxchars = 1224
            Case 9
                maxchars = 1377
            Case 10
                maxchars = 1530
        End Select
        Return maxchars

    End Function

    Protected Sub generateSMS(ByVal theurl As String)
        Dim inStream As StreamReader
        Dim webRequest As WebRequest
        Dim webresponse As WebResponse
        webRequest = webRequest.Create(theurl)
        webresponse = webRequest.GetResponse()
        inStream = New StreamReader(webresponse.GetResponseStream())
        Dim ResponseString As String = inStream.ReadToEnd()
    End Sub


   
    Function getdocumentname_fromdocumentindexid(ByVal indexid As Long) As String
        Dim SQL As String = "SELECT     tbl_email_templates.template_name" & _
                            " FROM         tbl_documents INNER JOIN" & _
                            " tbl_email_templates ON tbl_documents.document_template_indexid = tbl_email_templates.indexid" & _
                            " WHERE     (tbl_documents.indexid = @indexid)"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.BigInt).Value = indexid
        conn.Open()
        Dim x As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x
    End Function
    Function get_autorespondstatus_byindexid(ByVal indexid As Integer) As Boolean
        Dim SQL As String = "SELECT template_currentstatus FROM tbl_email_templates WHERE indexid=@indexid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.VarChar).Value = indexid
        conn.Open()
        Dim x As Integer = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return evaluateInt(x)
    End Function

    Function get_autorespondstatus(ByVal autoreponsetemplate_code As String) As Boolean
        Dim SQL As String = "SELECT template_currentstatus FROM tbl_email_templates WHERE autoreponsetemplate_code=@autoreponsetemplate_code AND domain_name=@domain_name"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@autoreponsetemplate_code", SqlDbType.VarChar).Value = autoreponsetemplate_code
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim x As Integer = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return evaluateInt(x)
    End Function
    Function StripTags(ByVal html As String) As String
        ' Remove HTML tags.
        Return Regex.Replace(html, "<.*?>", "")
    End Function

    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        'http://www.vbforums.com/showthread.php?t=407441
        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If
        Return EmailAddressCheck

    End Function

 

    Protected Sub update_documentstatus(ByVal document_status_indexid As Integer, ByVal indexid As Integer)
        Dim SQL As String = "UPDATE tbl_documents SET document_status_indexid=@document_status_indexid, document_status_date=@document_status_date WHERE indexid=@indexid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.Int).Value = indexid
        command.Parameters.Add("@document_status_indexid", SqlDbType.VarChar).Value = document_status_indexid
        command.Parameters.Add("@document_status_date", SqlDbType.DateTime).Value = Date.UtcNow
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

  
   
    Protected Sub write_mail(ByVal email_date As Date, ByVal mailid As String, ByVal email_subject As String, ByVal email_from As String, ByVal email_to As String, ByVal email_messagebody As String, ByVal ishtml As Integer, ByVal email_htmlbody As String)

        If email_messagebody = Nothing Then email_messagebody = StripHTMLTags(email_htmlbody)
        Dim SQL As String = "INSERT INTO tbl_email_messages (domain_name,client_indexid,email_uniqueid,email_date,email_to,email_from,email_subject,email_messagebody,ishtml, email_htmlbody) VALUES (@domain_name,@client_indexid,@email_uniqueid,@email_date,@email_to,@email_from,@email_subject,@email_messagebody,@ishtml,@email_htmlbody)"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        command.Parameters.Add("@client_indexid", SqlDbType.Int).Value = 0
        command.Parameters.Add("@email_uniqueid", SqlDbType.VarChar).Value = mailid
        'using Date.UtcNow because sometimes the date that is returned is time only
        command.Parameters.Add("@email_date", SqlDbType.DateTime).Value = Date.UtcNow
        command.Parameters.Add("@email_to", SqlDbType.VarChar).Value = email_to
        command.Parameters.Add("@email_from", SqlDbType.VarChar).Value = email_from
        If email_subject = "" Then
            command.Parameters.Add("@email_subject", SqlDbType.VarChar).Value = DBNull.Value
        Else
            command.Parameters.Add("@email_subject", SqlDbType.VarChar).Value = email_subject
        End If
        command.Parameters.Add("@email_messagebody", SqlDbType.VarChar).Value = email_messagebody
        command.Parameters.Add("@ishtml", SqlDbType.TinyInt).Value = ishtml
        If ishtml = 1 Then
            command.Parameters.Add("@email_htmlbody", SqlDbType.VarChar).Value = email_htmlbody
        Else
            command.Parameters.Add("@email_htmlbody", SqlDbType.VarChar).Value = DBNull.Value
        End If
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing



        conn.Close()
        conn = Nothing
    End Sub


    Protected Sub sendsms_IPNreceived(ByVal sms_message As String)
        Dim primary_senderid As String = Left(domain_name, 11)
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim SQL As String = "SELECT * FROM tbl_sms_senderids WHERE domain_name=@domain_name AND (alert_paypalpayment=1) AND (archived=0)"
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        Dim drR As SqlDataReader = command.ExecuteReader
        Do While drR.Read
            Dim maxsmsunits As Integer = drR.Item("alert_maxunits")
            Dim maxchars As Integer = truncateAt(maxsmsunits)
            sms_message = Left(sms_message, maxchars)
            Dim recipientid As String = drR.Item("description") & " (" & drR.Item("mobile_number") & ")"
            Dim senderid As String = Left(domain_name, 11)
            Dim messageToPost As String = System.Web.HttpUtility.UrlEncode(sms_message, System.Text.Encoding.GetEncoding("ISO-8859-1"))
            Dim NumSplit As Integer = calcSMSunits(Len(sms_message))
            Dim request_string As String = "http://www.smsglobal.com.au/http-api.php?action=sendsms&user=" & smsglobal_userid & "&password=" & smsglobal_pwd & "&from=" & primary_senderid & "&to=" & recipientid & "&maxsplit=" & NumSplit & "&text=" & messageToPost
            Dim new_indexid As Long = create_sms_request(domain_name, "", 0, "Alert", request_string, sms_message, recipientid, NumSplit, "Payment Alert")
            Dim t As Thread
            t = New Thread(AddressOf Me.SendHTTPRequest)
            t.Start(new_indexid)
        Loop
        drR.Close()
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Protected Sub sendsms_onlinemessage(ByVal sms_message As String)
        Dim primary_senderid As String = Left(domain_name, 11)
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim SQL As String = "SELECT * FROM tbl_sms_senderids WHERE domain_name=@domain_name AND (alert_onlinemessage=1) AND (archived=0)"
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        Dim drR As SqlDataReader = command.ExecuteReader
        Do While drR.Read
            Dim maxsmsunits As Integer = drR.Item("alert_maxunits")
            Dim maxchars As Integer = truncateAt(maxsmsunits)
            sms_message = Left(sms_message, maxchars)
            Dim recipientid As String = drR.Item("description") & " (" & drR.Item("mobile_number") & ")"
            Dim senderid As String = Left(domain_name, 11)
            Dim messageToPost As String = System.Web.HttpUtility.UrlEncode(sms_message, System.Text.Encoding.GetEncoding("ISO-8859-1"))
            Dim NumSplit As Integer = calcSMSunits(Len(sms_message))
            Dim request_string As String = "http://www.smsglobal.com.au/http-api.php?action=sendsms&user=" & smsglobal_userid & "&password=" & smsglobal_pwd & "&from=" & primary_senderid & "&to=" & recipientid & "&maxsplit=" & NumSplit & "&text=" & messageToPost
            Dim new_indexid As Long = create_sms_request(domain_name, "", 0, "Alert", request_string, sms_message, recipientid, NumSplit, "Email Alert")
            Dim t As Thread
            t = New Thread(AddressOf Me.SendHTTPRequest)
            t.Start(new_indexid)
        Loop
        drR.Close()
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Protected Sub sendsms_orderreceived(ByVal sms_message As String)
        Dim primary_senderid As String = Left(domain_name, 11)
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim SQL As String = "SELECT * FROM tbl_sms_senderids WHERE domain_name=@domain_name AND (alert_onlineorder=1) AND (archived=0)"
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        Dim drR As SqlDataReader = command.ExecuteReader
        Do While drR.Read
            Dim maxsmsunits As Integer = drR.Item("alert_maxunits")
            Dim maxchars As Integer = truncateAt(maxsmsunits)
            sms_message = Left(sms_message, maxchars)
            Dim recipientid As String = drR.Item("description") & " (" & drR.Item("mobile_number") & ")"
            Dim senderid As String = Left(domain_name, 11)
            Dim messageToPost As String = System.Web.HttpUtility.UrlEncode(sms_message, System.Text.Encoding.GetEncoding("ISO-8859-1"))
            Dim NumSplit As Integer = calcSMSunits(Len(sms_message))
            Dim request_string As String = "http://www.smsglobal.com.au/http-api.php?action=sendsms&user=" & smsglobal_userid & "&password=" & smsglobal_pwd & "&from=" & primary_senderid & "&to=" & recipientid & "&maxsplit=" & NumSplit & "&text=" & messageToPost
            Dim new_indexid As Long = create_sms_request(domain_name, "", 0, "Alert", request_string, sms_message, recipientid, NumSplit, "Order Alert")
            Dim t As Thread
            t = New Thread(AddressOf Me.SendHTTPRequest)
            t.Start(new_indexid)
        Loop
        drR.Close()
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub



    Protected Sub clear_tempdir_contents()

        Dim myFile As String = "" ' file strings

        Dim YOUR_DIRECTORY As String = "~/upload/tempdir/"
        For Each myFile In Directory.GetFiles(Server.MapPath(YOUR_DIRECTORY))
            File.Delete(myFile)
        Next

    End Sub
   
    Function derive_smsunitbalance_fromaccountbalance(ByVal response_string As String) As Double
        If response_string = "error" Then
            Return -1
        Else
            Dim numstr As String = ""
            Dim lenstr As Integer = Len(response_string)
            Dim startnum As Integer = 0
            For xyz As Integer = 1 To lenstr
                If IsNumeric(Mid(response_string, xyz, 1)) Then
                    startnum = xyz
                    Exit For
                End If
            Next
            Do Until Not IsNumeric(Mid(response_string, startnum, 1))
                numstr = numstr & Mid(response_string, startnum, 1)
                startnum += 1
            Loop
            Dim numcredits As Long = CLng(numstr) / 2.4
            Return numcredits

            'BALANCE: 1175.9999999999998; USER: jgs0002

        End If
    End Function
    Function getSMSGlobal_accountbalance() As Integer

        Dim request_string As String = "http://www.smsglobal.com/balance-api.php?user=" & smsglobal_userid & "&password=" & smsglobal_pwd
        Dim retryCount As Integer
        Dim response_string As String = ""
Retry:
        Try
            Dim inStream As StreamReader
            Dim webRequest As HttpWebRequest
            Dim webresponse As WebResponse
            webRequest = System.Net.WebRequest.Create(request_string)
            webRequest.CookieContainer = myCookies
            'You can increment the timeout, but I think that is better set it to a short time and do more retries if it timeout, you must take that decision.
            'webRequest.Timeout = 300000 'default 100000 ms
            webresponse = webRequest.GetResponse()
            inStream = New StreamReader(webresponse.GetResponseStream())
            response_string = inStream.ReadToEnd()
            inStream.Dispose()
            webresponse.Close()
            webRequest = Nothing
            Return derive_smsunitbalance_fromaccountbalance(response_string)
        Catch ex As Exception
            If retryCount < 3 Then
                ' Retry one more time
                retryCount += 1
                GoTo Retry
            End If
            Return derive_smsunitbalance_fromaccountbalance("error")
        End Try
    End Function

    Protected Sub SendHTTPRequest(ByVal indexid As Long)

        Dim request_string As String = get_request_string_fromindexid(indexid)
        Dim retryCount As Integer
        Dim response_string As String = ""
Retry:
        Try
            Dim inStream As StreamReader
            Dim webRequest As HttpWebRequest
            Dim webresponse As WebResponse
            webRequest = System.Net.WebRequest.Create(request_string)
            webRequest.CookieContainer = myCookies
            'You can increment the timeout, but I think that is better set it to a short time and do more retries if it timeout, you must take that decision.
            'webRequest.Timeout = 300000 'default 100000 ms
            webresponse = webRequest.GetResponse()
            inStream = New StreamReader(webresponse.GetResponseStream())
            response_string = inStream.ReadToEnd()
            inStream.Dispose()
            webresponse.Close()
            webRequest = Nothing
            Call write_sms_response(response_string, indexid)
        Catch ex As Exception
            If retryCount < 3 Then
                ' Retry one more time
                retryCount += 1
                GoTo Retry
            End If
        End Try
    End Sub



    Function create_sms_request(domain_name As String, ByVal booking_uniqueid As String, ByVal client_indexid As Integer, ByVal username As String, ByVal request_string As String, ByVal message_body As String, ByVal request_recipient As String, ByVal numsplit As Integer, ByVal senderid As String) As Long
        Dim SQL As String = "INSERT INTO tbl_sms_requests (domain_name,booking_uniqueid,client_indexid,username,request_string,date_requested,request_recipient,message_body,numsplit,senderid) VALUES (@domain_name,@booking_uniqueid,@client_indexid,@username,@request_string,@date_requested,@request_recipient,@message_body,@numsplit,@senderid)"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        conn.Open()
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name

        If booking_uniqueid = "" Then
            command.Parameters.Add("@booking_uniqueid", SqlDbType.VarChar).Value = DBNull.Value
        Else
            command.Parameters.Add("@booking_uniqueid", SqlDbType.VarChar).Value = booking_uniqueid
        End If

        If client_indexid = 0 Then
            command.Parameters.Add("@client_indexid", SqlDbType.VarChar).Value = DBNull.Value
        Else
            command.Parameters.Add("@client_indexid", SqlDbType.VarChar).Value = client_indexid
        End If
        command.Parameters.Add("@username", SqlDbType.VarChar).Value = username
        command.Parameters.Add("@request_string", SqlDbType.VarChar).Value = request_string
        command.Parameters.Add("@date_requested", SqlDbType.DateTime).Value = DateTime.UtcNow
        command.Parameters.Add("@request_recipient", SqlDbType.VarChar).Value = request_recipient
        command.Parameters.Add("@message_body", SqlDbType.VarChar).Value = message_body
        command.Parameters.Add("@numsplit", SqlDbType.TinyInt).Value = numsplit
        command.Parameters.Add("@senderid", SqlDbType.VarChar).Value = senderid
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        Dim SQL2 As String = "Select @@Identity as new_indexid;"
        Dim command2 As New SqlCommand(SQL2, conn)
        Dim new_indexid As Long = command2.ExecuteScalar
        command2 = Nothing
        conn.Close()
        conn = Nothing
        Return new_indexid

    End Function

    Protected Sub archive_sms_request(ByVal indexid As Long)
        Dim SQL As String = "UPDATE tbl_sms_requests SET archived=1 WHERE indexid=@indexid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.BigInt).Value = indexid
        conn.Open()
        Dim x As String = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing

    End Sub

    Protected Sub write_sms_response(ByVal response_string As String, ByVal indexid As Long)
        Dim SMSGlobalQMsgID As String = ExtractNextString(response_string, "Sent queued message ID: ")
        Dim SMSGlobalMsgID As String = ExtractNextString(response_string, "SMSGlobalMsgID:")
        Dim SQL As String = "UPDATE tbl_sms_requests SET response_string=@response_string,SMSGlobalQMsgID=@SMSGlobalQMsgID,SMSGlobalMsgID=@SMSGlobalMsgID WHERE indexid=@indexid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.BigInt).Value = indexid
        command.Parameters.Add("@response_string", SqlDbType.VarChar).Value = response_string
        command.Parameters.Add("@SMSGlobalQMsgID", SqlDbType.VarChar).Value = SMSGlobalQMsgID
        command.Parameters.Add("@SMSGlobalMsgID", SqlDbType.VarChar).Value = SMSGlobalMsgID
        conn.Open()
        Dim x As String = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing

    End Sub
    Function get_request_string_fromindexid(ByVal indexid As Long) As String
        Dim SQL As String = "SELECT request_string FROM tbl_sms_requests WHERE indexid=@indexid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.BigInt).Value = indexid
        conn.Open()
        Dim x As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return x

    End Function
    Function ExtractNextString(ByVal WebResponse As String, ByVal BaseStr As String) As String
        'this will find any given string and then return the string block immediately following
        Dim tempstr As String = ""
        Dim xyz As Integer = 0
        Dim valChar As Char = ""
        Dim tempPos As Integer = InStr(WebResponse, BaseStr)

        If tempPos = 0 Then
            Return ""
        Else
            For xyz = tempPos + Len(BaseStr) To Len(WebResponse)
                valChar = Mid(WebResponse, xyz, 1)
                'ONLY ACEPT alphanumeric characters
                If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(valChar)) > 0 Then
                    tempstr = tempstr & valChar
                ElseIf xyz = tempPos + Len(BaseStr) Then  'make sure its not the first char whuch may be inadvertantly a space
                    'do nothing but continue for next loop
                Else
                    'finished end of block
                    Exit For
                End If
            Next
            Return Trim(tempstr)
        End If

    End Function

    Function get_datetime_from_donedate(ByVal donedate As String) As Date
        Dim theyear As Integer = 2000 + CInt(Left(donedate, 2))
        Dim themonth As Integer = CInt(Mid(donedate, 3, 2))
        Dim theday As Integer = CInt(Mid(donedate, 5, 2))
        Dim thehour As Integer = CInt(Mid(donedate, 7, 2))
        Dim theminute As Integer = CInt(Right(donedate, 2))
        Dim thedate As Date = DateSerial(theyear, themonth, theday)
        Dim thedatewithhour As Date = DateAdd(DateInterval.Hour, thehour, thedate)
        Dim thedatewithhourandminutes As Date = DateAdd(DateInterval.Minute, theminute, thedatewithhour)
        Return thedatewithhourandminutes
    End Function


    Function config_internationalsmsnumber(ByVal client_phone_mobile_countrycode As String, ByVal client_phone_mobile As String) As String
        If Left(client_phone_mobile, 1) = "0" Then client_phone_mobile = Right(client_phone_mobile, Len(client_phone_mobile) - 1)
        Dim recipientnumber As String = client_phone_mobile_countrycode & client_phone_mobile
        Return recipientnumber
    End Function



    Function createTablefromunorderedlist(ByVal html_text As String, ByVal color As String) As String
        html_text = html_text.Replace("<ul>", "<table style='margin-left:20px'>")
        html_text = html_text.Replace("</ul>", "</table><br />")
        'do same for ordered lists
        html_text = html_text.Replace("<ol>", "<table style='margin-left:20px'>")
        html_text = html_text.Replace("</ol>", "</table><br />")
        html_text = html_text.Replace("<li>", "<tr><td class='bullet_" & color & "'>&nbsp;</td><td>")
        html_text = html_text.Replace("</li>", "</td></tr>")
        Return html_text
    End Function

    Function makebullet_lightbulb(ByVal apoint As String) As String
        Dim tempstr As New StringBuilder
        tempstr.AppendLine("<table>")
        tempstr.AppendLine("<tr>")
        tempstr.AppendLine("<td class='lightbulb16_black'>&nbsp;</td>")
        tempstr.AppendLine("<td>" & apoint & "</td>")
        tempstr.AppendLine("</tr>")
        tempstr.AppendLine("</table>")
        Return tempstr.ToString
    End Function

    Protected Sub prepare_image_new(ByVal physicalfilepath As String, ByVal img_width As Integer, ByVal img_height As Integer, ByVal fileuploadpath_target As String)
        'Get the image.    
        Dim fullSizeImg As System.Drawing.Image
        Dim image_maxheight As Integer = Convert.ToInt16(img_height)
        Dim image_maxwidth As Integer = Convert.ToInt16(img_width)
        Dim image_filename As String = Path.GetFileName(physicalfilepath)
        fullSizeImg = System.Drawing.Image.FromFile(physicalfilepath)
        'Determine width and height of uploaded image
        Dim image_width_px As Single = fullSizeImg.PhysicalDimension.Width
        Dim image_height_px As Single = fullSizeImg.PhysicalDimension.Height
        fullSizeImg.Dispose()
        Dim scalepercent As Double = 1
        If image_width_px > image_maxwidth Then
            scalepercent = image_maxwidth / image_width_px
        End If
        ResizeImage(physicalfilepath, scalepercent)
        image_height_px = image_height_px * scalepercent

        'now need to check the new height - if it is still too big then resize again
        If image_height_px > image_maxheight Then
            scalepercent = image_maxheight / image_height_px
            ResizeImage(physicalfilepath, scalepercent)
        End If
        If File.Exists(fileuploadpath_target) Then File.Delete(fileuploadpath_target)
        File.Copy(physicalfilepath, fileuploadpath_target)

    End Sub
    Sub ResizeImage(ByVal physicalfilepath As String, ByVal percentResize As Double)
        'http://www.thedesilva.com/2010/01/resize-image-using-vb-net/
        Dim bm As New Bitmap(physicalfilepath)
        Dim width As Integer = bm.Width * percentResize 'new image width. 
        Dim height As Integer = bm.Height * percentResize  'new image height
        Dim thumb As New Bitmap(width, height)
        Dim g As Graphics = Graphics.FromImage(thumb)
        g.InterpolationMode = Drawing2D.InterpolationMode.High
        g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)
        g.Dispose()
        bm.Dispose()

        'image path.
        Dim fileext As String = Path.GetExtension(physicalfilepath)
        Select Case fileext
            Case ".png"
                thumb.Save(physicalfilepath, System.Drawing.Imaging.ImageFormat.Png)
            Case ".gif"
                thumb.Save(physicalfilepath, System.Drawing.Imaging.ImageFormat.Gif)
            Case ".jpg", ".jpeg"
                thumb.Save(physicalfilepath, System.Drawing.Imaging.ImageFormat.Jpeg)
            Case ".bmp"
                thumb.Save(physicalfilepath, System.Drawing.Imaging.ImageFormat.Bmp)
            Case ".ico"
                thumb.Save(physicalfilepath, System.Drawing.Imaging.ImageFormat.Icon)
            Case Else
                thumb.Save(physicalfilepath, System.Drawing.Imaging.ImageFormat.Gif)
        End Select

        thumb.Dispose()

    End Sub


    Function generate_newsticker() As String
        Dim newsticker_html As New StringBuilder
        Dim SQL As String = "SELECT * FROM tbl_newsticker WHERE domain_name=@domain_name AND include=1 ORDER BY orderid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim drR As SqlDataReader = command.ExecuteReader
        command = Nothing
        newsticker_html.AppendLine("<ul id='newsticker_1' class='newsticker'>")
        Do While drR.Read()
            Dim validflag As Boolean = True
            Dim startdate As String
            Dim enddate As String
            If Convert.ToString(drR.Item("startdate")) <> "" Then
                startdate = drR.Item("startdate")
                If CDate(startdate) > DateTime.Now.Date Then validflag = False
            End If
            If Convert.ToString(drR.Item("enddate")) <> "" Then
                enddate = drR.Item("enddate")
                If CDate(enddate) < DateTime.Now.Date Then validflag = False
            End If
            If validflag Then newsticker_html.AppendLine("<li>" & drR.Item("ticker_text") & "</li>")
        Loop
        newsticker_html.AppendLine("</ul>")
        drR.Close()

        command = Nothing
        conn.Close()
        conn = Nothing
        Return newsticker_html.ToString
    End Function


    Protected Sub create_matchstrings(ByVal order_uniqueid As String)
        'first refresh the table delete all records
        Call refresh_matchstrings()
        Call generate_matchstrings_fromorder(order_uniqueid)

    End Sub
    Protected Sub refresh_matchstrings()
        'first refresh the table delete all records
        Dim SQL As String = "DELETE FROM tbl_email_matchstrings"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub
    Protected Sub generate_matchstrings_fromorder(ByVal order_uniqueid As String)
        'cycle through all the values in the booking table and create match strings for each value for the documents
        'first remove existing matchstrings

        Dim SQL As String = "SELECT tbl_orders.* FROM tbl_orders WHERE order_uniqueid=@order_uniqueid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        conn.Open()
        Dim drR As SqlDataReader = command.ExecuteReader
        command = Nothing
        drR.Read()
        Dim order_submitted_datetime As Date = drR.Item("order_submitted_datetime")
        Dim order_yourname As String = drR.Item("order_yourname")
        Dim order_orgname As String = Convert.ToString(drR.Item("order_orgname"))
        Dim order_emailaddress As String = Convert.ToString(drR.Item("order_emailaddress"))
        Dim order_phonenumber As String = Convert.ToString(drR.Item("order_phonenumber"))
        Dim order_deliveryaddress As String = Convert.ToString(drR.Item("order_deliveryaddress"))
        Dim order_specialinstructions As String = Convert.ToString(drR.Item("order_specialinstructions"))
        Dim order_value As Double = drR.Item("order_value")
        Dim order_summary As String = drR.Item("order_summary")
        drR.Close()
        conn.Close()
        conn = Nothing
        ' now add each value as a match string
        Call insert_matchstring("{order_submitted_datetime}", Format(order_submitted_datetime, "Long Date"))
        Call insert_matchstring("{order_yourname}", order_yourname)
        Call insert_matchstring("{order_orgname}", order_orgname)
        Call insert_matchstring("{order_emailaddress}", order_emailaddress)
        Call insert_matchstring("{order_phonenumber}", order_phonenumber)
        Call insert_matchstring("{order_deliveryaddress}", order_deliveryaddress)
        Call insert_matchstring("{order_specialinstructions}", order_specialinstructions)
        Call insert_matchstring("{order_value}", FormatCurrency(order_value))
        Call insert_matchstring("{order_summary}", order_summary)

    End Sub

    Protected Sub insert_matchstring(ByVal match_string As String, ByVal Value_string As String)
        Dim SQL As String = "INSERT INTO tbl_email_matchstrings (match_string,Value_string) VALUES(@match_string,@Value_string)"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@match_string", SqlDbType.VarChar).Value = match_string
        command.Parameters.Add("@Value_string", SqlDbType.VarChar).Value = Value_string
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Function iterate_constants(ByVal rawstr As String) As String
        Dim SQL As String = "SELECT * FROM tbl_email_global_parameters WHERE domain_name=@domain_name;"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim drR As SqlDataReader = command.ExecuteReader
        command = Nothing
        Do While drR.Read
            Dim parameter_string As String = drR.Item("parameter_string")
            Dim parameter_value As String = Convert.ToString(drR.Item("parameter_value"))
            rawstr = rawstr.Replace(parameter_string, parameter_value)
        Loop
        drR.Close()
        conn.Close()
        conn = Nothing
        Return rawstr
    End Function

    Function iterate_documentinputs(ByVal rawstr As String) As String
        Dim SQL As String = "SELECT * FROM tbl_email_matchstrings;"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        conn.Open()
        Dim drR As SqlDataReader = command.ExecuteReader
        command = Nothing
        Do While drR.Read
            Dim match_string As String = drR.Item("match_string")
            Dim Value_string As String = Convert.ToString(drR.Item("Value_string"))
            If Value_string <> "" Then
                rawstr = rawstr.Replace(match_string, Value_string)
            End If
        Loop
        drR.Close()
        conn.Close()
        conn = Nothing
        Return rawstr
    End Function

    Function iterate_customtags(ByVal rawstr As String, ByVal order_uniqueid As String) As String
        Dim customtag As String = ""
        '1 "#link_onlineordersummary#"
        customtag = "#link_ordersummary#"
        Dim tag_string As String = "<a href='http://www." & domain_name & "/ordersummary.aspx?id=" & order_uniqueid & "' target='_blank'>'YOUR ORDER SUMMARY'</a>"
        rawstr = rawstr.Replace(customtag, tag_string)
        '2 #link_thewebsite#
        customtag = "#link_thewebsite#"
        tag_string = "<a href='http://www." & domain_name & "' target='_blank'>www." & domain_name & "</a>"
        rawstr = rawstr.Replace(customtag, tag_string)

        '3) #link_ordersummary#
        'customtag = "#link_ordersummary#"
        'tag_string = "<a href='http://www." & domain_name & "/ordersummary.aspx?id=" & order_uniqueid & "' target='_blank'>'ORDER PAYMENT INSTRUCTIONS'</a>"
        'rawstr = rawstr.Replace(customtag, tag_string)

        Return rawstr
    End Function

    Function get_default_emailaddress() As String
        Dim SQL As String = "SELECT pop3_emailaddress FROM tbl_emailaddresses WHERE domain_name=@domain_name and isPrimary=1"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim pop3_emailaddress As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return pop3_emailaddress
    End Function

    Protected Sub send_email_jollygoodsolutions_support(ByVal sendfrom As String, ByVal subject As String, ByVal messagebody As String)
        Dim recipient_address As String = "support@jollygoodsolutions.com"
        Dim sender_address As String = get_default_emailaddress()
        sendEmail(recipient_address, "", "", sender_address, subject, messagebody, False)
    End Sub

    Protected Sub update_orderstatus(ByVal order_uniqueid As String, ByVal order_status_indexid As Integer, ByVal status_comment As String, ByVal status_updatedby As String)
        Dim SQL As String = "INSERT INTO tbl_orders_status_diary (order_uniqueid,order_status_indexid,status_datetime,status_updatedby,status_comment) VALUES (@order_uniqueid,@order_status_indexid,@status_datetime,@status_updatedby,@status_comment)"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command.Parameters.Add("@order_status_indexid", SqlDbType.Int).Value = order_status_indexid
        command.Parameters.Add("@status_comment", SqlDbType.VarChar).Value = status_comment
        command.Parameters.Add("@status_datetime", SqlDbType.DateTime).Value = DateTime.UtcNow
        command.Parameters.Add("@status_updatedby", SqlDbType.VarChar).Value = status_updatedby
        conn.Open()
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing

        Dim SQL2 As String = "Select @@Identity as new_indexid;"
        Dim command2 As New SqlCommand(SQL2, conn)
        Dim order_currentstatus_indexid As Long = command2.ExecuteScalar
        command2 = Nothing

        Dim SQL3 As String = "UPDATE tbl_orders SET order_currentstatus_indexid=@order_currentstatus_indexid WHERE order_uniqueid=@order_uniqueid"
        Dim command3 As New SqlCommand(SQL3, conn)
        command3.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command3.Parameters.Add("@order_currentstatus_indexid", SqlDbType.BigInt).Value = order_currentstatus_indexid
        Dim x3 As Integer = command3.ExecuteNonQuery
        command3 = Nothing

        conn.Close()
        conn = Nothing
    End Sub

    Protected Sub addtocart(ByVal order_uniqueid As String, ByVal order_product_indexid As Integer, ByVal listprice_perunit As Double)
        Dim SQL As String = "INSERT INTO tbl_order_cart (domain_name,order_uniqueid,order_product_indexid,listprice_perunit) VALUES (@domain_name,@order_uniqueid,@order_product_indexid,@listprice_perunit)"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command.Parameters.Add("@order_product_indexid", SqlDbType.Int).Value = order_product_indexid
        command.Parameters.Add("@listprice_perunit", SqlDbType.Money).Value = listprice_perunit
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Protected Sub removefromcart(ByVal order_uniqueid As String, ByVal order_product_indexid As Integer)
        Dim SQL As String = "DELETE FROM tbl_order_cart WHERE (order_uniqueid=@order_uniqueid) AND (order_product_indexid=@order_product_indexid)"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command.Parameters.Add("@order_product_indexid", SqlDbType.Int).Value = order_product_indexid
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Function get_listprice_fromproductindexid(ByVal indexid As Integer) As Double
        Dim SQL As String = "SELECT product_price FROM tbl_products WHERE indexid=@indexid"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@indexid", SqlDbType.Int).Value = indexid
        Dim product_price As Double = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return product_price
    End Function

    Protected Sub update_ordervalue(ByVal order_uniqueid As String, ByVal order_value As Double)
        Dim SQL As String = "UPDATE tbl_orders SET order_value=@order_value WHERE order_uniqueid=@order_uniqueid"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command.Parameters.Add("@order_value", SqlDbType.Money).Value = order_value
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Protected Sub update_ordersummary(ByVal order_uniqueid As String, ByVal order_summary As String)
        Dim SQL As String = "UPDATE tbl_orders SET order_summary=@order_summary WHERE order_uniqueid=@order_uniqueid"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command.Parameters.Add("@order_summary", SqlDbType.VarChar).Value = order_summary
        Dim x As Integer = command.ExecuteNonQuery
        command = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Function get_currentstatus(ByVal order_uniqueid As String) As String
        Dim SQL As String = "SELECT tbl_orderstatus.order_status FROM tbl_orders INNER JOIN tbl_orders_status_diary ON tbl_orders.order_currentstatus_indexid = tbl_orders_status_diary.indexid INNER JOIN tbl_orderstatus ON tbl_orders_status_diary.order_status_indexid = tbl_orderstatus.indexid WHERE (tbl_orders.order_uniqueid = @order_uniqueid)"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        Dim order_status As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return order_status
    End Function
    Function get_currentstatus_comment(ByVal order_uniqueid As String) As String
        Dim SQL As String = "SELECT  tbl_orders_status_diary.status_comment FROM tbl_orders INNER JOIN tbl_orders_status_diary ON tbl_orders.order_currentstatus_indexid = tbl_orders_status_diary.indexid INNER JOIN tbl_orderstatus ON tbl_orders_status_diary.order_status_indexid = tbl_orderstatus.indexid WHERE (tbl_orders.order_uniqueid = @order_uniqueid)"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        Dim status_comment As String = Convert.ToString(command.ExecuteScalar)
        command = Nothing
        conn.Close()
        conn = Nothing
        Return status_comment
    End Function

    Function get_currentstatus_date(ByVal order_uniqueid As String) As Date
        Dim SQL As String = "SELECT  tbl_orders_status_diary.status_datetime FROM tbl_orders INNER JOIN tbl_orders_status_diary ON tbl_orders.order_currentstatus_indexid = tbl_orders_status_diary.indexid INNER JOIN tbl_orderstatus ON tbl_orders_status_diary.order_status_indexid = tbl_orderstatus.indexid WHERE (tbl_orders.order_uniqueid = @order_uniqueid)"
        Dim conn As New SqlConnection(ConnString_production)
        conn.Open()
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        Dim status_datetime As Date = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return status_datetime
    End Function
    Protected Sub requery_currentstatus(ByVal order_uniqueid As String)
        Dim SQL As String = "SELECT MAX(indexid) AS maxofindexid FROM tbl_orders_status_diary WHERE order_uniqueid = @order_uniqueid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        conn.Open()
        Dim order_currentstatus_indexid As Long = command.ExecuteScalar
        command = Nothing
        Dim SQL2 As String = "UPDATE tbl_orders SET order_currentstatus_indexid=@order_currentstatus_indexid  WHERE order_uniqueid = @order_uniqueid"
        Dim command2 As New SqlCommand(SQL2, conn)
        command2.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        command2.Parameters.Add("@order_currentstatus_indexid", SqlDbType.BigInt).Value = order_currentstatus_indexid
        Dim x2 As Integer = command2.ExecuteNonQuery
        command2 = Nothing
        conn.Close()
        conn = Nothing
    End Sub

    Function Get_customername_from_orderuniqueid(ByVal order_uniqueid As String) As String
        Dim SQL As String = "SELECT order_yourname FROM tbl_orders WHERE order_uniqueid=@order_uniqueid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@order_uniqueid", SqlDbType.VarChar).Value = order_uniqueid
        conn.Open()
        Dim order_yourname As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return order_yourname
    End Function

    Protected Sub create_sharethisbutton_html()
        Dim SQL As String = "SELECT button_code FROM tbl_sharethis_buttoncode WHERE domain_name=@domain_name AND include=1 ORDER BY orderid"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        command.Parameters.Add("@domain_name", SqlDbType.VarChar).Value = domain_name
        conn.Open()
        Dim drR As SqlDataReader = command.ExecuteReader
        command = Nothing
        Dim html_text As New StringBuilder
        Do While drR.Read()
            html_text.AppendLine(Convert.ToString(drR.Item("button_code")))
        Loop
        drR.Close()
        conn.Close()
        conn = Nothing

        Call set_customfield_byliteralid("lit_sharethis_buttoncode", html_text.ToString)

    End Sub

    Function get_orderdeleted_statusindexid() As Integer
        Dim SQL As String = "SELECT indexid FROM tbl_orderstatus WHERE system_status=99"
        Dim conn As New SqlConnection(ConnString_production)
        Dim command As New SqlCommand(SQL, conn)
        conn.Open()
        Dim indexid As String = command.ExecuteScalar
        command = Nothing
        conn.Close()
        conn = Nothing
        Return indexid
    End Function

End Class
