<%@ page language="VB" MasterPageFile="MasterPage.master" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<script runat="server">
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        
        Dim session_name As Label
        session_name = Master.FindControl("Session_Name")
        session_name.Text = "Logged in as : " & UCase(Session("user")) & " "
        session_name.Visible = True
        
        Email.Text = Session("username")
        
    End Sub

   
    
    Sub UploadFile(ByVal sender As Object, ByVal e As System.EventArgs)
  
        Dim strSQL As String
      
        Dim mySqlConnection As SqlConnection
        Dim mySqlCommand As SqlCommand
        Dim myReader As SqlDataReader
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand("Select TOP 1 ID From FileUploads order by ID DESC", mySqlConnection)
        mySqlConnection.Open()
        
        myReader = mySqlCommand.ExecuteReader()
        
        If myReader.HasRows = True Then
        
            Do While myReader.Read()
                'Response.Write(myReader("DemandQty") & " - " & myReader("DemandMonth") & "<br>")
                ID.Text = myReader("ID")
            Loop
            
        End If
        mySqlConnection.Close()
        strSQL = "insert into FileUploads (Email,FileUpload,FileUploadDate) values ('" & Email.Text & "','" & ID.Text & FileUpload1.FileName & "', '" & Now() & "')"

        
        'mySqlCommand = New SqlCommand("insert into users (PeriodDecimalSystem,EmailAddress,surname,firstname,phone,site_code1,site_code2,site_code3,site_code4,site_code5,site_code6,site_code7,site_code8,site_code9,site_code10,site_code11,company,address,city,state,country,language,measurements,pwd,requestDate,zipcode,user_role,UserType) values ('" & PeriodDecimalSystem1.selectedValue & "','" & EmailAddress.Text & "','" & surname.Text & "','" & firstname.Text & "','" & phone.Text & "','" & site_c1 & "','" & site_code2.Text & "','" & site_code3.Text & "','" & site_code4.Text & "','" & site_code5.Text & "','" & site_code6.Text & "','" & site_code7.Text & "','" & site_code8.Text & "','" & site_code9.Text & "','" & site_code10.Text & "','" & site_code11.Text & "','" & company.Text & "','" & address.Text & "','" & city.Text & "','" & Mystate & "','" & country.Text & "','" & Language.Text & "','" & Measurements.Text & "','" & pwd.Text & "','" & Now() & "','" & zip.Text & "','Supplier','" & UserType.SelectedValue & "')", mySqlConnection)
        'mySqlCommand = New SqlCommand("insert into SCRs (GSDB,Company,Address,City,State,Zip) values ('" & Request.QueryString("GSDB") & "','" & Company.Text & "','" & Address.Text & "','" & City.Text & "','" & State.Text & "','" & Zip.Text & "','" & FirstName.Text & "','" & LastName.Text & "','" & Phone.Text & "','" & Email1.Text & "','" & Email2.Text & "','" & Email3.Text & "','" & Email4.Text & "','" & Email5.Text & "' )", mySqlConnection)
        mySqlCommand = New SqlCommand(strSQL, mySqlConnection)
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        mySqlConnection.Close()
        
        If FileUpload1.HasFile Then
            Try
                FileUpload1.SaveAs(Server.MapPath("FileUploads/" & ID.Text & FileUpload1.FileName))
            Catch ex As Exception
                Response.Write("file upload did not work")
            End Try
            'Exit Sub
        End If
        
        
        Dim mm As New MailMessage("admin@creativedatainc.com", "glafever@creativedatainc.com,efreeman@creativedatainc.com")
        ' mm.To = bt.CommandArgument
        ' mm.From = "admin@creativedatainc.com"
        mm.Subject = "New Planograph XLS Upload"
        mm.Body = "A new Planograph XLS file has been submitted to the Planograph website by " & Session("user") & " (" & Session("email") & ")."
        'mm.BodyFormat = Mail.MailFormat.Html
        mm.IsBodyHtml = True
        'Dim smtp As New SmtpClient
        'smtp.Send(mm)
        
        Dim sc As SmtpClient = New SmtpClient("localhost")
        ' sc.Credentials = New Net.NetworkCredential("efreeman@creativedatainc.com", "newcdi")
        sc.Send(mm)
        
        
        
        'Dim Email As MailMessage = _
        'New MailMessage("efreeman@creativedatainc.com", "efreeman@creativedatainc.com")
        'Email.IsBodyHtml = True
        'Email.Subject = "New Planograph Uplaod"
        'Email.Body = "There is a new Planograph XLS Upload"
        
        ''Email.Bcc.Add("efree@juno.com,glennlafever@yahoo.com")
        'Email.Priority = MailPriority.High
	
        'Dim sc As SmtpClient = New SmtpClient("localhost")
        'sc.Credentials = New Net.NetworkCredential("efreeman@creativedatainc.com", "newcdi")
        'sc.Send(Email)
        
        SuccessMessage.Visible = True
        
    End Sub
    

</script>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server"><center>

    <asp:Label ID="ID" runat="server" Visible="false"></asp:Label>

    <div>
        <asp:Image ID="Image1" runat="server" ImageUrl="images/UploadaPartsFile.jpg" /><br />
      
        

<center>
    <asp:Label ID="SuccessMessage" runat="server" Text="Your file has been uploaded successfully." Visible="false" Font-Size="Large" ForeColor="Red"></asp:Label>
<table width="90%"><tr><td colspan="2">
To run the Planograph tool with your own parts list, download the "<a href="PlanographTemplate.xls">PlanographTemplate.xls</a>" file.  Fill in all of the information for each part, using the row at the top of the file as a guide.  Then, browse to the file on your computer, and click on the "Upload File" button. 
</td></tr>
<tr><td></td><td>
    <asp:TextBox ID="Email" runat="server" Visible="false"></asp:TextBox></td></tr>
<tr><td>XLS File: </td><td>
    <asp:FileUpload ID="FileUpload1" runat="server" /></td></tr>
    <tr><td colspan="2" align="center"><asp:ImageButton ID="ImageButton2" runat="server" OnClick="UploadFile"  ImageUrl="images/UploadFile.jpg" CausesValidation="false" />  </td></tr>
</table>
    </center>
</asp:Content>
