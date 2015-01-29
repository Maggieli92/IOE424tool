<%@ page language="VB" MasterPageFile="MasterPageNotLoggedIn.master" %>
<script runat=server>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("user") = ""
        'Session("user_role") = ""

        ' Session("viewonly") = ""
    End Sub
    
    
    Protected Sub LogMe(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If useremailaddress.Text = "" Then
            LblMsg.Text = ""
            Err.visible = False
            Exit Sub
        End If
        Dim mySqlConnection As SqlConnection
        Dim mySqlCommand As SqlCommand
        Dim myReader As SqlDataReader
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand("select * from Users where username='" & useremailaddress.Text & "' and pw  = '" & pw.Text & "'", mySqlConnection)
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        If myReader.HasRows = True Then
            Do While myReader.Read
                Session("user") = myReader("FirstName") & " " & myReader("LastName")
                'Session("user_role") = myReader("UserRole")
                'Session("UID") = myReader("UID")
                Session("email") = myReader("username")
                'Session("sitecode") = myReader("SiteCode")
                'Session("company") = myReader("Company")
                'If IsDBNull(myReader("ParentCode")) = False Then
                '    Session("parentcode") = myReader("ParentCode")
                'Else
                '    Session("parentcode") = "No"
                'End If
                'If IsDBNull(myReader("ViewOnly")) = False Then
                '    Session("viewonly") = myReader("viewonly")
                'Else
                '    Session("viewonly") = "No"
                'End If
                
                Session.Timeout = 450
            Loop
        Else
            LblMsg.Text = "The userID/password combination you submitted was not recognized by the SDS Claims system. If you are a new user, please begin registration under the 'New User Access Request' area."
            ERR.Visible = True
            Exit Sub
        End If
        mySqlConnection.Close()
      
        
        
        Response.Redirect("Planograph.aspx")
      
    End Sub

    
    </script>
 <asp:content id="ContentPlaceHolder1" contentplaceholderid="ContentPlaceHolder1" runat="server">

 <table border="0" width="90%">
	 	<tr valign="top">
    <td align="left">  	<img src='images/cross.gif'  runat="server" id="ERR" visible=false align='absmiddle'> </td>
	<td align="left">		<asp:Label ID="LblMsg" runat="server" Text="" CssClass="ErrorMsg" >
			</asp:Label>
	
		</td>
		</tr></table>
 <table cellspacing="15">
<tr>
<td valign="top">
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/RegisteredUserLogin_01.jpg" width="15" height="57" /></td>
    <td><img src="images/RegisteredUserLogin_02.jpg" width="379" height="57" /></td>
    <td><img src="images/RegisteredUserLogin_03.jpg" width="17" height="57" /></td>
  </tr>
  <tr>
    <td><img src="images/RegisteredUserLogin_04.jpg" width="15" height="199" /></td>
    <td background="images/bluetile.jpg" bgcolor="2137a4">
    
    <table border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td class="LargeWhiteText"  align="center">User E-mail:&nbsp;</td>
        <td><asp:TextBox ID="useremailaddress" runat="server"  Width="170px" CssClass="textbox"></asp:TextBox></td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
    <tr>
        <td class="LargeWhiteText"  align="center">Password:&nbsp;</td>
        <td><asp:TextBox ID="pw" TextMode="Password" runat="server"  Width="170px" CssClass="textbox"></asp:TextBox></td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
     <tr>
        <td  colspan="2" align="center" ><asp:ImageButton ID="ImageButton1" runat="server" OnClick="LogMe" ImageUrl="images/Button_Login.jpg" CausesValidation="true" />      </td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
    <tr>
    <td colspan="2"></td>
    </tr>
    </table>&nbsp;</td>
    
    <td><img src="images/RegisteredUserLogin_06.jpg" width="17" height="199" /></td>
  </tr>
  <tr>
    <td><img src="images/RegisteredUserLogin_07.jpg" width="15" height="151" /></td>
    <td><img src="images/RegisteredUserLogin_08.jpg" width="379" height="151" /></td>
    <td><img src="images/RegisteredUserLogin_09.jpg" width="17" height="151" /></td>
  </tr>
</table>




&nbsp;

</td>

<td valign="top">












</td>
</tr>
</table>
Need help getting started?  <a href="ApplyingforAccess.html" class="Medium" target="_blank">Click here</a> to view a brief tutorial on Applying for Access.  <a href="LoggingIn.html" class="Medium" target="_blank">Click here</a> for a brief tutorial on Logging In.
  </asp:content>
