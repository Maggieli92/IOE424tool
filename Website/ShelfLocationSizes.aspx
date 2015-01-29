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
        NewLocationSize.Visible = False
        
    End Sub

    Sub ShowAddNewLocationSize(ByVal sender As Object, ByVal e As System.EventArgs)
        NewLocationSize.Visible = True
        
    End Sub
    
    Sub AddNewLocationSize(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim mySqlConnection As SqlConnection
        Dim strSQL As String
        Dim mySqlCommand As SqlCommand
        Dim myReader As SqlDataReader
        strSQL = "Insert INTO LocationSizes (LocationSizeCode"
        
        If NewOpeningWidth.Text <> "" Then
            strSQL = strSQL & ",OpeningWidth"
        End If
        If NewOpeningHeight.Text <> "" Then
            strSQL = strSQL & ",OpeningHeight"
        End If
        If NewOpeningDepth.Text <> "" Then
            strSQL = strSQL & ",OpeningDepth"
        End If
        If NewWhseCategory.Text <> "" Then
            strSQL = strSQL & ",WhseCategory"
        End If
   
        If NewWhseArea.Text <> "" Then
            strSQL = strSQL & ",WhseArea"
        End If
        If NewLocationName.Text <> "" Then
            strSQL = strSQL & ",LocationName"
        End If
       
        strSQL = strSQL & ") Values('" & NewLocationSizeCode.Text & "'"
        
        
        If NewOpeningWidth.Text <> "" Then
            strSQL = strSQL & ",'" & NewOpeningWidth.Text & "'"
        End If
        If NewOpeningHeight.Text <> "" Then
            strSQL = strSQL & ",'" & NewOpeningHeight.Text & "'"
        End If
        If NewOpeningDepth.Text <> "" Then
            strSQL = strSQL & ",'" & NewOpeningDepth.Text & "'"
        End If
        If NewWhseCategory.Text <> "" Then
            strSQL = strSQL & ",'" & NewWhseCategory.Text & "'"
        End If
        If NewWhseArea.Text <> "" Then
            strSQL = strSQL & ",'" & NewWhseArea.Text & "'"
        End If
        If NewLocationName.Text <> "" Then
            strSQL = strSQL & ",'" & NewLocationName.Text & "'"
        End If
      
        strSQL = strSQL & ")"
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand(strSQL, mySqlConnection)
        mySqlConnection.Open()
        mySqlCommand.ExecuteReader()
        mySqlConnection.Close()
    
        Response.Redirect("ShelfLocationSizes.aspx")
        
    End Sub
    

</script>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server"><center>

        <asp:Label ID="Email" runat="server" Visible="False"></asp:Label>


    <div>
        <asp:Image ID="Image1" runat="server" ImageUrl="images/PageTitleDefineShelfLocationSizes.jpg" /><br />
      
        

<center>
<table width="800px" class="LargeBlackText"><tr><td>To make updates to the Shelf Location Sizes, click on "Update" below. 
    <asp:LinkButton ID="LinkButton1" runat="server"  OnClick="ShowAddNewLocationSize" Text="Click here" Font-Size="Small">Click here</asp:LinkButton> to add a new Shelf Location Size.</td></tr></table>
<table id="NewLocationSize" runat="server" cellpadding="3" cellspacing="0">
<tr class="MediumRedText"><td colspan="8">Note: All fields are required.</td></tr>
<tr style="background-color:#000066" class="MediumWhiteText" align="center">
<td>Location Size Code</td>
<td>Opening Width</td>
<td>Opening Height</td>
<td>Opening Depth</td>
<td>Whse Category</td>
<td>Whse Area</td>
<td>Location Name</td>
<td></td>
</tr>
<tr>
<td>
    <asp:TextBox ID="NewLocationSizeCode" runat="server" Width="100"></asp:TextBox></td>
<td>
    <asp:TextBox ID="NewOpeningWidth" runat="server" Width="100"></asp:TextBox></td>
<td>
    <asp:TextBox ID="NewOpeningHeight" runat="server" Width="100"></asp:TextBox></td>
<td>
    <asp:TextBox ID="NewOpeningDepth" runat="server" Width="100"></asp:TextBox></td>
<td>
    <asp:TextBox ID="NewWhseCategory" runat="server" Width="100"></asp:TextBox></td>
<td>
    <asp:TextBox ID="NewWhseArea" runat="server" Width="100"></asp:TextBox></td>
<td>
    <asp:TextBox ID="NewLocationName" runat="server" ></asp:TextBox></td>
<td>
    <asp:LinkButton ID="LinkButton2" runat="server" OnClick="AddNewLocationSize" Font-Size="Small">Add</asp:LinkButton>   </td>
</tr>

</table>
    <asp:GridView ID="GridView1" runat="server" DataSourceID="SQLDataSource1" DataKeyNames="LocationSizeCode" AutoGenerateColumns="false" EmptyDataText="Currently there are no Location Sizes."  BorderColor="black" AlternatingRowStyle-BackColor="#CCCCCC"  HeaderStyle-BackColor="#000066" AllowSorting="true"  AlternatingRowStyle-ForeColor="DarkGray" RowStyle-ForeColor="White" RowStyle-BackColor="White">
     <Columns>
              
		

				<asp:BoundField DataField="LocationSizeCode" HeaderText="Location Size Code" ItemStyle-Width="150"  itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80"/>
				<asp:BoundField DataField="OpeningWidth" HeaderText="Opening Width" ItemStyle-Width="150"   itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80" />
				<asp:BoundField DataField="OpeningHeight" HeaderText="Opening Height"   ItemStyle-Width="150" itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80" />
				<asp:BoundField DataField="OpeningDepth" HeaderText="Opening Depth" ItemStyle-Width="150"   itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80" />
				<asp:BoundField DataField="WhseCategory" HeaderText="Whse Category" ItemStyle-Width="150"   itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80" />
				<asp:BoundField DataField="WhseArea" HeaderText="Whse Area" ItemStyle-Width="150"   itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80" />

				<asp:BoundField DataField="LocationName" HeaderText="Location Name" ItemStyle-Width="150"   itemstyle-horizontalalign="center" ItemStyle-ForeColor="Black" HeaderStyle-ForeColor="White" ControlStyle-Width="80" />

				<asp:templatefield headertext="Edit" HeaderStyle-ForeColor="White">
                  <ItemTemplate>
					<asp:linkbutton id="btnEdit" runat="server" causesvalidation="true" commandname="Edit"  text="Update" />
				</ItemTemplate>
				<edititemtemplate>
					<asp:linkbutton id="btnUpdate" runat="server" commandname="Update" text="Save" />
					<asp:linkbutton id="btnCancel" runat="server" causesvalidation="false" commandname="Cancel"	text="Cancel" />
					<asp:linkbutton id="btnDelete" runat="server" causesvalidation="false" commandname="Delete"	text="Delete" />

				</edititemtemplate>
			</asp:templatefield>
				
    </Columns>
    </asp:GridView></center>
         </ContentTemplate>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString1 %>" SelectCommand="Select OpeningWidth,OpeningHeight,OpeningDepth,WhseCategory,WhseArea,LocationSizeCode,LocationName From LocationSizes Order By LocationSizeCode"  UpdateCommand="update LocationSizes set LocationSizeCode=@LocationSizeCode,OpeningWidth=@OpeningWidth,OpeningHeight=@OpeningHeight,OpeningDepth=@OpeningDepth,WhseCategory=@WhseCategory,WhseArea=@WhseArea,LocationName=@LocationName where LocationSizeCode=@LocationSizeCode" DeleteCommand="Delete from LocationSizes where LocationSizeCode=@LocationSizeCode">
    
    <UpdateParameters>
		       <asp:FormParameter Type="string" Name="LocationSizeCode"/>
               <asp:FormParameter Type="string" Name="OpeningWidth" />
			    <asp:FormParameter Type="String" Name="OpeningHeight" />
			    <asp:FormParameter Type="String" Name="OpeningDepth" />
			    <asp:FormParameter Type="String" Name="WhseCategory" />
			    <asp:FormParameter Type="String" Name="WhseArea"  />
			       <asp:FormParameter Type="String" Name="LocationName"  />
         </UpdateParameters></asp:SqlDataSource>

 
    </div>

    </center>
</asp:Content>
