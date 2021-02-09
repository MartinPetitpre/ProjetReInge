<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Projet._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    
            <asp:Button ID="Button1" runat="server" Text="Button" onclick="Button1_Click" 
            /><br />
            <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
            <asp:DataGrid id="Datagrid1" runat="server" />
	        <br />



            <center>
                <table cellpadding="3" cellspacing="3" border="3" bordercolor="green">
                    <tr>
                    <td colspan="2" align="center"><asp:button id="Button" Text="Générer_XML depuis fichier access " OnClick="access_Vers_XML" runat="server" /></td></tr>
                    <td colspan="2" align="center"><asp:button id="Button2" Text="Générer_XML depuis fichier excel " OnClick="excel_Vers_XML" runat="server" /></td></tr>
                </table>				
            </center>

    <h3>SQL Server</h3>
        <center>
            <br /><br /><br /><br /><br />
                <asp:Button ID="Button3" runat="server" Text="Afficher La Table" onclick="OpenSqlConnection"  /><br />
                    <asp:GridView ID="GridView1" runat="server"
                        DataSourceID="SqlDataSource1" AllowPaging="True" OnSelectedIndexChanged="GridView1_SelectedIndexChanged">
                        <AlternatingRowStyle BorderStyle="Double" BorderWidth="5px" />
                    </asp:GridView>
	            <asp:SqlDataSource ID="SqlDataSource1" runat="server"></asp:SqlDataSource>
	        <br />
        </center>

        
    

</asp:Content>
