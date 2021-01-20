<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Projet._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <<center><asp:Button ID="Button1" runat="server" Text="Button" onclick="Button1_Click" 
            /><br />
            <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
    <asp:DataGrid id="Datagrid1" runat="server" />
	<br />
        
        </center>

</asp:Content>
