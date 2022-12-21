<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BrokenLink.aspx.cs" Inherits="ClopayDoorBrokenLinks.BrokenLink" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Broken Links</title>
    <style>
        .pointer {cursor: pointer;Width:200px;height:50px;}
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div style="margin-top:200px;text-align:center;">
        <asp:Button ID="btnResidential" runat="server" Text="Get Residential Data" OnClick="btnResidential_Click" class="pointer"/>
        <br /> <br />
        <asp:Button ID="btnModel" runat="server" Text="Get Model Data" OnClick="btnModel_Click" class="pointer"/>
        <br /> <br />
        <asp:Button ID="btnCommercial" runat="server" Text="Get Commercial Data" OnClick="btnCommercial_Click" class="pointer"/>
        <br /> <br />
        <asp:Label ID="lblmsg" ForeColor="Red" runat="server"></asp:Label>
    </div>
    </form>
</body>
</html>
