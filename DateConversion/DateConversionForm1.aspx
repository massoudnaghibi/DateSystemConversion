<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DateConversionForm1.aspx.vb" Inherits="DateConversion.DateConversionForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
        </div>
        <asp:TextBox ID="InDateTextBox" runat="server"></asp:TextBox>
        &nbsp;<asp:RadioButtonList ID="RadioButtonList1" runat="server">
            <asp:ListItem Value="PI">Pers_Islamic</asp:ListItem>
            <asp:ListItem Value="PJ">Pers_Julian</asp:ListItem>
            <asp:ListItem Value="IP">Islamic_Pers</asp:ListItem>
            <asp:ListItem Value="IJ">Islamic_Julian</asp:ListItem>
            <asp:ListItem Value="JP">Julian_Pers</asp:ListItem>
            <asp:ListItem Value="JI">Julian_Islamic</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:RadioButtonList ID="RadioButtonList2" runat="server" Width="168px">
            <asp:ListItem Value="P">Pers</asp:ListItem>
            <asp:ListItem Value="I">Islamic</asp:ListItem>
            <asp:ListItem Value="J">Julian</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="OutDateLabel" runat="server"></asp:Label>
        &nbsp;
        <asp:Label ID="TodayDateLabel" runat="server" Text="تاریخ امروز"></asp:Label>
        <br />
        <asp:Button ID="DateConvButton" runat="server" Text="Convert Date" />
        &nbsp;&nbsp;&nbsp;
        <asp:Button ID="DateDescButton" runat="server" Text="Show Date Desc" />
    </form>
</body>
</html>
