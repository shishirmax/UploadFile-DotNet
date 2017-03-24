<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadRHList.aspx.cs" Inherits="UploadHoliday.UploadRHList" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="FileUploadToServer" runat="server" />
        &nbsp;
        <asp:Button ID="btnUpload" runat="server" Text="Upload RH List" OnClick="btnUpload_Click" />
        <br />
        <asp:Label ID="lblMsg" runat="server" Text=""></asp:Label>
        <br />
        <asp:GridView ID="GridView1" runat="server" EmptyDataText="No record found !" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Vertical">
            <AlternatingRowStyle BackColor="#DCDCDC" />
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <HeaderStyle BackColor="#000084" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#0000A9" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#000065" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
