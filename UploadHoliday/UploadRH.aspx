<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadRH.aspx.cs" Inherits="UploadHoliday.UploadRH" %>

<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <telerik:RadScriptManager runat="server" ID="RadScriptManager1" />
        <div>
            <asp:Label ID="lblMessage" runat="server" Visible="false"></asp:Label>
        </div>
    <div>
        <telerik:RadUpload ID="UploadRHFile" runat="server" TargetFolder="~/File/" ControlObjectsVisibility="None"></telerik:RadUpload>
        &nbsp;
        <asp:Button ID="btnUpload" runat="server" Text="Import RH" OnClick="btnUpload_Click" />
    </div>
        <div>
            <asp:GridView ID="GridViewData" runat="server"></asp:GridView>
        </div>
    </form>
</body>
</html>
