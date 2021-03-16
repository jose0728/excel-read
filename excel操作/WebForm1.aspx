<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="excel操作.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="FileUploadStatus" runat="server" Text="Label"></asp:Label>
        </div>
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="上传" />
        <br />
        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="导入" />
    </form>
</body>
</html>
