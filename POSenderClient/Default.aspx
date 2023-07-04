<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="POSenderClient._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>PO Sender Client</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Label ID="LblPONumber" runat="server">PO Number: </asp:Label>
        <asp:TextBox ID="TxtPONumber" runat="server"></asp:TextBox>
        <asp:Button ID="BtnClickMe" runat="server" Text="Submit" 
            onclick="BtnClickMe_Click" />
        
    </div>
    </form>
</body>
</html>
