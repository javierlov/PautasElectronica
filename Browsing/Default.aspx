<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:Button OnClientClick="showDirectory();" ID="Button1" runat="server" Text="Browse" /></div>
    </form>
    <script language=javascript>
    function showDirectory()
    {
       document.all.TextBox1.value=  window.showModalDialog("browseDirectory.aspx",'jain',"dialogHeight: 560px; dialogWidth: 360px; edge: Raised; center: Yes; help: Yes; resizable: Yes; status: No;");   
        return false;   
     }
    
    </script>
</body>
</html>
