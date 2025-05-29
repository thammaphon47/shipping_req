<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Home.aspx.vb" Inherits="SHIPPING_REQ.Home" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Welcome to Shipping System</title>
    <style>
body {
    font-family: Arial, sans-serif;
    background-image: url('');
    background-size: cover;
    background-position: center;
    background: #EBD3F8;
    margin: 0;
    padding: 0;
    display: flex;
    justify-content: center;
    align-items: center;
    height: 100vh;
    text-align: center;
}

.container {
    margin-top: 0;
    background-color: white;
    padding: 40px;
    border-radius: 15px;
    box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2), 0 4px 6px rgba(0, 0, 0, 0.1);
}

img.logo {
    width: 250px;
    height: auto;
    margin-bottom: 30px;
}

h1 {
    color: #333;
    margin-bottom: 10px;
}

.btn {
    display: inline-block;
    padding: 10px 25px;
    font-size: 18px;
    background-color: #EB3678;
    color: white;
    text-decoration: none;
    border-radius: 5px;
    margin-top: 20px;
}

.btn:hover {
    background-color: #005a9e;
}
</style>

</head>
<body>
<form id="form1" runat="server">
    <div class="container">
        <img src="ima/10112476.png" alt="Shipping Logo" class="logo" />
        <h1>Welcome to Shipping Request System</h1>
        <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label><br />

        <label for="txtUsername">Username:</label><br />
        <asp:TextBox ID="txtUsername" runat="server" /><br /><br />

        <label for="txtPassword">Password:</label><br />
        <asp:TextBox ID="txtPassword" runat="server" TextMode="Password" /><br /><br />

        <asp:Button ID="btnLogin" runat="server" Text="Login" OnClick="btnLogin_Click" CssClass="btn" />
    </div>
</form>

</body>
</html>
