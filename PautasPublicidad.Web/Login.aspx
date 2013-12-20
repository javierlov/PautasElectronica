<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="PautasPublicidad.Web.Login" %>

<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="/Styles/Accendo.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%" height="50px" class="headerMedium">
        <tr>
            <td align="left">
                <table>
                    <tr>
                        <td>
                            <asp:Image runat ="server" ImageUrl ="~/Images/Accendo.png" height="30px"/>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="right" style="color: White;">
                <asp:Label ID="lblEmpresa" runat="server" Text="Sistema de Pautas Publicitarias"></asp:Label>
            </td>
        </tr>
    </table>
    <div style="width: 100%; text-align: center;">
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <div align="center">
            <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" Width="200px" 
                HeaderText="Iniciar Sesión" 
                CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                CssPostfix="Office2010Silver" EnableDefaultAppearance="False" 
                GroupBoxCaptionOffsetX="6px" GroupBoxCaptionOffsetY="-19px" 
                SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" 
                    PaddingTop="10px" />
                <HeaderImage Url="~/Images/Icons/entity16_1036.png">
                </HeaderImage>
                <HeaderStyle>
                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" 
                    PaddingTop="3px" />
                </HeaderStyle>
                <PanelCollection>
                    <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                        <table>
                            <tr>
                                <td>
                                    <asp:Label runat="server" Text="Empresa:" ID="Label7"></asp:Label>
                                </td>
                                <td>
                                    <asp:DropDownList runat="server" Width="195px" ID="ddlEmpresa1">
                                        <asp:ListItem Selected="True">Sprayette S.A.</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label runat="server" Text="Usuario:" ID="Label8"></asp:Label>
                                </td>
                                <td>
                                    <asp:TextBox runat="server" Width="191px" ID="txUserName"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label runat="server" Text="Contrase&#241;a:" ID="Label9"></asp:Label>
                                </td>
                                <td>
                                    <asp:TextBox runat="server" Width="191px" ID="txPassword" TextMode="Password">admin</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="right">
                                    <dx:ASPxLabel ID="lblMsg" runat="server" ForeColor="Red">
                                    </dx:ASPxLabel>
                                    <dx:ASPxButton ID="ASPxButton1" runat="server" 
                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                        CssPostfix="Office2010Silver" 
                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                        Text="Iniciar Sesión" OnClick="ASPxButton1_Click" Width="140px">
                                        <Image Url="~/Images/Icons/entity16_1036.png">
                                        </Image>
                                    </dx:ASPxButton>
                                    
                                </td>
                            </tr>
                        </table>
                    </dx:PanelContent>
                </PanelCollection>
            </dx:ASPxRoundPanel>
        </div>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
    </div>
    </form>
</body>
</html>
