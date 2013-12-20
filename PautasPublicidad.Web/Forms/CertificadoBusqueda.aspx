<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="CertificadoBusqueda.aspx.cs" Inherits="PautasPublicidad.Web.Forms.CertificadoBusqueda" %>

<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="uccombobox" TagPrefix="uc1" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <table width="100%">
        <tr>
            <td align="center">
                <asp:ScriptManager ID="ScriptManager1" runat="server">
                </asp:ScriptManager>
                <dx:ASPxRoundPanel ID="pnlMain" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                    CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                    GroupBoxCaptionOffsetY="-19px" HeaderText="Selección de Certificado" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                    Width="500px">
                    <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                    <ContentPaddings PaddingLeft="9px" PaddingTop="10px" PaddingRight="11px" PaddingBottom="10px">
                    </ContentPaddings>
                    <HeaderStyle>
                        <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                        <Paddings PaddingLeft="9px" PaddingTop="3px" PaddingRight="11px" PaddingBottom="6px">
                        </Paddings>
                    </HeaderStyle>
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server" SupportsDisabledAttribute="True">
                            <table width="100%">
                                <tr>
                                    <td>
                                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                            <ContentTemplate>
                                                <table width="100%">
                                                    <tr>
                                                        <td>
                                                            <table width="100%" style="border: 1px solid grey;">
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Label ID="Label3" runat="server" Text="Espacio:"></asp:Label>
                                                                    </td>
                                                                    <td>
                                                                        <uc1:uccombobox ID="ucIdentifEspacio" runat="server" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right">
                                                                        <asp:Label ID="Label1" runat="server" Text="Año - Més de la Pauta:"></asp:Label>
                                                                    </td>
                                                                    <td align="left">
                                                                        <table>
                                                                            <tr>
                                                                                <td>
                                                                                    <dx:ASPxSpinEdit ID="seAño" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Height="21px" Number="2012" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="60px" MaxValue="2050" MinValue="2010" NumberType="Integer">
                                                                                        <SpinButtons HorizontalSpacing="0">
                                                                                        </SpinButtons>
                                                                                    </dx:ASPxSpinEdit>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxSpinEdit ID="seMes" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Height="21px" Number="1" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="40px" MaxValue="12" MinValue="1" NumberType="Integer">
                                                                                        <SpinButtons HorizontalSpacing="0">
                                                                                        </SpinButtons>
                                                                                    </dx:ASPxSpinEdit>
                                                                                </td>
                                                                            </tr>

                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                <td width="30%" align="right">
                                                                
                                                                    <asp:Label ID="Label5" runat="server" Text="Origen:"></asp:Label>
                                                                
                                                                </td>
                                                                <td align="left">
                                                                
                                                                    <uc1:uccombobox ID="ucIdentifOrigen1" runat="server" />
                                                                
                                                                </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" colspan="2">
                                                                        <dx:ASPxButton ID="btnBuscarEspacioPeriodo" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" OnClick="btnBuscarEspacioPeriodo_Click" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Text="Buscar por Espacio y Período" Width="250px">
                                                                            <Image Url="~/Images/Icons/16_L_check.gif">
                                                                            </Image>
                                                                        </dx:ASPxButton>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <br />
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <table width="100%" style="border: 1px solid grey;">
                                                                <tr>
                                                                    <td width="30%" align="right">
                                                                        <asp:Label ID="Label2" runat="server" Text="Número de Pauta:"></asp:Label>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxTextBox ID="txNroPauta" runat="server" Width="170px">
                                                                        </dx:ASPxTextBox>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                <td width="30%" align="right">
                                                                
                                                                    <asp:Label ID="Label4" runat="server" Text="Origen:"></asp:Label>
                                                                
                                                                </td>
                                                                <td align="left">
                                                                
                                                                    <uc1:uccombobox ID="ucIdentifOrigen2" runat="server" />
                                                                
                                                                </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" colspan="2">
                                                                        <dx:ASPxButton ID="btnBuscarPauta" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" OnClick="btnBuscarPauta_Click" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Text="Buscar por Nro. Pauta" Width="250px">
                                                                            <Image Url="~/Images/Icons/16_L_check.gif">
                                                                            </Image>
                                                                        </dx:ASPxButton>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <asp:Label ID="lblMsg" runat="server" ForeColor="Red"></asp:Label>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </ContentTemplate>
                                        </asp:UpdatePanel>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>
    </table>
</asp:Content>
