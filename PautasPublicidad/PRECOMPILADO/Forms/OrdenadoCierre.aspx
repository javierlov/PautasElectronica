<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="OrdenadoCierre.aspx.cs" Inherits="PautasPublicidad.Web.Forms.OrdenadoCierre" %>

<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
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
                    GroupBoxCaptionOffsetY="-19px" HeaderText="Cierre Ordenado" 
                    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                    <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                    <HeaderStyle>
                        <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                    </HeaderStyle>
                    <PanelCollection>
                        <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                            <table width="100%">
                            <tr>
                            <td>
                            <asp:UpdatePanel ID="UpdatePanel1" runat="server" width="100%">
                                <ContentTemplate>
                                    <table width="100%">
                                        <tr>
                                            <td align="right">
                                                <asp:Label ID="Label1" runat="server" Text="Año - Més de la Pauta:"></asp:Label>
                                            </td>
                                            <td align="left" width="60">
                                                <dx:ASPxSpinEdit ID="seAño" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                    CssPostfix="Office2010Silver" Height="21px" Number="2012" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                    Width="60px" onnumberchanged="seAñoMes_NumberChanged" AutoPostBack="True" 
                                                    MaxValue="2050" MinValue="2010" NumberType="Integer">
                                                    <SpinButtons HorizontalSpacing="0">
                                                    </SpinButtons>
                                                </dx:ASPxSpinEdit>
                                            </td>
                                            <td align="right">
                                                <dx:ASPxSpinEdit ID="seMes" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                    CssPostfix="Office2010Silver" Height="21px" Number="1" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                    Width="40px" onnumberchanged="seAñoMes_NumberChanged" AutoPostBack="True" 
                                                    MaxValue="12" MinValue="1" NumberType="Integer">
                                                    <SpinButtons HorizontalSpacing="0">
                                                    </SpinButtons>
                                                </dx:ASPxSpinEdit>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="right">
                                                <asp:Label ID="Label2" runat="server" Text="Fecha del Cierre:" Visible="False"></asp:Label>
                                            </td>
                                            <td colspan="2" align="left">
                                                <dx:ASPxDateEdit ID="deFechaCierre" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                    CssPostfix="Office2010Silver" Date="2012-01-01" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                    Width="110px" ClientVisible="False">
                                                    <CalendarProperties>
                                                        <HeaderStyle Spacing="1px" />
                                                    </CalendarProperties>
                                                    <ButtonStyle Width="13px">
                                                    </ButtonStyle>
                                                </dx:ASPxDateEdit>
                                            </td>
                                        </tr>                                        
                                    </table>
                                </ContentTemplate>
                            </asp:UpdatePanel>
                            </td>
                            </tr>
                            <tr>
                            <td align="right">
                            
                                <dx:ASPxButton ID="btnCierre" runat="server" 
                                    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                    CssPostfix="Office2010Silver" 
                                    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                    Text="Realizar Cierre" Width="150px" OnClick="btn_ShowCierre">
                                    <Image Url="~/Images/Icons/16_L_check.gif">
                                    </Image>
                                </dx:ASPxButton>
                            </td>
                            </tr>
                            <tr>
                            <td align="left">
                            
                                <asp:Label ID="lblMsg" runat="server" ForeColor="Red"></asp:Label>
                            
                            </td>
                            </tr>
                            </table>

                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
                <table runat="server" id="tblCerrar" width="100%">
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td>
                                                <table align="left" id="TblCierre">
                                                    <tr>
                                                        <td>
                                                            <asp:Label ID="Label3" runat="server" Text="Label">¿Esta seguro que quiere Cerrar el periodo indicado?</asp:Label>
                                                        </td>
                                                    </tr>
                                                </table>
                                                <table align="right" >
                                                        <tr>
                                                            <td>
                                                                <dx:ASPxButton ID="Btn_CerrarOrdenado" runat="server" OnClick="btn_CerrarOrdenado"
                                                                    Text="Cerrar" Width="150px">
                                                                    <Image Url="~/Images/Icons/16_succeeded.png">
                                                                    </Image>
                                                                </dx:ASPxButton>
                                                            </td>
                                                            <td>
                                                                <dx:ASPxButton ID="Btn_CancelarCierreOrdenado" runat="server" OnClick="btn_CancelarCierreOrdenado" 
                                                                    Text="Cancelar"
                                                                    Width="150px">
                                                                    <Image Url="~/Images/Crud/Delete_16.png">
                                                                    </Image>
                                                                </dx:ASPxButton>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    
                                                </td>
                                            </tr>
                                        </table>
            </td>
        </tr>
    </table>
</asp:Content>
