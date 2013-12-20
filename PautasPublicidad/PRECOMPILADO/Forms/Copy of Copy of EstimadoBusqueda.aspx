<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true" CodeBehind="Copy of Copy of EstimadoBusqueda.aspx.cs" Inherits="PautasPublicidad.Web.Forms.BusquedaPautasMensuales2" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="ucComboBox" TagPrefix="uc1" %>
<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">

    <style type="text/css">
        .style1
        {
            height: 24px;
        }
    </style>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<asp:ScriptManager ID="ScriptManager0" runat="server"></asp:ScriptManager>
    <table width="100%">
        <tr>
            <td align="center">
                    <ContentTemplate>
                        <br />
<dx:ASPxRoundPanel ID="pnlMain" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                    CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                    GroupBoxCaptionOffsetY="-19px" HeaderText="Planilla de Pautas Mensual"
                    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
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
                        <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                            <table width="100%">
                                <tr>
                                    <td>
                                            <ContentTemplate>
                                                <table width="100%">
                                                    <tr>
                                                        <td>
                                                            <table width="100%" style="border: 1px solid grey;">
                                                                <tr>
                                                                                                                                                        <td align="right">
                                                                                        <asp:Label ID="Label121" runat="server" Text="Año - Mes:"></asp:Label>
                                                                                    </td>
                                                                                    <td align="left">
                                                                                        <table>
                                                                                            <tr>
                                                                                                <td align="left" class="style1">
                                                                                                    <dx:ASPxDateEdit ID="deAnoMes" runat="server" 
                                                                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                        CssPostfix="Office2010Silver" DisplayFormatString="yyyy-MM" EditFormat="Custom" 
                                                                                                        EditFormatString="yyyy-MM" Spacing="0" 
                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                        OnValueChanged = "deAnoMes_DateChanged" 
                                                                                                         >
                                                                                                        <CalendarProperties>
                                                                                                            
<HeaderStyle Spacing="1px" />
                                                                                                        
</CalendarProperties>
                                                                                                        <ButtonStyle Width="13px">
                                                                                                        </ButtonStyle>
                                                                                                    </dx:ASPxDateEdit>
                                                                                                </td>

                                                                                                <td align="left">
                                                                                                    <asp:ImageButton ID="btnRefresh" runat="server" ImageUrl="~/Images/Icons/16_find.gif"
                                                                                                    OnClick="btnRefresh_Click" ToolTip="Actualizar" 
                                                                                                    Style="width: 16px;" />
                                                                                                </td>

                                                                                            </tr>
                                                                                        </table>
                                                                                    </td>
                                                                </tr>

                                                                <tr>
                                                                    <td align="right" width="30%" class="style1">
                                                                        <asp:Label ID="Label3" runat="server" Text="Estado:"></asp:Label>

                                                                    </td>
                                                                    <td align=left class="style1">
                                                                        <dx:ASPxComboBox ID="ucEstado" runat="server" 
                                                                            OnSelectedIndexChanged="ucEstado_SelectedIndexChanged" AutoPostBack="True" 
                                                                            EnableCallbackMode="True">
                                                                        </dx:ASPxComboBox>
                                                                    </td>
                                                                    </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Label ID="lblOrigen" runat="server" Text="Origen:" Visible="False"></asp:Label>
                                                                    </td>
                                                                    <td align=left>
                                                                        <dx:ASPxComboBox ID="ucOrigen" runat="server" 
                                                                            OnSelectedIndexChanged="ucOrigen_SelectedIndexChanged" AutoPostBack="True" 
                                                                            EnableCallbackMode="True" Visible="False">
                                                                        </dx:ASPxComboBox>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        &nbsp;</td>
                                                                    <td align= "justify" colspan="2">
                                                                        <dx:ASPxButton ID="btnBuscar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" OnClick="btnBuscar_Click" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Text="Buscar" Width="171px" HorizontalAlign="Center" Enabled="False">
                                                                            <Image Url="~/Images/Icons/16_L_check.gif">
                                                                            </Image>
                                                                        </dx:ASPxButton>
                                                                    </td>
                                                                    <caption>
                                                                        <br />
                                                                        <tr>
                                                                            <td>
                                                                            </td>               
                                                                             <td>
                                                                            </td>
                                                                        </tr>
                                                                    </caption>
                                                                </tr>
                                                                <caption>
                                                                    <br />
                                                                </caption>
                                                            </table>
                                                            <br />
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <asp:Label ID="lblMsg" runat="server" ForeColor="Red"></asp:Label>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </ContentTemplate>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
                <br />
                <br />
                    </ContentTemplate>
            </td>
        </tr>
    </table>
</asp:Content>
