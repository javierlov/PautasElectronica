<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true" CodeBehind="PlanillaPautasMensuales.aspx.cs" Inherits="PautasPublicidad.Web.Forms.PlanillaPautasMensual2" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx" %>
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="ucComboBox" TagPrefix="uc1" %>
<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>

<%@ Register assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxCallbackPanel" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server"></asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

    <asp:ScriptManager ID="ScriptManager1" runat="server" />

            <dx:ASPxGridViewExporter ID="ASPxGridViewExporter1" runat="server" 
                GridViewID="gv">
            </dx:ASPxGridViewExporter>

    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItemsPlanillas.xml" XPath="/MenuItems/*">
    </asp:XmlDataSource>



                <dx:ASPxMenu runat="server" AutoPostBack="True" 
    AutoSeparators="RootOnly" EnableCallBacks="True" EnableClientSideAPI="True" 
    ShowPopOutImages="True" 
    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
    DataSourceID="XmlDataSource1" CssPostfix="Office2010Silver" 
    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" Width="100%" 
    ID="mnuPrincipal" OnItemClick="mnuPrincipal_ItemClick"></dx:ASPxMenu>


    <dx:ASPxGridView ID="gv" runat="server" Width="100%">
        <SettingsBehavior AllowSort="False" />
        <SettingsPager AlwaysShowPager="True">
        </SettingsPager>
        <Settings ShowFooter="True" ShowVerticalScrollBar="True" />
    </dx:ASPxGridView>



        <table width="100%">


                                                    <tr runat="server" id="trButtons">
                                                <td align="right">
                                                    <table align="right">
                                                        <tr>
                                                            <td runat="server" id ="tdValidar" align="left">
                                                                &nbsp;</td>
                                                            <td runat="server" id ="tdCancel">
                                                            <dx:ASPxButton ID="btnVolver" runat="server" OnClick="btnVolver_Click" Text="Volver"
                                                            Width="150px" ToolTip="Volver">
                                                            </dx:ASPxButton>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
    </table>
</asp:Content>
