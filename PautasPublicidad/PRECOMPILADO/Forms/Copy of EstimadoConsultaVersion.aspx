<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EstimadoConsultaVersion.aspx.cs"
    Inherits="PautasPublicidad.Web.Forms.EstimadoConsultaVersion" %>

<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxSplitter" TagPrefix="dx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Consulta de Versiones de Estimado</title>
    <style type="text/css">
        .style1
        {
            width: 788px;
        }
    </style>
</head>
<body <%--style="width: 100%"--%> >
    <form id="form1" runat="server">
    <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" Width="400px" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
        CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
        GroupBoxCaptionOffsetY="-19px" 
        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Height="774px">
        <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
        <HeaderStyle>
            <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
        </HeaderStyle>
        <PanelCollection>
            <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                <table width="800px" height="600px">
                    <tr>
                        <td>
                            <table width="100%">
                                <tr>
                                    <td class="style1">
                                        <dx:ASPxGridView ID="gvCabecera" runat="server" Width="100%" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                            CssPostfix="Office2010Silver" KeyFieldName="RecId">
                                            <Settings ShowVerticalScrollBar="True" />
                                            <Images SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                <LoadingPanelOnStatusBar Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                                </LoadingPanelOnStatusBar>
                                                <LoadingPanel Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                                </LoadingPanel>
                                            </Images>
                                            <ImagesFilterControl>
                                                <LoadingPanel Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                                </LoadingPanel>
                                            </ImagesFilterControl>
                                            <Styles CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver">
                                                <Header ImageSpacing="5px" SortingImageSpacing="5px">
                                                </Header>
                                                <LoadingPanel ImageSpacing="5px">
                                                </LoadingPanel>
                                            </Styles>
                                            <StylesEditors ButtonEditCellSpacing="0">
                                                <ProgressBar Height="21px">
                                                </ProgressBar>
                                            </StylesEditors>
                                        </dx:ASPxGridView>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style1">
                                        <dx:ASPxButton ID="btnRefresh" runat="server" OnClick="btnRefreshSKU_Click" Text="Actualizar"
                                            Width="150px">
                                            <Image Url="~/Images/Crud/16_L_refresh.gif">
                                            </Image>
                                        </dx:ASPxButton>
                                        <dx:ASPxGridView ID="gvDetalle" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                            CssPostfix="Office2010Silver" Width="100%">
                                            <Settings ShowVerticalScrollBar="True" />
                                            <Images SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                <LoadingPanelOnStatusBar Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                                </LoadingPanelOnStatusBar>
                                                <LoadingPanel Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                                </LoadingPanel>
                                            </Images>
                                            <ImagesFilterControl>
                                                <LoadingPanel Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                                </LoadingPanel>
                                            </ImagesFilterControl>
                                            <Styles CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                CssPostfix="Office2010Silver">
                                                <Header ImageSpacing="5px" SortingImageSpacing="5px">
                                                </Header>
                                                <LoadingPanel ImageSpacing="5px">
                                                </LoadingPanel>
                                            </Styles>
                                            <StylesEditors ButtonEditCellSpacing="0">
                                                <ProgressBar Height="21px">
                                                </ProgressBar>
                                            </StylesEditors>
                                        </dx:ASPxGridView>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            <dx:ASPxGridView ID="gvSKUs" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                CssPostfix="Office2010Silver" Width="100%">
                                <Settings ShowVerticalScrollBar="True" />
                                <Images SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                    <LoadingPanelOnStatusBar Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                    </LoadingPanelOnStatusBar>
                                    <LoadingPanel Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                    </LoadingPanel>
                                </Images>
                                <ImagesFilterControl>
                                    <LoadingPanel Url="~/App_Themes/Office2010Silver/GridView/Loading.gif">
                                    </LoadingPanel>
                                </ImagesFilterControl>
                                <Styles CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver">
                                    <Header ImageSpacing="5px" SortingImageSpacing="5px">
                                    </Header>
                                    <LoadingPanel ImageSpacing="5px">
                                    </LoadingPanel>
                                </Styles>
                                <StylesEditors ButtonEditCellSpacing="0">
                                    <ProgressBar Height="21px">
                                    </ProgressBar>
                                </StylesEditors>
                            </dx:ASPxGridView>
                        </td>
                    </tr>
                </table>
            </dx:PanelContent>
        </PanelCollection>
    </dx:ASPxRoundPanel>
    </form>
</body>
</html>
