<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="Frecuencia.aspx.cs" Inherits="PautasPublicidad.Web.Forms.Frecuencia" %>

<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItems.xml"
        XPath="/MenuItems/*"></asp:XmlDataSource>
    <dx:ASPxGridViewExporter ID="ASPxGridViewExporter1" runat="server">
    </dx:ASPxGridViewExporter>
    <dx:ASPxMenu ID="ASPxMenu1" runat="server" AutoSeparators="RootOnly" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
        CssPostfix="Office2010Silver" DataSourceID="XmlDataSource1" ShowPopOutImages="True"
        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
        Width="100%" onitemclick="ASPxMenu1_ItemClick1">
        <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif">
        </LoadingPanelImage>
        <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
        <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
        <LoadingPanelStyle ImageSpacing="5px">
        </LoadingPanelStyle>
        <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
    </dx:ASPxMenu>
    <div align="center" style="vertical-align: top; height: 95%; overflow: auto;">
        <dx:ASPxRoundPanel ID="pnlControls" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
            GroupBoxCaptionOffsetY="-19px" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
            Width="100%" Visible="False">
            <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
            <HeaderStyle>
                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
            </HeaderStyle>
            <PanelCollection>
                <dx:PanelContent ID="PanelContent1" runat="server" SupportsDisabledAttribute="True">
                    <table width="100%">
                        <tr runat="server" id="trMsg">
                            <td align="center">
                                <asp:Label ID="lblMsg" runat="server" Text="¿Está completamente seguro de que desea eliminar los registros seleccionados? Esta operación no puede deshacerse."
                                    Font-Bold="True" ForeColor="Blue"></asp:Label>
                            </td>
                        </tr>
                        <tr runat="server" id="trAbm">
                            <td>
                                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                    <ContentTemplate>
                                        <table runat="server" id="tblControls" width="100%">
                                            <tr>
                                                <td align="right" width="30%">
                                                    <asp:Literal ID="Literal1" runat="server" Text="Frecuencia"></asp:Literal>
                                                </td>
                                                <td align="left">
                                                    <dx:ASPxTextBox ID="txIdentifFrecuencia" runat="server" Width="100%">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td align="right" width="30%">
                                                    <asp:Literal ID="Literal2" runat="server" Text="Descripción"></asp:Literal>
                                                </td>
                                                <td align="left">
                                                    <dx:ASPxTextBox ID="txName" runat="server" Width="100%">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td align="right" width="30%">
                                                    <asp:Literal ID="Literal3" runat="server" Text="Semana / Mes"></asp:Literal>
                                                </td>
                                                <td align="left">
                                                    <dx:ASPxRadioButtonList ID="rbSemMes" runat="server" Width="100%" 
                                                        onselectedindexchanged="rbSemMes_SelectedIndexChanged1">
                                                        <Items>
                                                            <dx:ListEditItem Text="Semana" Value="SEMANA" />
                                                            <dx:ListEditItem Text="Mes" Value="MES" />
                                                        </Items>
                                                    </dx:ASPxRadioButtonList>
                                                </td>
                                            </tr>
                                            <tr runat="server" id="trDias">
                                                <td colspan="2" align="center">
                                                    <dx:ASPxRoundPanel ID="pnlDias" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                        CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                                                        GroupBoxCaptionOffsetY="-19px" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                        Width="400px" HeaderText="Días de la Semana / Días del Mes">
                                                        <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                                                        <HeaderStyle>
                                                            <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                        </HeaderStyle>
                                                        <PanelCollection>
                                                            <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                                                                <table width="100%">
                                                                    <tr>
                                                                        <td align="right" width="20%">
                                                                            <asp:Literal ID="Literal4" runat="server" Text="Día"></asp:Literal>
                                                                        </td>
                                                                        <td>
                                                                            <dx:ASPxComboBox ID="cbDias" runat="server" Width="100%">
                                                                            </dx:ASPxComboBox>
                                                                        </td>
                                                                    </tr>
                                                                    <tr>
                                                                        <td colspan="2" align="right">
                                                                            <table>
                                                                                <tr>
                                                                                    <td>
                                                                                        <asp:Label ID="lblErrorDia" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                                                                                    </td>
                                                                                    <td>
                                                                                        <dx:ASPxButton ID="btnAddDia" runat="server" Text="Agregar" Width="100px" 
                                                                                            OnClick="btnAddDia_Click">
                                                                                            <Image Url="~/Images/Crud/cmd_add.png">
                                                                                            </Image>
                                                                                        </dx:ASPxButton>
                                                                                    </td>
                                                                                    <td>
                                                                                        <dx:ASPxButton ID="btnDeleteDia" runat="server" Text="Eliminar" Width="100px" 
                                                                                            OnClick="btnDeleteDia_Click">
                                                                                            <Image Url="~/Images/Crud/16_cancel.png">
                                                                                            </Image>
                                                                                        </dx:ASPxButton>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                    <tr>
                                                                        <td colspan="2">
                                                                            <dx:ASPxGridView ID="gvABM" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                CssPostfix="Office2010Silver" Width="100%">
                                                                                <Columns>
                                                                                    <dx:GridViewDataComboBoxColumn Caption="Día de la Semana" FieldName="DiaSemana" ShowInCustomizationForm="True"
                                                                                        VisibleIndex="1">
                                                                                    </dx:GridViewDataComboBoxColumn>
                                                                                    <dx:GridViewDataSpinEditColumn Caption="Nro. de Día" FieldName="Dia" ShowInCustomizationForm="True"
                                                                                        VisibleIndex="2">
                                                                                        <PropertiesSpinEdit DisplayFormatString="g">
                                                                                        </PropertiesSpinEdit>
                                                                                    </dx:GridViewDataSpinEditColumn>
                                                                                </Columns>
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
                                                </td>
                                            </tr>
                                        </table>
                                    </ContentTemplate>
                                </asp:UpdatePanel>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table align="right">
                                    <tr>
                                        <td>
                                            <dx:ASPxButton ID="btnCancel" runat="server" Text="Cancelar" Width="150px" OnClick="btnCancel_Click">
                                                <Image Url="~/Images/Crud/Delete_16.png">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                        <td runat="server" id="tdAdd">
                                            <dx:ASPxButton ID="btnAdd" runat="server" Text="Agregar Registro" Width="150px" OnClick="btnAdd_Click">
                                                <Image Url="~/Images/Crud/cmd_add.png">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                        <td runat="server" id="tdSave">
                                            <dx:ASPxButton ID="btnSave" runat="server" Text="Guardar Registro" Width="150px"
                                                OnClick="btnSave_Click">
                                                <Image Url="~/Images/Crud/16_save.gif">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                        <td runat="server" id="tdDelete">
                                            <dx:ASPxButton ID="btnDelete" runat="server" Text="Eliminar Registros" Width="150px"
                                                OnClick="btnDelete_Click">
                                                <Image Url="~/Images/Crud/16_cancel.png">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="lblError" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </dx:PanelContent>
            </PanelCollection>
        </dx:ASPxRoundPanel>
        <dx:ASPxGridView ID="gv" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" Width="100%">
            <Columns>
                <dx:GridViewDataTextColumn Caption="Frecuencia" FieldName="IdentifFrecuencia" VisibleIndex="0">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="Descripción" FieldName="Name" VisibleIndex="1">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataComboBoxColumn Caption="Semana / Mes" FieldName="SemMes" VisibleIndex="3">
                    <PropertiesComboBox>
                        <Items>
                            <dx:ListEditItem Text="Semana" Value="SEMANA" />
                            <dx:ListEditItem Text="Mes" Value="MES" />
                        </Items>
                    </PropertiesComboBox>
                </dx:GridViewDataComboBoxColumn>
            </Columns>
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
            <Templates>
                <DetailRow>
                    <dx:ASPxGridView ID="detailGrid" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                        CssPostfix="Office2010Silver" Width="100%" OnBeforePerformDataSelect="detailGrid_DataSelect">
                        <Columns>
                            <dx:GridViewDataColumn FieldName="IdentifFrecuencia" Caption="IdentifFrecuencia"
                                VisibleIndex="1" />
                            <dx:GridViewDataColumn FieldName="DiaSemana" VisibleIndex="2" />
                            <dx:GridViewDataColumn FieldName="Dia" VisibleIndex="3" />
                        </Columns>
                        <Settings ShowFooter="False" />
                        <TotalSummary>
                            <dx:ASPxSummaryItem FieldName="IdentifFrecuencia" SummaryType="Count" />
                        </TotalSummary>
                    </dx:ASPxGridView>
                </DetailRow>
            </Templates>
            <SettingsDetail ShowDetailRow="true" />
        </dx:ASPxGridView>
    </div>
</asp:Content>
