<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="Avisos.aspx.cs" Inherits="PautasPublicidad.Web.Forms.Avisos" %>

<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="ucComboBox" TagPrefix="uc1" %>
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
                                <table runat="server" id="tblControls" width="100%">
                                    <tr>
                                        <td>
                                            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                                <ContentTemplate>
                                                    <table width="100%">
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal1" runat="server" Text="Aviso"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxTextBox ID="txIdentifAviso" runat="server" Width="100%" MaxLength="15">
                                                                </dx:ASPxTextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal2" runat="server" Text="Descripción"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxTextBox ID="txName" runat="server" Width="100%">
                                                                </dx:ASPxTextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal5" runat="server" Text="Espacio de Contenido"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <uc1:ucComboBox ID="ucIdentifEspacio" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal4" runat="server" Text="Formato del Aviso"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <uc1:ucComboBox ID="ucIdentifFormAviso" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal6" runat="server" Text="Pieza de Arte"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <uc1:ucComboBox ID="ucIdentifPieza" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr runat="server" id="trDuracion">
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal7" runat="server" Text="Duración"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxSpinEdit ID="spDuracion" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" Height="21px" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                    Width="100%">
                                                                    <SpinButtons HorizontalSpacing="0">
                                                                    </SpinButtons>
                                                                </dx:ASPxSpinEdit>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal3" runat="server" Text="Etiqueta de Producto Externa"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxTextBox ID="txEtiquetaProd" runat="server" Width="100%">
                                                                </dx:ASPxTextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal8" runat="server" Text="Zócalo"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxTextBox ID="txZocalo" runat="server" Width="100%">
                                                                </dx:ASPxTextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal14" runat="server" Text="Nro. Ingesta"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxTextBox ID="txNroIngesta" runat="server" Width="100%">
                                                                </dx:ASPxTextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal9" runat="server" Text="Vigencia Desde"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxDateEdit ID="deVigDesde" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                    Width="100%">
                                                                    <CalendarProperties>
                                                                        <HeaderStyle Spacing="1px" />
                                                                    </CalendarProperties>
                                                                    <ButtonStyle Width="13px">
                                                                    </ButtonStyle>
                                                                </dx:ASPxDateEdit>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" width="20%">
                                                                <asp:Literal ID="Literal10" runat="server" Text="Vigencia Hasta"></asp:Literal>
                                                            </td>
                                                            <td align="left">
                                                                <dx:ASPxDateEdit ID="deVigHasta" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                    Width="100%">
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
                                        <td align="center">
                                            <br />
                                            <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                                <ContentTemplate>
                                                    <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" Width="500px" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                        CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                                                        GroupBoxCaptionOffsetY="-19px" HeaderText="Identificadores de Atención" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                        <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                                                        <HeaderStyle>
                                                            <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                        </HeaderStyle>
                                                        <PanelCollection>
                                                            <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                                                                <table width="100%">                                                                    
                                                                    <tr>
                                                                        <td align="right" width="20%">
                                                                            <asp:Literal ID="Literal12" runat="server" Text="Atención"></asp:Literal>
                                                                        </td>
                                                                        <td align="left">
                                                                            <table width="100%">
                                                                            <tr>
                                                                            <td><dx:ASPxComboBox ID="cbIdentifIdentAte" runat="server" DropDownRows="10" 
                                                                                FilterMinLength="3" IncrementalFilteringMode="Contains" 
                                                                                OnItemRequestedByValue="cbIdentifIdentAte_ItemRequestedByValue" 
                                                                                
                                                                                    OnItemsRequestedByFilterCondition="cbIdentifIdentAte_ItemsRequestedByFilterCondition" 
                                                                                    Width="100%">
                                                                            </dx:ASPxComboBox></td>
                                                                            <td style="width: 20px;" align="center"><asp:ImageButton ID="btnAddIdentifIdentAte" runat="server" ImageUrl="~/Images/Crud/cmd_add.png"
                        ToolTip="Agregar" OnClientClick="window.open('IdentificadoresAtencion.aspx' ,'','height=300', 'width=300');" /></td>
                                                                            </tr>
                                                                            </table>
                                                                            
                                                                            
                                                                        </td>
                                                                    </tr>
                                                                    <tr>
                                                                        <td colspan="2" align="right">
                                                                            <table>
                                                                                <tr>
                                                                                    <td>
                                                                                        <asp:Label ID="lblErrorProducto" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                                                                                    </td>
                                                                                    <td>
                                                                                        <dx:ASPxButton ID="btnAddAtencion" runat="server"
                                                                                            Text="Agregar" Width="100px">
                                                                                            <Image Url="~/Images/Crud/cmd_add.png">
                                                                                            </Image>
                                                                                        </dx:ASPxButton>
                                                                                    </td>
                                                                                    <td>
                                                                                        <dx:ASPxButton ID="btnDeleteAtencion" runat="server" OnClick="btnDeleteAtencion_Click"
                                                                                            Text="Eliminar" Width="100px">
                                                                                            <Image Url="~/Images/Crud/16_cancel.png">
                                                                                            </Image>
                                                                                        </dx:ASPxButton>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                    <tr runat="server" id="trDias">
                                                                        <td colspan="2">
                                                                            <dx:ASPxGridView ID="gvABM" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                CssPostfix="Office2010Silver" Width="100%">
                                                                                <Columns>
                                                                                    <dx:GridViewDataTextColumn Caption="Aviso" FieldName="IdentifAviso" ShowInCustomizationForm="True"
                                                                                        VisibleIndex="1" Visible="False">
                                                                                    </dx:GridViewDataTextColumn>
                                                                                    <dx:GridViewDataTextColumn Caption="Atención" FieldName="IdentifIdentAte" ShowInCustomizationForm="True"
                                                                                        VisibleIndex="2">
                                                                                    </dx:GridViewDataTextColumn>
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
                                                </ContentTemplate>
                                            </asp:UpdatePanel>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table align="right">
                                    <tr>
                                        <td>
                                            <dx:ASPxButton ID="btnCancel" runat="server" Text="Cancelar" Width="150px">
                                                <Image Url="~/Images/Crud/Delete_16.png">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                        <td runat="server" id="tdAdd">
                                            <dx:ASPxButton ID="btnAdd" runat="server" Text="Agregar Registro" Width="150px">
                                                <Image Url="~/Images/Crud/cmd_add.png">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                        <td runat="server" id="tdSave">
                                            <dx:ASPxButton ID="btnSave" runat="server" Text="Guardar Registro" Width="150px">
                                                <Image Url="~/Images/Crud/16_save.gif">
                                                </Image>
                                            </dx:ASPxButton>
                                        </td>
                                        <td runat="server" id="tdDelete">
                                            <dx:ASPxButton ID="btnDelete" runat="server" Text="Eliminar Registros" Width="150px">
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
                <dx:GridViewDataTextColumn Caption="Frecuencia" FieldName="IdentifPieza" VisibleIndex="0">
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
                            <dx:GridViewDataColumn FieldName="IdentifAviso" Caption="IdentifAviso" VisibleIndex="1"
                                Visible="false" />
                            <dx:GridViewDataColumn FieldName="IdentifIdentAte" Caption="Identificador de Atención"
                                VisibleIndex="2" />
                        </Columns>
                        <Settings ShowFooter="False" />
                        <TotalSummary>
                            <dx:ASPxSummaryItem FieldName="IdentifPieza" SummaryType="Count" />
                        </TotalSummary>
                    </dx:ASPxGridView>
                </DetailRow>
            </Templates>
            <SettingsDetail ShowDetailRow="true" />
        </dx:ASPxGridView>
    </div>
</asp:Content>
