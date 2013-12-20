<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="EspacioContenidos.aspx.cs" Inherits="PautasPublicidad.Web.Forms.EspacioContenidos" %>

<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
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
        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%">
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
                                        <td align="right" width="20%">
                                            <asp:Literal ID="Literal1" runat="server" Text="Espacio Contenido"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <dx:ASPxTextBox ID="txIdentifEspacio" runat="server" Width="100%">
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
                                            <asp:Literal ID="Literal3" runat="server" Text="Medio Publicitario"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <uc1:ucComboBox ID="ucIdentifMedio" runat="server" />
                                        </td>
                                    </tr>
                                    <tr runat="server" id="trControlesDinamicos">
                                        <td colspan="2">
                                            <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                                <ContentTemplate>
                                                    <table width="100%">
                                                        <tr>
                                                            <td width="20%" align="right">
                                                                <asp:Literal ID="Literal4" runat="server" Text="Tipo de espacio "></asp:Literal>
                                                            </td>
                                                            <td>
                                                                <uc1:ucComboBox ID="ucIdentifTipoEsp" runat="server" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table width="100%">
                                                        <tr runat="server" id="trFrecuencia">
                                                            <td width="20%" align="right">
                                                                <asp:Literal ID="Literal5" runat="server" Text="Frecuencia Semanal"></asp:Literal>
                                                            </td>
                                                            <td>
                                                                <uc1:ucComboBox ID="ucIdentifFrecuencia" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr runat="server" id="trHoraInicio">
                                                            <td width="20%" align="right">
                                                                <asp:Literal ID="Literal8" runat="server" Text="Hora de inicio de la pauta"></asp:Literal>
                                                            </td>
                                                            <td>
                                                                <dx:ASPxTimeEdit ID="teHoraInicio" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                    Width="100%" EditFormat="Custom" EditFormatString="HH:mm">
                                                                </dx:ASPxTimeEdit>
                                                            </td>
                                                        </tr>
                                                        <tr runat="server" id="trHoraFin">
                                                            <td width="20%" align="right">
                                                                <asp:Literal ID="Literal9" runat="server" Text="Hora del fin de la pauta"></asp:Literal>
                                                            </td>
                                                            <td>
                                                                <dx:ASPxTimeEdit ID="teHoraFin" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                    Width="100%" EditFormat="Custom" EditFormatString="HH:mm">
                                                                </dx:ASPxTimeEdit>
                                                            </td>
                                                        </tr>
                                                        <tr runat="server" id="trIntervalo">
                                                            <td width="20%" align="right">
                                                                <asp:Literal ID="Literal6" runat="server" Text="Intervalos de las salidas de las pautas"></asp:Literal>
                                                            </td>
                                                            <td>
                                                                <uc1:ucComboBox ID="ucIdentifIntervalo" runat="server" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </ContentTemplate>
                                            </asp:UpdatePanel>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="20%">
                                            <asp:Literal ID="Literal7" runat="server" Text="Formato Orden de Publicidad"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <dx:ASPxRadioButtonList ID="rbFormatoOP" runat="server" Width="100%" RepeatColumns="2">
                                                <Items>
                                                    <dx:ListEditItem Text="Gráfica" Value="GRAFICA" />
                                                    <dx:ListEditItem Text="PNT Salida" Value="PNT_SALIDA" />
                                                    <dx:ListEditItem Text="PNT Producto" Value="PNT_PRODUCTO" />
                                                    <dx:ListEditItem Text="Calendario Descriptivo" Value="CALENDARIO_DESCRIPTIVO" />
                                                    <dx:ListEditItem Text="Calendario Numérico" Value="CALENDARIO_NUMERICO" />
                                                    <dx:ListEditItem Text="Ninguno" Value="NINGUNO" />
                                                </Items>
                                            </dx:ASPxRadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="20%">
                                            <asp:Literal ID="Literal10" runat="server" Text="Responsable"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <dx:ASPxTextBox ID="txResponsable" runat="server" Width="100%">
                                            </dx:ASPxTextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="20%">
                                            <asp:Literal ID="Literal11" runat="server" Text="Nombre Contacto"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <dx:ASPxTextBox ID="txContacto" runat="server" Width="100%">
                                            </dx:ASPxTextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="20%">
                                            <asp:Literal ID="Literal12" runat="server" Text="Email"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <dx:ASPxTextBox ID="txEmail" runat="server" Width="100%">
                                            </dx:ASPxTextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="20%" class="style1">
                                            <asp:Literal ID="Literal13" runat="server" Text="Dirección "></asp:Literal>
                                        </td>
                                        <td align="left" class="style1">
                                            <dx:ASPxTextBox ID="txDireccion" runat="server" Width="100%">
                                            </dx:ASPxTextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="20%">
                                            <asp:Literal ID="Literal14" runat="server" Text="Teléfono"></asp:Literal>
                                        </td>
                                        <td align="left">
                                            <dx:ASPxTextBox ID="txTelefono" runat="server" Width="100%">
                                            </dx:ASPxTextBox>
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
        <dx:ASPxGridView ID="gv" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" Width="100%">
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
    </div>
</asp:Content>
