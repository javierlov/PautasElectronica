<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="Costos.aspx.cs" Inherits="PautasPublicidad.Web.Forms.Costos" %>

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
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxTabControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxClasses" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItemsCosto.xml"
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
                    <table runat="server" id="tblControls" width="100%">
                        <tr runat="server" id="trMsg">
                            <td align="center">
                                <asp:Label ID="lblMsg" runat="server" Text="¿Está completamente seguro de que desea eliminar los registros seleccionados? Esta operación no puede deshacerse."
                                    Font-Bold="True" ForeColor="Blue"></asp:Label>
                            </td>
                        </tr>
                        <tr runat="server" id="trAbm">
                            <td align="center">
                                <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                    <ContentTemplate>
                                        <dx:ASPxPageControl ID="ASPxPageControl1" runat="server" ActiveTabIndex="0" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                            CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                            TabSpacing="0px" Width="100%">
                                            <TabPages>
                                                <dx:TabPage Text="Costo">
                                                    <ContentCollection>
                                                        <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                            <table runat="server" id="tblCosto" width="100%">
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal1" runat="server" Text="Espacio de Contenido"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <uc1:ucComboBox ID="ucIdentifEspacio" runat="server" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
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
                                                                    <td align="right" width="30%">
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
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal5" runat="server" Text="Tipo de Frecuencia"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxRadioButtonList ID="rbFrecuencia" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" RepeatColumns="2" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%">
                                                                            <Items>
                                                                                <dx:ListEditItem Text="Todo" Value="TODO" />
                                                                                <dx:ListEditItem Text="Detallado" Value="DETALLADO" />
                                                                            </Items>
                                                                        </dx:ASPxRadioButtonList>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="2">
                                                                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                                                            <ContentTemplate>
                                                                                <table width="100%">
                                                                                    <tr runat="server" id="trFrecuencia">
                                                                                        <td align="right" width="30%">
                                                                                            <asp:Literal ID="Literal6" runat="server" Text="Frecuencia"></asp:Literal>
                                                                                        </td>
                                                                                        <td align="left">
                                                                                            <uc1:ucComboBox ID="ucIdentifFrecuencia" runat="server" />
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr runat="server" id="trTipoHorario">
                                                                                        <td align="right" width="30%">
                                                                                            <asp:Literal ID="Literal7" runat="server" Text="Tipo de Horario"></asp:Literal>
                                                                                        </td>
                                                                                        <td align="left">
                                                                                            <dx:ASPxRadioButtonList ID="rbHorario" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                CssPostfix="Office2010Silver" RepeatColumns="2" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                Width="100%">
                                                                                                <Items>
                                                                                                    <dx:ListEditItem Text="Todo" Value="TODO" />
                                                                                                    <dx:ListEditItem Text="Detallado" Value="DETALLADO" />
                                                                                                </Items>
                                                                                            </dx:ASPxRadioButtonList>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </ContentTemplate>
                                                                        </asp:UpdatePanel>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal3" runat="server" Text="Ultima Versión confirmada"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxSpinEdit ID="spVersion" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Height="21px" Number="0" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%" ReadOnly="True">
                                                                            <SpinButtons HorizontalSpacing="0">
                                                                            </SpinButtons>
                                                                        </dx:ASPxSpinEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal8" runat="server" Text="Confirmada por"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxTextBox ID="txConfirmado" runat="server" Width="100%" ReadOnly="True">
                                                                        </dx:ASPxTextBox>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal17" runat="server" Text="Fecha Confirmación"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxDateEdit ID="deFecConfirmado" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
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
                                                        </dx:ContentControl>
                                                    </ContentCollection>
                                                </dx:TabPage>
                                                <dx:TabPage Text="Costo por Frecuencia">
                                                    <ContentCollection>
                                                        <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                            <table width="100%" runat="server" id="tblFrecuencia">
                                                                <tr runat="server" id="trDia">
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal11" runat="server" Text="Día"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxSpinEdit ID="spDia" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Height="21px" MaxValue="31" Number="0" Spacing="0"
                                                                            SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%">
                                                                            <SpinButtons HorizontalSpacing="0">
                                                                            </SpinButtons>
                                                                        </dx:ASPxSpinEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr runat="server" id="trDiaSemana">
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal13" runat="server" Text="Día de la Semana"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxCheckBoxList ID="clDiaSemana" runat="server" 
                                                                            CssFilePath="~/App_Themes/Office2003Silver/{0}/styles.css" 
                                                                            CssPostfix="Office2003Silver" RepeatColumns="2" 
                                                                            SpriteCssFilePath="~/App_Themes/Office2003Silver/{0}/sprite.css" Width="100%">
                                                                            <Items>
                                                                                <dx:ListEditItem Text="Lunes" Value="LUNES" />
                                                                                <dx:ListEditItem Text="Martes" Value="MARTES" />
                                                                                <dx:ListEditItem Text="Miercoles" Value="MIERCOLES" />
                                                                                <dx:ListEditItem Text="Jueves" Value="JUEVES" />
                                                                                <dx:ListEditItem Text="Viernes" Value="VIERNES" />
                                                                                <dx:ListEditItem Text="Sabado" Value="SABADO" />
                                                                                <dx:ListEditItem Text="Domingo" Value="DOMINGO" />
                                                                            </Items>
                                                                        </dx:ASPxCheckBoxList>
                                                                    </td>
                                                                </tr>
                                                                <tr runat="server" id="trHoraDesde">
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal22" runat="server" Text="Hora Desde"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxTimeEdit ID="teHoraDesde" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%" EditFormatString="HH:mm" EditFormat="Custom">
                                                                        </dx:ASPxTimeEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr runat="server" id="trHoraHasta">
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal21" runat="server" Text="Hora Hasta"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxTimeEdit ID="teHoraHasta" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%" EditFormatString="HH:mm" EditFormat="Custom">
                                                                        </dx:ASPxTimeEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr runat="server" id="trCosto">
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal12" runat="server" Text="Costo"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxSpinEdit ID="spCostoFrecuencia" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Height="21px" Number="0" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%">
                                                                            <SpinButtons HorizontalSpacing="0">
                                                                            </SpinButtons>
                                                                        </dx:ASPxSpinEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="2" align="right">
                                                                        <table>
                                                                            <tr>
                                                                                <td>
                                                                                    <asp:Label ID="lblErrorFrecuencia" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                                                                                </td>
                                                                                <td>
                                                                                    <dx:ASPxButton ID="btnAddFrecuencia" runat="server" OnClick="btnAddFrecuencia_Click"
                                                                                        Text="Agregar" Width="100px">
                                                                                        <Image Url="~/Images/Crud/cmd_add.png">
                                                                                        </Image>
                                                                                    </dx:ASPxButton>
                                                                                </td>
                                                                                <td>
                                                                                    <dx:ASPxButton ID="btnDeleteFrecuencia" runat="server" OnClick="btnDeleteFrecuencia_Click"
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
                                                                        <dx:ASPxGridView ID="gvABMFrecuencia" runat="server" AutoGenerateColumns="False"
                                                                            CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                                                                            Width="100%">
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
                                                        </dx:ContentControl>
                                                    </ContentCollection>
                                                </dx:TabPage>
                                                <dx:TabPage Text="Costo por Proveedor">
                                                    <ContentCollection>
                                                        <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                            <table width="100%" runat="server" id="tblProveedor">
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal14" runat="server" Text="Proveedor"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <uc1:ucComboBox ID="ucIdentifProv" runat="server" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal15" runat="server" Text="Categoría de Costo"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxRadioButtonList ID="rbCategoria" runat="server" Width="100%" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                            <Items>
                                                                                <dx:ListEditItem Text="Directo" Value="DIRECTO" />
                                                                                <dx:ListEditItem Text="Indirecto" Value="INDIRECTO" />
                                                                            </Items>
                                                                        </dx:ASPxRadioButtonList>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal2" runat="server" Text="Incluido en Orden de Publicidad"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxCheckBox ID="cbIncluidoOP" runat="server" CheckState="Unchecked" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                        </dx:ASPxCheckBox>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal4" runat="server" Text="Estimado"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxCheckBox ID="cbEstimado" runat="server" CheckState="Unchecked" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                        </dx:ASPxCheckBox>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal23" runat="server" Text="Genera Orden de Compra"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxCheckBox ID="cbGeneraOC" runat="server" CheckState="Unchecked" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" 
                                                                            SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                        </dx:ASPxCheckBox>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal18" runat="server" Text="Tipo de Costo"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxRadioButtonList ID="rbTipoCosto" runat="server" Width="100%" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                            <Items>
                                                                                <dx:ListEditItem Text="Fijo Mensual" Value="FIJO_MENSUAL" />
                                                                                <dx:ListEditItem Text="Segundo Fijo" Value="SEGUNDO_FIJO" />
                                                                                <dx:ListEditItem Text="Salida" Value="SALIDA" />
                                                                                <dx:ListEditItem Text="Unidad Pautada" Value="UNIDAD_PAUTADA" />
                                                                            </Items>
                                                                        </dx:ASPxRadioButtonList>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal19" runat="server" Text="Moneda"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <uc1:ucComboBox ID="ucIdentifMon" runat="server" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal20" runat="server" Text="Grossing Up"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxSpinEdit ID="spGrossingUp" runat="server" Number="1" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%">
                                                                            <SpinButtons HorizontalSpacing="0">
                                                                            </SpinButtons>
                                                                        </dx:ASPxSpinEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr id="trCoeficiente0" runat="server">
                                                                    <td runat="server" align="right" width="30%">
                                                                        <asp:Literal ID="Literal16" runat="server" Text="Costo"></asp:Literal>
                                                                    </td>
                                                                    <td runat="server" align="left">
                                                                        <dx:ASPxSpinEdit ID="spCostoProveedor" runat="server" Number="0" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%">
                                                                            <SpinButtons HorizontalSpacing="0">
                                                                            </SpinButtons>
                                                                        </dx:ASPxSpinEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" colspan="2">
                                                                        <table>
                                                                            <tr>
                                                                                <td>
                                                                                    <asp:Label ID="lblErrorProveedor" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                                                                                </td>
                                                                                <td>
                                                                                    <dx:ASPxButton ID="btnAddProveedor" runat="server" OnClick="btnAddProveedor_Click"
                                                                                        Text="Agregar" Width="100px">
                                                                                        <Image Url="~/Images/Crud/cmd_add.png">
                                                                                        </Image>
                                                                                    </dx:ASPxButton>
                                                                                </td>
                                                                                <td>
                                                                                    <dx:ASPxButton ID="btnDeleteProveedor" runat="server" OnClick="btnDeleteProveedor_Click"
                                                                                        Text="Eliminar" Width="100px">
                                                                                        <Image Url="~/Images/Crud/16_cancel.png">
                                                                                        </Image>
                                                                                    </dx:ASPxButton>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                                <tr id="trDias0" runat="server">
                                                                    <td runat="server" colspan="2">
                                                                        <dx:ASPxGridView ID="gvABMProveedor" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </dx:ContentControl>
                                                    </ContentCollection>
                                                </dx:TabPage>
                                                <dx:TabPage Text="Copiar Costos">
                                                    <ContentCollection>
                                                        <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                            <table runat="server" id="Table1" width="100%">
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal24" runat="server" Text="Espacio de Contenido"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <uc1:ucComboBox ID="UcIdentifEspacio0" runat="server"/>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal25" runat="server" Text="Vigencia Desde"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxDateEdit ID="deVigDesde0" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%" OnDateChanged="deVigDesde0_DateChanged">
                                                                            <CalendarProperties>
                                                                                <HeaderStyle Spacing="1px" />
                                                                            </CalendarProperties>
                                                                            <ButtonStyle Width="13px">
                                                                            </ButtonStyle>
                                                                        </dx:ASPxDateEdit>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td align="right" width="30%">
                                                                        <asp:Literal ID="Literal26" runat="server" Text="Vigencia Hasta"></asp:Literal>
                                                                    </td>
                                                                    <td align="left">
                                                                        <dx:ASPxDateEdit ID="deVigHasta0" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                            Width="100%" OnDateChanged="deVigHasta0_DateChanged">
                                                                            <CalendarProperties>
                                                                                <HeaderStyle Spacing="1px" />
                                                                            </CalendarProperties>
                                                                            <ButtonStyle Width="13px">
                                                                            </ButtonStyle>
                                                                        </dx:ASPxDateEdit>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </dx:ContentControl>
                                                    </ContentCollection>
                                                </dx:TabPage>
                                            </TabPages>
                                            <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif">
                                            </LoadingPanelImage>
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                            <Paddings Padding="2px" PaddingLeft="5px" PaddingRight="5px" />
                                            <Paddings Padding="2px" PaddingLeft="5px" PaddingRight="5px" />
                                            <Paddings Padding="2px" PaddingLeft="5px" PaddingRight="5px" />
                                            <Paddings Padding="2px" PaddingLeft="5px" PaddingRight="5px" />
                                            <ContentStyle>
                                                <Paddings Padding="12px" />
                                                <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px" />
                                                <Paddings Padding="12px" />
                                                <Paddings Padding="12px" />
                                                <Paddings Padding="12px" />
                                                <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px"></Border>
                                            </ContentStyle>
                                        </dx:ASPxPageControl>
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
        <dx:ASPxRoundPanel ID="pnlCommit" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
            GroupBoxCaptionOffsetY="-19px" HeaderText="Confirmar" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
            Width="100%" Visible="False">
            <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
            <HeaderImage Url="~/Images/Crud/ActivateQuote_16.png">
            </HeaderImage>
            <HeaderStyle>
                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
            </HeaderStyle>
            <PanelCollection>
                <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                    <table align="center" width="100%">
                        <tr>
                            <td align="center">
                                <asp:Label ID="lblMsgCommit" runat="server" Font-Bold="True" ForeColor="Blue" Text="¿Confirma el Costo?"></asp:Label>
                            </td>
                        </tr>
                        <tr>

                        <td align="right">
                        <table>
                        <tr>
                        <td>
                                <dx:ASPxButton ID="btnCancel2" runat="server" OnClick="btnCancel2_Click" Text="Cancelar"
                                    Width="150px">
                                    <Image Url="~/Images/Crud/Delete_16.png">
                                    </Image>
                                </dx:ASPxButton>
                            </td>
                            <td>
                                <dx:ASPxButton ID="btnCommit" runat="server" OnClick="btnCommit_Click" Text="Confirmar Costo"
                                    Width="150px">
                                    <Image Url="~/Images/Crud/ActivateQuote_16.png">
                                    </Image>
                                </dx:ASPxButton>
                            </td>
                        </tr>
                        </table>
                        </td>

                            
                        </tr>
                    </table>
                </dx:PanelContent>
            </PanelCollection>
        </dx:ASPxRoundPanel>
        <dx:ASPxGridView runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" AutoGenerateColumns="False" Width="100%" ID="gv">
            <Columns>
                <dx:GridViewDataTextColumn FieldName="IdentifPieza" ShowInCustomizationForm="True"
                    Caption="Frecuencia" VisibleIndex="0">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn FieldName="Name" ShowInCustomizationForm="True" Caption="Descripci&#243;n"
                    VisibleIndex="1">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataComboBoxColumn FieldName="SemMes" ShowInCustomizationForm="True"
                    Caption="Semana / Mes" VisibleIndex="3">
                    <PropertiesComboBox>
                        <Items>
                            <dx:ListEditItem Text="Semana" Value="SEMANA"></dx:ListEditItem>
                            <dx:ListEditItem Text="Mes" Value="MES"></dx:ListEditItem>
                        </Items>
                    </PropertiesComboBox>
                </dx:GridViewDataComboBoxColumn>
            </Columns>
            <SettingsDetail ShowDetailRow="True"></SettingsDetail>
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
            <Styles CssPostfix="Office2010Silver" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css">
                <Header SortingImageSpacing="5px" ImageSpacing="5px">
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
                    <b>Costo por Frecuencia:</b>
                    <dx:ASPxGridView ID="detailGridFrecuencia" runat="server" AutoGenerateColumns="true"
                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                        OnBeforePerformDataSelect="detailGrid_DataSelect" Width="100%">
                        <Columns>
                            <dx:GridViewDataColumn Caption="IdentifPieza" FieldName="IdentifPieza" Visible="false"
                                VisibleIndex="1" />
                        </Columns>
                        <Settings ShowFooter="False" />
                        <TotalSummary>
                            <dx:ASPxSummaryItem FieldName="IdentifEspacio" SummaryType="Count" />
                        </TotalSummary>
                    </dx:ASPxGridView>
                    <br />
                    <b>Costo por Proveedor:</b>
                    <dx:ASPxGridView ID="detailGridProveedor" runat="server" AutoGenerateColumns="true"
                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                        OnBeforePerformDataSelect="detailGrid_DataSelect" Width="100%">
                        <Columns>
                            <dx:GridViewDataColumn Caption="IdentifPieza" FieldName="IdentifPieza" Visible="false"
                                VisibleIndex="1" />
                        </Columns>
                        <Settings ShowFooter="False" />
                        <TotalSummary>
                            <dx:ASPxSummaryItem FieldName="IdentifProv" SummaryType="Count" />
                        </TotalSummary>
                    </dx:ASPxGridView>
                </DetailRow>
            </Templates>
        </dx:ASPxGridView>
        <dx:ASPxRoundPanel ID="pnlVersiones" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
            GroupBoxCaptionOffsetY="-19px" HeaderText="Consulta de Versiones" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
            Width="100%" Visible="False">
            <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
            <HeaderImage Url="~/Images/Icons/16_find.gif">
            </HeaderImage>
            <HeaderStyle>
                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
            </HeaderStyle>
            <PanelCollection>
                <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                    <table align="center" width="100%">
                        <tr>
                            <td colspan="2" align="center">
                                <dx:ASPxGridView ID="gvVersiones" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                    CssPostfix="Office2010Silver" Width="100%">
                                    <Columns>
                                        <dx:GridViewDataTextColumn Caption="IdentifEspacio" FieldName="IdentifEspacio" ShowInCustomizationForm="True"
                                            VisibleIndex="1">
                                        </dx:GridViewDataTextColumn>
                                    </Columns>
                                    <SettingsDetail ShowDetailRow="True" />
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
                    <b>Costo por Frecuencia:</b>
                    <dx:ASPxGridView ID="detailGridFrecuenciaVer" runat="server" AutoGenerateColumns="true"
                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                        OnBeforePerformDataSelect="detailGrid_DataSelect" Width="100%">
                        <Columns>
                            <dx:GridViewDataColumn Caption="IdentifPieza" FieldName="IdentifPieza" Visible="false"
                                VisibleIndex="1" />
                        </Columns>
                        <Settings ShowFooter="False" />
                        <TotalSummary>
                            <dx:ASPxSummaryItem FieldName="IdentifEspacio" SummaryType="Count" />
                        </TotalSummary>
                    </dx:ASPxGridView>
                    <br />
                    <b>Costo por Proveedor:</b>
                    <dx:ASPxGridView ID="detailGridProveedorVer" runat="server" AutoGenerateColumns="true"
                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                        OnBeforePerformDataSelect="detailGrid_DataSelect" Width="100%">
                        <Columns>
                            <dx:GridViewDataColumn Caption="IdentifPieza" FieldName="IdentifPieza" Visible="false"
                                VisibleIndex="1" />
                        </Columns>
                        <Settings ShowFooter="False" />
                        <TotalSummary>
                            <dx:ASPxSummaryItem FieldName="IdentifProv" SummaryType="Count" />
                        </TotalSummary>
                    </dx:ASPxGridView>
                </DetailRow>
                                    </Templates>
                                </dx:ASPxGridView>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;
                            </td>
                            <td align="right">
                                <dx:ASPxButton ID="btnCancel3" runat="server" OnClick="btnCancel3_Click" Text="Cerrar"
                                    Width="150px">
                                    <Image Url="~/Images/Crud/Delete_16.png">
                                    </Image>
                                </dx:ASPxButton>
                            </td>
                        </tr>
                    </table>
                </dx:PanelContent>
            </PanelCollection>
        </dx:ASPxRoundPanel>
    </div>
</asp:Content>
