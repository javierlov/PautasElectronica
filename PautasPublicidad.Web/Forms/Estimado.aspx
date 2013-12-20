<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true" CodeBehind="Estimado.aspx.cs" Inherits="PautasPublicidad.Web.Forms.Estimado" %>
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
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="ucComboBox" TagPrefix="uc1" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxTabControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxClasses" TagPrefix="dx" %>
    <asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItemsOrdenado.xml"
        XPath="/MenuItems/*"></asp:XmlDataSource>
    <asp:XmlDataSource ID="XmlDataSource2" runat="server" DataFile="~/App_Data/MenuItemsOrdenadoDetalle.xml"
        XPath="/MenuItems/*"></asp:XmlDataSource>
    <dx:ASPxGridViewExporter ID="ASPxGridViewExporter1" runat="server">
    </dx:ASPxGridViewExporter>
    <dx:ASPxMenu ID="ASPxMenu1" runat="server" DataSourceID="XmlDataSource1" Width="100%"
        AutoSeparators="RootOnly" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
        CssPostfix="Office2010Silver" ShowPopOutImages="True" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
        <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif">
        </LoadingPanelImage>
        <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
        <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
        <LoadingPanelStyle ImageSpacing="5px">
        </LoadingPanelStyle>
        <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
    </dx:ASPxMenu>
    <div align="center" style="vertical-align: top; height: 95%; overflow: auto;">
        <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
            GroupBoxCaptionOffsetY="-19px" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
            Width="100%" HeaderText="Estimado">
            <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
<ContentPaddings PaddingLeft="9px" PaddingTop="10px" PaddingRight="11px" 
                PaddingBottom="10px"></ContentPaddings>

            <HeaderStyle>
                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
<Paddings PaddingLeft="9px" PaddingTop="3px" PaddingRight="11px" PaddingBottom="6px"></Paddings>
            </HeaderStyle>
            <PanelCollection>
                <dx:PanelContent ID="PanelContent1" runat="server" SupportsDisabledAttribute="True">
                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                        <ContentTemplate>
                            <table width="100%">
                                <tr runat="server" id="trFind">
                                    <td align="center">
                                        <table width="600px">
                                            <tr>
                                                <td align="right" width="20%">
                                                    <asp:Literal ID="Literal1" runat="server" Text="Espacio Contenido"></asp:Literal>
                                                </td>
                                                <td>
                                                    <uc1:ucComboBox ID="ucIdentifEspacio" runat="server" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td align="right">
                                                    <asp:Literal ID="Literal2" runat="server" Text="Medio"></asp:Literal>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txMedio" runat="server" Width="100%" ReadOnly="True">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td align="right">
                                                    <asp:Literal ID="Literal3" runat="server" Text="Año - Mes"></asp:Literal>
                                                </td>
                                                <td align="left">
                                                    <table>
                                                        <tr>
                                                            <td>
                                                                <dx:ASPxDateEdit ID="deAnoMes" runat="server" EditFormat="Custom" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" DisplayFormatString="yyyy-MM" EditFormatString="yyyy-MM"
                                                                    Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                    <CalendarProperties>
                                                                        <HeaderStyle Spacing="1px" />
                                                                    </CalendarProperties>
                                                                    <ButtonStyle Width="13px">
                                                                    </ButtonStyle>
                                                                </dx:ASPxDateEdit>
                                                            </td>
                                                            <td>
                                                                <asp:ImageButton ID="btnRefresh" runat="server" ImageUrl="~/Images/Icons/16_find.gif"
                                                                    OnClick="btnRefresh_Click" ToolTip="Actualizar" Style="width: 16px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">
                                                                <asp:Label ID="lblValidaAñoMes" runat="server" ForeColor="Red"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr runat="server" id="trPauta">
                                    <td align="center">
                                        <table width="100%">
                                            <tr>
                                                <td align="right">
                                                    <asp:Literal ID="litCambiarPauta" runat="server" Text="Seleccionar otra Pauta"></asp:Literal>
                                                    <asp:ImageButton ID="btnBack" runat="server" ImageUrl="~/Images/Crud/16_L_refresh.gif"
                                                        Style="width: 16px" ToolTip="Actualizar" OnClick="btnBack_Click" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="lblErrorLineas" runat="server" ForeColor="Red"></asp:Label>
                                                    <dx:ASPxPageControl ID="ASPxPageControl1" runat="server" ActiveTabIndex="0" Width="100%"
                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                        TabSpacing="0px">
                                                        <TabPages>
                                                            <dx:TabPage Text="Pauta">
                                                                <ContentCollection>
                                                                    <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                                        <table width="100%">
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal4" runat="server" Text="Nro. Pauta"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxSpinEdit ID="spPautaID" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Height="21px" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="100%">
                                                                                        <SpinButtons HorizontalSpacing="0">
                                                                                        </SpinButtons>
                                                                                    </dx:ASPxSpinEdit>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal5" runat="server" Text="Frecuencia"></asp:Literal>
                                                                                </td>
                                                                                <td>
                                                                                    <uc1:ucComboBox ID="ucIdentifFrecuencia" runat="server" />
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal6" runat="server" Text="Hora Inicio"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxTimeEdit ID="teHoraInicio" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="100%" DisplayFormatString="HH:mm" EditFormat="Custom" EditFormatString="HH:mm">
                                                                                    </dx:ASPxTimeEdit>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal7" runat="server" Text="Hora Fin"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxTimeEdit ID="teHoraFin" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="100%" DisplayFormatString="HH:mm" EditFormat="Custom" EditFormatString="HH:mm">
                                                                                    </dx:ASPxTimeEdit>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal8" runat="server" Text="Intervalo"></asp:Literal>
                                                                                </td>
                                                                                <td>
                                                                                    <uc1:ucComboBox ID="ucIdentifIntervalo" runat="server" />
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="2">
                                                                                    <hr />
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal14" runat="server" Text="Versión Costo"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxSpinEdit ID="spVersionCosto" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Height="21px" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="100%">
                                                                                        <SpinButtons HorizontalSpacing="0">
                                                                                        </SpinButtons>
                                                                                    </dx:ASPxSpinEdit>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal9" runat="server" Text="Usuario Cálculo Costo"></asp:Literal>
                                                                                </td>
                                                                                <td>
                                                                                    <dx:ASPxTextBox ID="txUsuCosto" runat="server" ReadOnly="True" Width="100%">
                                                                                    </dx:ASPxTextBox>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal10" runat="server" Text="Fecha Cálculo Costo"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxDateEdit ID="deFecCosto" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        ReadOnly="True" Width="100%">
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
                                                                                    <asp:Literal ID="Literal11" runat="server" Text="Usuario Cierre"></asp:Literal>
                                                                                </td>
                                                                                <td>
                                                                                    <dx:ASPxTextBox ID="txUsuCierre" runat="server" ReadOnly="True" Width="100%">
                                                                                    </dx:ASPxTextBox>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td align="right" width="30%">
                                                                                    <asp:Literal ID="Literal12" runat="server" Text="Fecha Cierre"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxDateEdit ID="deFecCierre" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        ReadOnly="True" Width="100%">
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
                                                                                    <asp:Literal ID="Literal13" runat="server" Text="Cantidad de Salidas"></asp:Literal>
                                                                                </td>
                                                                                <td align="left">
                                                                                    <dx:ASPxSpinEdit ID="spCantSalidas" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                        CssPostfix="Office2010Silver" Height="21px" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                        Width="100%">
                                                                                        <SpinButtons HorizontalSpacing="0">
                                                                                        </SpinButtons>
                                                                                    </dx:ASPxSpinEdit>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="2">
                                                                                    &nbsp;
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="2">
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </dx:ContentControl>
                                                                </ContentCollection>
                                                            </dx:TabPage>
                                                            <dx:TabPage Text="Detalle">
                                                                <ContentCollection>
                                                                    <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                                        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                                                            <ContentTemplate>
                                                                                <table width="100%">
                                                                                    <tr>
                                                                                        <td>
                                                                                            <dx:ASPxMenu ID="mnuDetalle" runat="server" AutoSeparators="RootOnly" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                CssPostfix="Office2010Silver" DataSourceID="XmlDataSource2" ShowPopOutImages="True"
                                                                                                SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%"
                                                                                                OnItemClick="mnuDetalle_ItemClick">
                                                                                                <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif">
                                                                                                </LoadingPanelImage>
                                                                                                <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
                                                                                                <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
                                                                                                <LoadingPanelStyle ImageSpacing="5px">
                                                                                                </LoadingPanelStyle>
                                                                                                <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
                                                                                            </dx:ASPxMenu>
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr runat="server" id="trAccion">
                                                                                        <td align="center">
                                                                                            <dx:ASPxPageControl ID="ASPxPageControl2" runat="server" ActiveTabIndex="0" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                TabSpacing="0px" Width="100%">
                                                                                                <TabPages>
                                                                                                    <dx:TabPage Text="Insertar Líneas">
                                                                                                        <ContentCollection>
                                                                                                            <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                                                                                <table width="100%">
                                                                                                                    <tr>
                                                                                                                        <td width="30%" align="right">
                                                                                                                            <asp:Literal ID="Literal15" runat="server" Text="Días"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            <dx:ASPxCheckBoxList ID="ceDiasInsertar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                Width="100%" RepeatColumns="3">
                                                                                                                            </dx:ASPxCheckBoxList>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal17" runat="server" Text="Hora de Inicio"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            <dx:ASPxTimeEdit ID="teHoraInicioInsertar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                Width="100%" DisplayFormatString="HH:mm" EditFormat="Custom" EditFormatString="HH:mm">
                                                                                                                            </dx:ASPxTimeEdit>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal18" runat="server" Text="Hora de Fin"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            <dx:ASPxTimeEdit ID="teHoraFinInsertar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                Width="100%" DisplayFormatString="HH:mm" EditFormat="Custom" EditFormatString="HH:mm">
                                                                                                                            </dx:ASPxTimeEdit>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal19" runat="server" Text="Número de Salidas"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            <dx:ASPxSpinEdit ID="spSalidasInsertar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                CssPostfix="Office2010Silver" Height="21px" Number="0" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                Width="100%">
                                                                                                                                <SpinButtons HorizontalSpacing="0">
                                                                                                                                </SpinButtons>
                                                                                                                            </dx:ASPxSpinEdit>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal20" runat="server" Text="Aviso"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            <uc1:ucComboBox ID="ucIdentifAviso" runat="server" />
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right" style="font-style: italic;">
                                                                                                                            <asp:Literal ID="Literal21" runat="server" Text="Duración"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            <dx:ASPxSpinEdit ID="spDuracionInsertar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                CssPostfix="Office2010Silver" Height="21px" Number="0" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                Width="100%" ReadOnly="True">
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
                                                                                                                                        <dx:ASPxButton ID="btnCancelInsertar" runat="server" OnClick="btnCancel_Click" Text="Cancelar"
                                                                                                                                            Width="150px">
                                                                                                                                            <Image Url="~/Images/Crud/Delete_16.png">
                                                                                                                                            </Image>
                                                                                                                                        </dx:ASPxButton>
                                                                                                                                    </td>
                                                                                                                                    <td>
                                                                                                                                        <dx:ASPxButton ID="btnInsertarLineas" runat="server" OnClick="btnGenerarLineas_Click"
                                                                                                                                            Text="Insertar Líneas" Width="150px">
                                                                                                                                            <Image Url="~/Images/Icons/18_addView.gif">
                                                                                                                                            </Image>
                                                                                                                                        </dx:ASPxButton>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                </table>
                                                                                                            </dx:ContentControl>
                                                                                                        </ContentCollection>
                                                                                                    </dx:TabPage>
                                                                                                    <dx:TabPage Text="Copiar Períodos">
                                                                                                        <ContentCollection>
                                                                                                            <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                                                                                <table width="100%">
                                                                                                                    <tr>
                                                                                                                        <td width="20%" align="right" style="font-weight: bold;">
                                                                                                                            <asp:Literal ID="Literal27" runat="server" Text="Origen:"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            &nbsp;
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal28" runat="server" Text="Fecha Desde"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <dx:ASPxDateEdit ID="deFechaDesdeOrigenCopiar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal29" runat="server" Text="Fecha Hasta"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <dx:ASPxDateEdit ID="deFechaHastaOrigenCopiar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                        <td align="right" style="font-weight: bold;">
                                                                                                                            <asp:Literal ID="Literal30" runat="server" Text="Destino:"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            &nbsp;
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal31" runat="server" Text="Fecha Desde:"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <dx:ASPxDateEdit ID="deFechaDesdeDestinoCopiar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                    <tr runat="server" id="trFechaHastaReemplazo">
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal32" runat="server" Text="Fecha Hasta:"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <dx:ASPxDateEdit ID="deFechaHastaDestinoCopiar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                        <td align="right" colspan="2">
                                                                                                                            <table>
                                                                                                                                <tr>
                                                                                                                                    <td>
                                                                                                                                        <dx:ASPxButton ID="btnCancelCopiar" runat="server" OnClick="btnCancel_Click" Text="Cancelar"
                                                                                                                                            Width="150px">
                                                                                                                                            <Image Url="~/Images/Crud/Delete_16.png">
                                                                                                                                            </Image>
                                                                                                                                        </dx:ASPxButton>
                                                                                                                                    </td>
                                                                                                                                    <td>
                                                                                                                                        <dx:ASPxButton ID="btnCopiarPeriodos" runat="server" OnClick="btnCopiarPeriodos_Click"
                                                                                                                                            Text="Copiar Períodos" Width="150px">
                                                                                                                                            <Image Url="~/Images/Icons/18_editForm.gif">
                                                                                                                                            </Image>
                                                                                                                                        </dx:ASPxButton>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                </table>
                                                                                                            </dx:ContentControl>
                                                                                                        </ContentCollection>
                                                                                                    </dx:TabPage>
                                                                                                    <dx:TabPage Text="Reemplazar Avisos">
                                                                                                        <ContentCollection>
                                                                                                            <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                                                                                <table width="100%">
                                                                                                                    
              <tr>
                                                                                                                        <td width="20%" align="right" style="font-weight: bold; color">
                                                                                                                            <asp:Literal ID="Literal16" runat="server" Text="Desde:"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>

                                                                                                                        <table>
                                                                                                                        <tr>
                                                                                                                        
                                                                                                                        <td><dx:ASPxRadioButton ID="opEditPeriodo" runat="server" Text="Período Ingresado" 
                                                                                                                                Checked="True" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                CssPostfix="Office2010Silver" GroupName="editDesde" 
                                                                                                                                SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                                                                            </dx:ASPxRadioButton></td>
                                                                                                                        <td><dx:ASPxRadioButton ID="opEditSeleccionados" runat="server" 
                                                                                                                                Text="Todos los Seleccionados" 
                                                                                                                                CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                CssPostfix="Office2010Silver" GroupName="editDesde" 
                                                                                                                                SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"></dx:ASPxRadioButton></td>
                                                                                                                        <td><dx:ASPxRadioButton ID="opEditTodas" runat="server" Text="Todas las Líneas" 
                                                                                                                                CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                CssPostfix="Office2010Silver" GroupName="editDesde" 
                                                                                                                                SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                                                                            </dx:ASPxRadioButton></td>
                                                                                                                        </tr>
                                                                                                                        
                                                                                                                        </table>

                                                                                                                            
                                                                                                                            
                                                                                                                            
                                                                                                                            
                                                                                                                            
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal22" runat="server" Text="Fecha Desde"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <dx:ASPxDateEdit ID="deFechaDesdeReemplazar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal23" runat="server" Text="Fecha Hasta"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <dx:ASPxDateEdit ID="deFechaHastaReemplazar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                        <td align="right" style="font-weight: bold;">
                                                                                                                            <asp:Literal ID="Literal24" runat="server" Text="Avisos:"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td>
                                                                                                                            &nbsp;
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal33" runat="server" Text="Aviso a Reemplazar"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <uc1:ucComboBox ID="ucIdentifAvisoOrigenReemplazar" runat="server" />
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right">
                                                                                                                            <asp:Literal ID="Literal34" runat="server" Text="Nuevo Aviso"></asp:Literal>
                                                                                                                        </td>
                                                                                                                        <td align="left">
                                                                                                                            <uc1:ucComboBox ID="ucIdentifAvisoDestinoReemplazar" runat="server" />
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td align="right" colspan="2">
                                                                                                                            <table>
                                                                                                                                <tr>
                                                                                                                                    <td>
                                                                                                                                        <dx:ASPxButton ID="btnCancelReemplazar" runat="server" OnClick="btnCancel_Click"
                                                                                                                                            Text="Cancelar" Width="150px">
                                                                                                                                            <Image Url="~/Images/Crud/Delete_16.png">
                                                                                                                                            </Image>
                                                                                                                                        </dx:ASPxButton>
                                                                                                                                    </td>
                                                                                                                                    <td>
                                                                                                                                        <dx:ASPxButton ID="btnReemplazarAvisos" runat="server" OnClick="btnReemplazarAvisos_Click"
                                                                                                                                            Text="Reemplazar Avisos" Width="150px">
                                                                                                                                            <Image Url="~/Images/Icons/16_runworkflow.gif">
                                                                                                                                            </Image>
                                                                                                                                        </dx:ASPxButton>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
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
                                                                                                <ContentStyle>
                                                                                                    <Paddings Padding="12px" />
                                                                                                    <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px" />
                                                                                                    <Paddings Padding="12px" />
                                                                                                    <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px"></Border>
                                                                                                </ContentStyle>
                                                                                            </dx:ASPxPageControl>
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr runat="server" id="trEditLine">
                                                                                        <td align="center">
                                                                                            <dx:ASPxRoundPanel ID="pnlEditLine" runat="server" Width="100%" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                                                                                                GroupBoxCaptionOffsetY="-19px" HeaderText="Modificar Línea" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                                                <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                                                                                                <HeaderStyle>
                                                                                                    <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                                                                </HeaderStyle>
                                                                                                <PanelCollection>
                                                                                                    <dx:PanelContent ID="PanelContent2" runat="server" SupportsDisabledAttribute="True">
                                                                                                        <table width="100%">
                                                                                                            <tr>
                                                                                                                <td align="right" style="font-style: italic;">
                                                                                                                    <asp:Literal ID="Literal36" runat="server" Text="Día"></asp:Literal>
                                                                                                                </td>
                                                                                                                <td>
                                                                                                                    <dx:ASPxDateEdit ID="deFechaEdit" runat="server" 
                                                                                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                        CssPostfix="Office2010Silver" Spacing="0" 
                                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                                        Width="100%" ReadOnly="True">
                                                                                                                        <CalendarProperties>
                                                                                                                            <HeaderStyle Spacing="1px" />
                                                                                                                        </CalendarProperties>
                                                                                                                        <ButtonStyle Width="13px">
                                                                                                                        </ButtonStyle>
                                                                                                                    </dx:ASPxDateEdit>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="right">
                                                                                                                    <asp:Literal ID="Literal35" runat="server" Text="Hora"></asp:Literal>
                                                                                                                </td>
                                                                                                                <td>
                                                                                                                    <dx:ASPxTimeEdit ID="teHoraEdit" runat="server" 
                                                                                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                        CssPostfix="Office2010Silver" DisplayFormatString="HH:mm" EditFormat="Custom" 
                                                                                                                        EditFormatString="HH:mm" Spacing="0" 
                                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%">
                                                                                                                    </dx:ASPxTimeEdit>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="right">
                                                                                                                    <asp:Literal ID="Literal25" runat="server" Text="Aviso"></asp:Literal>
                                                                                                                </td>
                                                                                                                <td>
                                                                                                                    <uc1:ucComboBox ID="ucIdentifAvisoEdit" runat="server" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="right" style="font-style: italic;">
                                                                                                                    <asp:Literal ID="Literal26" runat="server" Text="Duración"></asp:Literal>
                                                                                                                </td>
                                                                                                                <td>
                                                                                                                    <dx:ASPxSpinEdit ID="spDuracionEdit" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                        CssPostfix="Office2010Silver" Height="21px" Number="0" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                        Width="100%" ReadOnly="True">
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
                                                                                                                                <dx:ASPxButton ID="btnCancelEdit" runat="server" OnClick="btnCancelEdit_Click" Text="Cancelar"
                                                                                                                                    Width="150px">
                                                                                                                                    <Image Url="~/Images/Crud/Delete_16.png">
                                                                                                                                    </Image>
                                                                                                                                </dx:ASPxButton>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <dx:ASPxButton ID="btnUpdateEdit" runat="server" OnClick="btnUpdateEdit_Click"
                                                                                                                                    Text="Aceptar" Width="150px">
                                                                                                                                    <Image Url="~/Images/Crud/16_save.gif">
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
                                                                                            &nbsp;
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr runat="server" id="trQuerySKU">
                                                                                    <td>
                                                                                    
                                                                                    <dx:ASPxRoundPanel ID="ASPxRoundPanel2" runat="server" Width="100%" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                                                                                                GroupBoxCaptionOffsetY="-19px" HeaderText="Consulta por SKU" 
                                                                                            SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                                                <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                                                                                                <HeaderStyle>
                                                                                                    <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                                                                </HeaderStyle>
                                                                                                <PanelCollection>
                                                                                                    <dx:PanelContent ID="PanelContent3" runat="server" SupportsDisabledAttribute="True">
                                                                                                        <table width="100%">
                                                                                                       
                                                                                                            <tr>
                                                                                                                <td>
                                                                                                                    <table width="100%">
                                                                                                                    <tr>
                                                                                                                    <td>
                                                                                                                        <dx:ASPxGridView ID="gvSKU" runat="server" AutoGenerateColumns="False" 
                                                                                                                            CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                            CssPostfix="Office2010Silver" OnRowUpdating="gv_RowUpdating" 
                                                                                                                            OnStartRowEditing="gv_StartRowEditing" Width="100%">
                                                                                                                            
                                                                                                                            <Columns>
                                                                                                                                
                                                                                                                                <dx:GridViewDataTextColumn Caption="Producto" FieldName="IdentifSKU" 
                                                                                                                                    ShowInCustomizationForm="True" VisibleIndex="0" Name="IdentifSKU">
                                                                                                                                </dx:GridViewDataTextColumn>
                                                                                                                                <dx:GridViewDataTextColumn Caption="Descripción" FieldName="Name" 
                                                                                                                                    ShowInCustomizationForm="True" VisibleIndex="1" Name="Name">
                                                                                                                                </dx:GridViewDataTextColumn>
                                                                                                                                <dx:GridViewDataTextColumn FieldName="CantSalidas" 
                                                                                                                                    ShowInCustomizationForm="True" VisibleIndex="2" Caption="Salidas" 
                                                                                                                                    Name="CantSalidas">
                                                                                                                                </dx:GridViewDataTextColumn>
                                                                                                                            </Columns>
                                                                                                                            <SettingsPager AlwaysShowPager="True">
                                                                                                                            </SettingsPager>
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
                                                                                                                <td align="right" colspan="2">
                                                                                                                    <table width="100%">
                                                                                                                        <tr>
                                                                                                                        <td align="left" width="100%">
                                                                                                                        <asp:Label ID="lblSKUTotalSalidas" runat="server" Text="" Font-Bold="True"></asp:Label>
                                                                                                                        &nbsp;</td>

                                                                                                                            <td align="right">
                                                                                                                                <dx:ASPxButton ID="btnCancelSKU" runat="server" Text="Cerrar"
                                                                                                                                    Width="150px" OnClick="btnCancelSKU_Click">
                                                                                                                                    <Image Url="~/Images/Crud/Delete_16.png">
                                                                                                                                    </Image>
                                                                                                                                </dx:ASPxButton>
                                                                                                                            </td>

                                                                                                                            <td align="right">
                                                                                                                                <dx:ASPxButton ID="btnRefreshSKU" runat="server" Text="Actualizar"
                                                                                                                                    Width="150px" OnClick="btnRefreshSKU_Click">
                                                                                                                                    <Image Url="~/Images/Crud/16_L_refresh.gif">
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
                                                                                    </td>
                                                                                    
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <td>
                                                                                            <dx:ASPxGridView ID="gv" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                CssPostfix="Office2010Silver" Width="100%" onrowupdating="gv_RowUpdating" 
                                                                                                onstartrowediting="gv_StartRowEditing">
                                                                                                <ClientSideEvents RowDblClick="function(s, e) { s.StartEditRow(e.visibleIndex); }" />
                                                                                                <Columns>
                                                                                                <dx:GridViewCommandColumn ShowInCustomizationForm="True" ShowSelectCheckbox="True"
                                                                                                        VisibleIndex="0" ButtonType="Image" Width="40px">
                                                                                                       
                                                                                                    </dx:GridViewCommandColumn>
                                                                                                    <dx:GridViewCommandColumn ShowInCustomizationForm="True" ShowSelectCheckbox="False"
                                                                                                        VisibleIndex="7" ButtonType="Image" Width="50px">
                                                                                                        <EditButton Visible="True">
                                                                                                            <Image Url="~/Images/Crud/EditProperties_16.png">
                                                                                                            </Image>
                                                                                                        </EditButton>
                                                                                                        <CancelButton>
                                                                                                            <Image Url="~/Images/Crud/Delete_16.png">
                                                                                                            </Image>
                                                                                                        </CancelButton>
                                                                                                        <UpdateButton>
                                                                                                            <Image Url="~/Images/Crud/Save_16.png">
                                                                                                            </Image>
                                                                                                        </UpdateButton>
                                                                                                    </dx:GridViewCommandColumn>
                                                                                                    <dx:GridViewDataTextColumn FieldName="Dia" ShowInCustomizationForm="True" VisibleIndex="0"
                                                                                                        Caption="Día" ReadOnly="True">
                                                                                                    </dx:GridViewDataTextColumn>
                                                                                                    <dx:GridViewDataTextColumn FieldName="DiaSemana" ShowInCustomizationForm="True" VisibleIndex="1"
                                                                                                        Caption="Día Semana" ReadOnly="True">
                                                                                                    </dx:GridViewDataTextColumn>
                                                                                                    <dx:GridViewDataTextColumn FieldName="Hora" ShowInCustomizationForm="True" 
                                                                                                        VisibleIndex="2" ReadOnly="True">
                                                                                                    </dx:GridViewDataTextColumn>
                                                                                                    <dx:GridViewDataTextColumn FieldName="Salida" ShowInCustomizationForm="True" 
                                                                                                        VisibleIndex="3" ReadOnly="True">
                                                                                                    </dx:GridViewDataTextColumn>
                                                                                                    <dx:GridViewDataTextColumn FieldName="Duracion" ShowInCustomizationForm="True" VisibleIndex="5"
                                                                                                        Caption="Duración" ReadOnly="True">
                                                                                                    </dx:GridViewDataTextColumn>
                                                                                                    <dx:GridViewDataComboBoxColumn Caption="Aviso" FieldName="IdentifAviso" 
                                                                                                        VisibleIndex="4">
                                                                                                    </dx:GridViewDataComboBoxColumn>
                                                                                                    <dx:GridViewDataDateColumn Caption="Fecha" VisibleIndex="6" FieldName="Fecha" 
                                                                                                        ReadOnly="True">
                                                                                                    </dx:GridViewDataDateColumn>
                                                                                                </Columns>
                                                                                                <SettingsPager AlwaysShowPager="True">
                                                                                                </SettingsPager>
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
                                                                                    <tr runat="server" id="trLinesButtons">
                                                                                        <td>
                                                                                            <table align="right">
                                                                                                <tr>
                                                                                                    <td runat="server" id="td1">
                                                                                                        <dx:ASPxButton ID="btnSaveLines" runat="server" Text="Guardar Líneas" Width="150px"
                                                                                                            OnClick="btnSaveLines_Click">
                                                                                                            <Image Url="~/Images/Crud/16_save.gif">
                                                                                                            </Image>
                                                                                                        </dx:ASPxButton>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <dx:ASPxButton ID="btnDeleteLines" runat="server" OnClick="btnDeleteLines_Click"
                                                                                                            Text="Eliminar Líneas" Width="150px">
                                                                                                            <Image Url="~/Images/Crud/16_cancel.png">
                                                                                                            </Image>
                                                                                                        </dx:ASPxButton>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </ContentTemplate>
                                                                        </asp:UpdatePanel>
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
                                                        <paddings padding="2px" paddingleft="5px" 
                                                            paddingright="5px" />
                                                        <ContentStyle>
                                                            <Paddings Padding="12px" />
                                                            <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px" />
                                                            <Paddings Padding="12px" />
                                                            <paddings padding="12px" />
                                                            <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px"></Border>
                                                        </ContentStyle>
                                                    </dx:ASPxPageControl>
                                                </td>
                                            </tr>
                                            <tr runat="server" id="trButtons">
                                                <td align="right">
                                                    <table align="right">
                                                        <tr>
                                                            <td runat="server" id="tdAdd">
                                                                <dx:ASPxButton ID="btnAdd" runat="server" Text="Nuevo Ordenado" Width="150px" OnClick="btnAdd_Click">
                                                                    <Image Url="~/Images/Crud/cmd_add.png">
                                                                    </Image>
                                                                </dx:ASPxButton>
                                                            </td>
                                                            <td runat="server" id="tdSave">
                                                                <dx:ASPxButton ID="btnSave" runat="server" Text="Guardar Ordenado" Width="150px"
                                                                    OnClick="btnSave_Click">
                                                                    <Image Url="~/Images/Crud/16_save.gif">
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
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </dx:PanelContent>
            </PanelCollection>
        </dx:ASPxRoundPanel>
    </div>
</asp:Content>
