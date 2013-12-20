<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true" CodeBehind="Certificado.aspx.cs" Inherits="PautasPublicidad.Web.Forms.Certificado" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="ucComboBox" TagPrefix="uc1" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxTabControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxClasses" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxLoadingPanel" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxCallback" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">

    <%--<style type="text/css">
        .style1
        {
            height: 30px;
        }
    </style>
--%>
    <style type="text/css">
        .style1
        {
            height: 25px;
        }
    </style>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePartialRendering="false" EnablePageMethods="True" />

    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItemsOrdenado.xml" XPath="/MenuItems/*" />

    <asp:TextBox ID="TextBox1" runat="server" BorderStyle="None" ForeColor="White" 
        Width="1px"></asp:TextBox>

    <dx:ASPxCallback ID="ASPxCallback2" runat="server" ClientInstanceName="Callback">
        <ClientSideEvents CallbackComplete="function(s, e) { lp.Hide(); }" />
    </dx:ASPxCallback>


        <dx:ASPxLoadingPanel runat="server" Modal="True" 
        Text="Creando Orden de Publicidad..." ClientInstanceName="lp" 
        HorizontalAlign="Center" VerticalAlign="Middle" ID="ASPxLoadingPanel1"></dx:ASPxLoadingPanel>



    <dx:ASPxPanel ID="ASPxPanel1" runat="server" Width="100%" ClientVisible="true">

         <PanelCollection>

            <dx:PanelContent runat="server" SupportsDisabledAttribute="True">

<%--                        <script type="text/javascript">

                            function showDirectory(s, e) {

                                var valor = "";

                                if (e.item.name == "btnOP") {

                                    var boton = e.item.name.toString();

                                    valor = window.showModalDialog("BrowseDirectory.aspx", 'jain', "dialogHeight: 560px; dialogWidth:360px; edge: Raised; center: Yes; help: Yes; resizable: Yes; status: No;");

                                    var textbox = document.getElementById('ASPxSplitter1_MainContent_TextBox1');

                                    textbox.value = valor;

                                }

                                return valor;

                            }    
                </script>
--%>
                </script>
                <script type="text/javascript">
                    function MostrarLoading(s, e) {

                        if (e.item.name == "btnOP") {

                            Callback.PerformCallback();

                            lp.Show();

                        }

                    }
                </script>

                <dx:ASPxMenu ID="mnuPrincipal" 
                             runat="server" 
                             AutoPostBack="True" 
                             AutoSeparators="RootOnly" 
                             CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                             CssPostfix="Office2010Silver" 
                             DataSourceID="XmlDataSource1" 
                             EnableCallBacks="True" 
                             EnableClientSideAPI="True" 
                             OnItemClick="mnuPrincipal_ItemClick" 
                             ShowPopOutImages="True" 
                             SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                             Width="100%">
                            <ClientSideEvents ItemClick="function(s, e) {MostrarLoading(s,e);}" />
                            <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif" />
                            <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
                            <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
                            <LoadingPanelStyle ImageSpacing="5px" />
                            <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
                </dx:ASPxMenu>

            </dx:PanelContent>

         </PanelCollection>

    </dx:ASPxPanel>
                                                    
    <asp:XmlDataSource ID="XmlDataSource2" runat="server" DataFile="~/App_Data/MenuItemsOrdenadoDetalle.xml" XPath="/MenuItems/*" />

    <asp:HiddenField ID="HiddenFieldId" runat="server" />

    <br />
    <dx:ASPxCallback ID="ASPxCallback1" runat="server" ClientIDMode ="AutoID"
        OnCallback="ASPxCallback1_Callback" ClientInstanceName="ASPxCallback1" 
        ViewStateMode="Enabled">
    </dx:ASPxCallback>
    <asp:TextBox ID="hdnField" runat="server" Visible="False" AutoPostBack="True" 
        ClientIDMode="AutoID" BorderStyle="None" Width="1px"></asp:TextBox>


    <dx:ASPxGridViewExporter ID="ASPxGridViewExporter1" runat="server" GridViewID="gv" OnRenderBrick="ASPxGridViewExporter1_RenderBrick" />

    <div align="center" style="vertical-align: top; height: 95%; overflow: auto;">
        <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
            GroupBoxCaptionOffsetY="-19px" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
            Width="100%" HeaderText="Certificado">
            <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
<ContentPaddings PaddingLeft="9px" PaddingTop="10px" PaddingRight="11px" 
                PaddingBottom="10px"></ContentPaddings>

            <HeaderStyle>
                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
<Paddings PaddingLeft="9px" PaddingTop="3px" PaddingRight="11px" PaddingBottom="6px"></Paddings>
            </HeaderStyle>
            <PanelCollection>
                <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                        <ContentTemplate>
                            <table width="100%">
                                <tr runat="server" id="trFind">
                                    <td align="center">
                                        <br />
                                        <dx:ASPxRoundPanel ID="pnlMain" runat="server" Width="500px"
                                        CssPostfix="Office2010Silver" 
                                        EnableDefaultAppearance="False" 
                                        GroupBoxCaptionOffsetX="6px"
                                        GroupBoxCaptionOffsetY="-19px" 
                                        HeaderText="Selección de Certificado" 
                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                        <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                                        <ContentPaddings PaddingLeft="9px" PaddingTop="10px" PaddingRight="11px" PaddingBottom="10px" />
                                            <HeaderStyle>
                                                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                <Paddings PaddingLeft="9px" PaddingTop="3px" PaddingRight="11px" PaddingBottom="6px" />
                                            </HeaderStyle>
                                                <PanelCollection>
                                                    <dx:PanelContent ID="PanelContent2" runat="server" SupportsDisabledAttribute="True">
                                                    <table width="100%">
                                                        <tr>
                                                            <td>
                                                                <asp:UpdatePanel ID="UpdatePanel3" runat="server">
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
                                                                                        <asp:Label ID="Label121" runat="server" Text="Año - Més de la Pauta:"></asp:Label>
                                                                                    </td>
                                                                                    <td align="left">
                                                                                        <table>
                                                                                            <tr>
                                                                                                <td class="style1">
                                                                                                    </td>
                                                                                                <td align="left" class="style1">
                                                                                                    <dx:ASPxDateEdit ID="deAnoMes" runat="server" 
                                                                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                        CssPostfix="Office2010Silver" DisplayFormatString="yyyy-MM" EditFormat="Custom" 
                                                                                                        EditFormatString="yyyy-MM" Spacing="0" 
                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css">
                                                                                                        <CalendarProperties>
                                                                                                            <HeaderStyle Spacing="1px" />
                                                                                                        </CalendarProperties>
                                                                                                        <ButtonStyle Width="13px">
                                                                                                        </ButtonStyle>
                                                                                                    </dx:ASPxDateEdit>
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
                                                                                <asp:Label ID="Label1" runat="server" Text="Número de Pauta:"></asp:Label>
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
                                        <asp:DropDownList ID="ddlNroPauta" runat="server" Visible="False">
                                        </asp:DropDownList>
                                        <br />
                                        <asp:Literal ID="Literal37" runat="server" Text="Consultar Certificados Grabados"></asp:Literal>
                                        <table width="100%">
                                            <tr>
                                                <td align="left">
                                                    <dx:ASPxButton ID="btnSelOrd" runat="server" Height="39px" OnClick="btn_ConsultarGvHome"
                                                        Text="Consultar Seleccionado" Width="150px" align="left" ClientInstanceName="btnSelOrd">
                                                        <Image Url="~/Images/Icons/16_find.gif">
                                                        </Image>
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right">
                                                    <dx:ASPxButton ID="btnDelOrd" runat="server" Height="39px" OnClick="btn_ShowDelete"
                                                        Text="Eliminar Seleccionado" Width="150px" align="right">
                                                        <Image Url="~/Images/Crud/16_cancel.png">
                                                        </Image>
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>
                                        <table runat="server" id="tblDelete" width="100%">
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td>
                                                <table align="left" id="TblEliminar">
                                                    <tr>
                                                        <td>
                                                            <asp:Label ID="Label2" runat="server" 
                                                                Text="¿Esta seguro de que desea eliminar el Certificado Seleccionado?"></asp:Label>
                                                        </td>
                                                    </tr>
                                                </table>
                                                <table align="right" >
                                                        <tr>
                                                            <td>
                                                                <dx:ASPxButton ID="Btn_EliminarLineaGvHome" runat="server" OnClick="btn_EliminarLineaGvHome"
                                                                    Text="Eliminar" Width="150px">
                                                                    <Image Url="~/Images/Crud/16_cancel.png">
                                                                    </Image>
                                                                </dx:ASPxButton>
                                                            </td>
                                                            <td>
                                                                <dx:ASPxButton ID="btn_CancelarDelete" runat="server" OnClick="btn_CancelDelete" Text="Cancelar"
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
                                        
                                        <asp:Label ID="lblErrorHome" runat="server" ForeColor="Red"></asp:Label>
                                            <script type="text/javascript">
                                                function OnRowDblClick(s, e) {
                                                    document.getElementById('<%=btnSelOrd.ClientID%>').click();
                                                }
                                            </script>

                                        <dx:ASPxGridView ID="gvHome" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                            CssPostfix="Office2010Silver" OnCustomColumnDisplayText="gvHome_CustomColumnDisplayText" ClientSideEvents-RowDblClick="gvHome_RowDblClick"
                                            Width="100%">
                                            <ClientSideEvents RowDblClick="OnRowDblClick" EndCallback="function(s, e) {}"/>
                                            <Columns>
                                                <dx:GridViewDataTextColumn Caption="Pauta" Name="Pauta" VisibleIndex="1" FieldName="PautaId">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="Espacio" Name="Espacio" VisibleIndex="2" FieldName="IdentifEspacio">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="Año" Name="AñoCosto" VisibleIndex="4" FieldName="año">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="Mes" Name="MesCosto" VisibleIndex="5" FieldName="mes">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="Hora Inicio" Name="HoraInicio" VisibleIndex="6"
                                                    FieldName="HoraInicio">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="Hora Fin" Name="HoraFin" VisibleIndex="7" FieldName="HoraFin">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="CertValido" Name="CertValido" VisibleIndex="9"
                                                    FieldName="CertValido">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="FecCertValido" Name="FecCertValido" VisibleIndex="10"
                                                    FieldName="FecCertValido">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="Identif. Origen" Name="IdentifOrigen" VisibleIndex="8"
                                                    FieldName="IdentifOrigen">
                                                </dx:GridViewDataTextColumn>
                                                <dx:GridViewDataTextColumn Caption="AnoMes" FieldName="AnoMes" Visible="False" VisibleIndex="11">
                                                </dx:GridViewDataTextColumn>
                                            </Columns>
                                            <SettingsBehavior AutoExpandAllGroups="True" />
                                            <SettingsPager AlwaysShowPager="True">
                                            </SettingsPager>
                                            <Settings ShowFilterRow="True" />
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
                                        &nbsp;
                                        
                                    </td>
                                </tr>
                                
                                <tr runat="server" id="trPauta">
                                    <td align="center">
                                        <table width="100%">
                                            <tr id="trMnuPrincipal">
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td align="right">
                                                    <asp:Literal ID="litCambiarPauta" runat="server" Text="Seleccionar otra Pauta"></asp:Literal>
                                                    <asp:ImageButton ID="btnBack" runat="server" ImageUrl="~/Images/Crud/16_L_refresh.gif"
                                                        Style="width: 16px" ToolTip="Actualizar" OnClick="btnBack_Click" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table width="100%">
                                                        
                                                        <tr runat="server" id="trQuerySKU">
                                                            <td>
                                                                <dx:ASPxRoundPanel ID="pnlQuerySKU" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                    CssPostfix="Office2010Silver" EnableDefaultAppearance="False" GroupBoxCaptionOffsetX="6px"
                                                                    GroupBoxCaptionOffsetY="-19px" HeaderText="Consulta por SKU" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                    Width="100%">
                                                                    <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="10px" />
                                                                    <HeaderStyle>
                                                                        <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                                    </HeaderStyle>
                                                                    <PanelCollection>
                                                                        <dx:PanelContent ID="PanelContent1" runat="server" SupportsDisabledAttribute="True">
                                                                            <table width="100%">
                                                                                <tr>
                                                                                    <td>
                                                                                        <table width="100%">
                                                                                            <tr>
                                                                                                <td>
                                                                                                    <dx:ASPxGridView ID="gvSKU" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                        CssPostfix="Office2010Silver" OnRowUpdating="gv_RowUpdating" OnStartRowEditing="gv_StartRowEditing"
                                                                                                        Width="100%">
                                                                                                        <Columns>
                                                                                                            <dx:GridViewDataTextColumn Caption="SKU" FieldName="IdentifSKU" Name="IdentifSKU"
                                                                                                                ShowInCustomizationForm="True" VisibleIndex="0">
                                                                                                            </dx:GridViewDataTextColumn>
                                                                                                            <dx:GridViewDataTextColumn Caption="Descripción" FieldName="Name" Name="Name" ShowInCustomizationForm="True"
                                                                                                                VisibleIndex="1">
                                                                                                            </dx:GridViewDataTextColumn>
                                                                                                            <dx:GridViewDataTextColumn Caption="Salidas" FieldName="CantSalidas" Name="CantSalidas"
                                                                                                                ShowInCustomizationForm="True" VisibleIndex="2">
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
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td align="right">
                                                                                        <table width="100%">
                                                                                            <tr>
                                                                                                <td align="left" width="100%">
                                                                                                    <asp:Label ID="lblSKUTotalSalidas" runat="server" Font-Bold="True" Text=""></asp:Label>
                                                                                                    &nbsp;
                                                                                                </td>
                                                                                                <td align="right">
                                                                                                    <dx:ASPxButton ID="btnRefreshSKU" runat="server" OnClick="btnRefreshSKU_Click" Text="Actualizar"
                                                                                                        Width="150px">
                                                                                                        <Image Url="~/Images/Crud/16_L_refresh.gif">
                                                                                                        </Image>
                                                                                                    </dx:ASPxButton>
                                                                                                </td>
                                                                                                <td align="right">
                                                                                                    <dx:ASPxButton ID="btnCancelSKU" runat="server" OnClick="btnCancelSKU_Click" Text="Cerrar"
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
                                                                        </dx:PanelContent>
                                                                    </PanelCollection>
                                                                </dx:ASPxRoundPanel>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <dx:ASPxPageControl ID="ASPxPageControl1" runat="server" ActiveTabIndex="1" Width="100%"
                                                                    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                                                                    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                    TabSpacing="0px" onactivetabchanged="ASPxPageControl2_ActiveTabChanged">
                                                                    <TabPages>
                                                                        <dx:TabPage Text="Cabecera">
                                                                            <ContentCollection>
                                                                                <dx:ContentControl runat="server" SupportsDisabledAttribute="True">

                                                                                    <table width="100%">
                                                                                       <tr align="center">
                                                                                           <td>
                                                                                               <asp:Label ID="Labelx" runat="server" ForeColor="Red"></asp:Label>
                                                                                           </td>
                                                                                       </tr>
                                                                                    </table>
                                                                                    <table width="100%">

                                                                                        <tr>
                                                                                            <td align="right" width="30%">
                                                                                                <asp:Literal ID="Espacio" runat="server" Text="Espacio"></asp:Literal>
                                                                                            </td>
                                                                                            <td align="left">
                                                                                                <dx:ASPxSpinEdit ID="spEspacio" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                    CssPostfix="Office2010Silver" Height="21px" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                    Width="100%">
                                                                                                    <SpinButtons HorizontalSpacing="0">
                                                                                                    </SpinButtons>
                                                                                                </dx:ASPxSpinEdit>
                                                                                            </td>
                                                                                        </tr>

                                                                                        <tr>
                                                                                            <td align="right" width="30%">
                                                                                                <asp:Literal ID="Medio" runat="server" Text="Medio"></asp:Literal>
                                                                                            </td>
                                                                                            <td align="left">
                                                                                                <dx:ASPxSpinEdit ID="spMedio" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                    CssPostfix="Office2010Silver" Height="21px" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                    Width="100%">
                                                                                                    <SpinButtons HorizontalSpacing="0">
                                                                                                    </SpinButtons>
                                                                                                </dx:ASPxSpinEdit>
                                                                                            </td>
                                                                                        </tr>

                                                                                        <tr>
                                                                                            <td align="right" width="30%">
                                                                                                <asp:Literal ID="AnioMes" runat="server" Text="Año-Mes"></asp:Literal>
                                                                                            </td>
                                                                                            <td align="left">
                                                                                                <dx:ASPxSpinEdit ID="spAnioMes" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                    CssPostfix="Office2010Silver" Height="21px" ReadOnly="True" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                    Width="100%">
                                                                                                    <SpinButtons HorizontalSpacing="0">
                                                                                                    </SpinButtons>
                                                                                                </dx:ASPxSpinEdit>
                                                                                            </td>
                                                                                        </tr>

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
                                                                                                <asp:Literal ID="OrigenCertificado" runat="server" Text="Origen de Certificado"></asp:Literal>
                                                                                            </td>
                                                                                            <td align="left">
                                                                                                <dx:ASPxSpinEdit ID="spOrigenCertificado" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                    ReadOnly="True" Width="100%" EditFormat="DateTime">
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
                                                                                                <asp:Literal ID="Literal11" runat="server" Text="Usuario Certificado Válido"></asp:Literal>
                                                                                            </td>
                                                                                            <td>
                                                                                                <dx:ASPxTextBox ID="txUsuCierre" runat="server" ReadOnly="True" Width="100%">
                                                                                                </dx:ASPxTextBox>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr>
                                                                                            <td align="right" width="30%">
                                                                                                <asp:Literal ID="Literal12" runat="server" Text="Fecha Certificado Válido"></asp:Literal>
                                                                                            </td>
                                                                                            <td align="left">
                                                                                                <dx:ASPxDateEdit ID="deFecCierre" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                    CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                    ReadOnly="True" Width="100%" EditFormat="DateTime" 
                                                                                                    OnDateChanged="deFecCierre_DateChanged">
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
                                                                        <dx:TabPage Text="Detalle" Name="Detalle">
                                                                            <ContentCollection>
                                                                                <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                                                                    <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                                                                        <ContentTemplate>
                                                                                            <table width="100%">
                                                                                            <tr>
                                                                                                <td>
                                                                                                    <asp:Label ID="lblErrorLineas" runat="server" ForeColor="Red"></asp:Label>
                                                                                                </td>
                                                                                            </tr>
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
                                                                                                        <dx:ASPxPageControl ID="ASPxPageControl2" runat="server" ActiveTabIndex="1" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                            CssPostfix="Office2010Silver" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                            TabSpacing="0px" Width="100%" AutoPostBack="True" 
                                                                                                            OnActiveTabChanged="ASPxPageControl2_ActiveTabChanged">
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
                                                                                                                                            Width="100%" DisplayFormatString="HH:mm" EditFormat="Custom" 
                                                                                                                                            EditFormatString="HH:mm">
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
                                                                                                                                            Width="100%" DisplayFormatString="HH:mm" EditFormat="Custom" 
                                                                                                                                            EditFormatString="HH:mm" AutoPostBack="True" 
                                                                                                                                            OnValueChanged="teHoraFinInsertar_ValueChanged" 
                                                                                                                                            OnDateChanged="teHoraFinInsertar_DateChanged">
                                                                                                                                        </dx:ASPxTimeEdit>
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
                                                                                                                                    <td align="right">
                                                                                                                                        <asp:Literal ID="Literal19" runat="server" Text="Número de Salida"></asp:Literal>
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
                                                                                                                                                    <dx:ASPxButton ID="btnInsertarLineas" runat="server" OnClick="btnGenerarLineas_Click"
                                                                                                                                                        Text="Insertar Líneas" Width="150px">
                                                                                                                                                        <Image Url="~/Images/Icons/18_addView.gif">
                                                                                                                                                        </Image>
                                                                                                                                                    </dx:ASPxButton>
                                                                                                                                                </td>
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxButton ID="btnCancelInsertar" runat="server" OnClick="btnCancel_Click" Text="Cancelar"
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
                                                                                                                                            Width="20%">
                                                                                                                                            <CalendarProperties>
                                                                                                                                                

<HeaderStyle Spacing="1px" />
                                                                                                                                            

</CalendarProperties>
                                                                                                                                            <ButtonStyle Width="20%">
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
                                                                                                                                            Width="20%">
                                                                                                                                            <CalendarProperties>
                                                                                                                                                

<HeaderStyle Spacing="1px" />
                                                                                                                                            

</CalendarProperties>
                                                                                                                                            <ButtonStyle Width="20%">
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
                                                                                                                                            Width="20%">
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
                                                                                                                                        <asp:Literal ID="Literal39" runat="server" Text="Fecha Hasta"></asp:Literal>
                                                                                                                                    </td>
                                                                                                                                    <td align="left">
                                                                                                                                        <dx:ASPxDateEdit ID="deFechaHastaDestinoCopiar" runat="server" 
                                                                                                                                            CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                            CssPostfix="Office2010Silver" Spacing="0" 
                                                                                                                                            SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                                                            Width="20%">
                                                                                                                                            <CalendarProperties>
                                                                                                                                                <HeaderStyle Spacing="1px" />
                                                                                                                                            </CalendarProperties>
                                                                                                                                            <ButtonStyle Width="20%">
                                                                                                                                            </ButtonStyle>
                                                                                                                                        </dx:ASPxDateEdit>
                                                                                                                                    </td>
                                                                                                                                </tr>

                                                                                                                               
                                                                                                                                <tr>
                                                                                                                                    <td align="right" colspan="2">
                                                                                                                                        <table>
                                                                                                                                            <tr>
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxButton ID="btnCopiarPeriodos" runat="server" OnClick="btnCopiarPeriodos_Click"
                                                                                                                                                        Text="Copiar Períodos" Width="150px">
                                                                                                                                                        <Image Url="~/Images/Icons/18_editForm.gif">
                                                                                                                                                        </Image>
                                                                                                                                                    </dx:ASPxButton>
                                                                                                                                                </td>
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxButton ID="btnCancelCopiar" runat="server" OnClick="btnCancel_Click" Text="Cancelar"
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
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxRadioButton ID="opEditPeriodo" runat="server" Text="Período Ingresado" Checked="True"
                                                                                                                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                                                                                                                                                        GroupName="editDesde" 
                                                                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                                                                        OnCheckedChanged="opEditPeriodo_CheckedChanged">
                                                                                                                                                    </dx:ASPxRadioButton>
                                                                                                                                                </td>
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxRadioButton ID="opEditSeleccionados" runat="server" Text="Todos los Seleccionados"
                                                                                                                                                        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" CssPostfix="Office2010Silver"
                                                                                                                                                        GroupName="editDesde" 
                                                                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                                                                        OnCheckedChanged="opEditSeleccionados_CheckedChanged">
                                                                                                                                                    </dx:ASPxRadioButton>
                                                                                                                                                </td>
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxRadioButton ID="opEditTodas" runat="server" Text="Todas las Líneas" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                                        CssPostfix="Office2010Silver" GroupName="editDesde" 
                                                                                                                                                        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                                                                        OnCheckedChanged="opEditTodas_CheckedChanged" OnValueChanged="opEditTodas_CheckedChanged">
                                                                                                                                                    </dx:ASPxRadioButton>
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                        </table>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr>
                                                                                                                                    <td align="right">
                                                                                                                                        <asp:Literal ID="Literal22" runat="server" Text="Fecha y Hora Desde"></asp:Literal>
                                                                                                                                    </td>
                                                                                                                                    <td align="left">
                                                                                                                                        <dx:ASPxDateEdit ID="deFechaDesdeReemplazar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                            Width="20%">
                                                                                                                                            <CalendarProperties>
                                                                                                                                                

<HeaderStyle Spacing="1px" />
                                                                                                                                            

</CalendarProperties>
                                                                                                                                            <ButtonStyle Width="13px">
                                                                                                                                            </ButtonStyle>
                                                                                                                                        </dx:ASPxDateEdit>
                                                                                                                                        <dx:ASPxTimeEdit ID="deHoraDesdeOrigenReemplazar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                            CssPostfix="Office2010Silver" DisplayFormatString="HH:mm" EditFormat="Custom"
                                                                                                                                            EditFormatString="HH:mm" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                            Width="20%">
                                                                                                                                        </dx:ASPxTimeEdit>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr>
                                                                                                                                    <td align="right">
                                                                                                                                        <asp:Literal ID="Literal23" runat="server" Text="Fecha y Hora Hasta"></asp:Literal>
                                                                                                                                    </td>
                                                                                                                                    <td align="left">
                                                                                                                                        <dx:ASPxDateEdit ID="deFechaHastaReemplazar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                            CssPostfix="Office2010Silver" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                            Width="20%">
                                                                                                                                            <CalendarProperties>
                                                                                                                                                

<HeaderStyle Spacing="1px" />
                                                                                                                                            

</CalendarProperties>
                                                                                                                                            <ButtonStyle Width="13px">
                                                                                                                                            </ButtonStyle>
                                                                                                                                        </dx:ASPxDateEdit>
                                                                                                                                        <dx:ASPxTimeEdit ID="deHoraHastaOrigenReemplazar" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                                                            CssPostfix="Office2010Silver" DisplayFormatString="HH:mm" EditFormat="Custom"
                                                                                                                                            EditFormatString="HH:mm" Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css"
                                                                                                                                            Width="20%">
                                                                                                                                        </dx:ASPxTimeEdit>
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
                                                                                                                                    <td align="right">
                                                                                                                                        <asp:Literal ID="Literal38" runat="server" Text="Duración"></asp:Literal>
                                                                                                                                    </td>
                                                                                                                                    <td>
                                                                                                                                        <dx:ASPxSpinEdit ID="spAvisoReempDuracion" runat="server" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
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
                                                                                                                                                    <dx:ASPxButton ID="btnReemplazarAvisos" runat="server" OnClick="btnReemplazarAvisos_Click"
                                                                                                                                                        Text="Reemplazar Avisos" Width="150px">
                                                                                                                                                        <Image Url="~/Images/Icons/16_runworkflow.gif">
                                                                                                                                                        </Image>
                                                                                                                                                    </dx:ASPxButton>
                                                                                                                                                </td>
                                                                                                                                                <td>
                                                                                                                                                    <dx:ASPxButton ID="btnCancelReemplazar" runat="server" OnClick="btnCancel_Click"
                                                                                                                                                        Text="Cancelar" Width="150px">
                                                                                                                                                        <Image Url="~/Images/Crud/Delete_16.png">
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
                                                                                                            <paddings padding="2px" paddingleft="5px" 
                                                                                                                paddingright="5px" />
                                                                                                            <paddings padding="2px" paddingleft="5px" 
                                                                                                                paddingright="5px" />
                                                                                                            <ContentStyle>
                                                                                                                <Paddings Padding="12px" />
                                                                                                                <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px" />
                                                                                                                <Paddings Padding="12px" />
                                                                                                                <paddings padding="12px" />
                                                                                                                <paddings padding="12px" />
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
                                                                                                            <contentpaddings paddingbottom="10px" paddingleft="9px" 
                                                                                                                paddingright="11px" paddingtop="10px" />
                                                                                                            <contentpaddings paddingbottom="10px" paddingleft="9px" 
                                                                                                                paddingright="11px" paddingtop="10px" />
                                                                                                            <HeaderStyle>
                                                                                                                <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" PaddingTop="3px" />
                                                                                                            <paddings paddingbottom="6px" paddingleft="9px" paddingright="11px" 
                                                                                                                paddingtop="3px" />
                                                                                                            <paddings paddingbottom="6px" paddingleft="9px" paddingright="11px" 
                                                                                                                paddingtop="3px" />
                                                                                                            </HeaderStyle>
                                                                                                            <PanelCollection>
                                                                                                                <dx:PanelContent runat="server" SupportsDisabledAttribute="True">
                                                                                                                    <table width="100%">
                                                                                                                        <tr>
                                                                                                                            <td align="right">
                                                                                                                                <asp:Literal ID="Literal25" runat="server" Text="Aviso"></asp:Literal>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <uc1:ucComboBox ID="ucIdentifAvisoEdit" runat="server" />
                                                                                                                            </td>
                                                                                                                        </tr>
                                                                                                                        <tr>
                                                                                                                            <td align="right">
                                                                                                                                <asp:Literal ID="Literal36" runat="server" Text="Hora de Inicio"></asp:Literal>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <dx:ASPxTimeEdit ID="teHoraInicioModificar" runat="server" 
                                                                                                                                    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                    CssPostfix="Office2010Silver" DisplayFormatString="HH:mm" EditFormat="Custom" 
                                                                                                                                    EditFormatString="HH:mm" Spacing="0" 
                                                                                                                                    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%">
                                                                                                                                </dx:ASPxTimeEdit>
                                                                                                                            </td>
                                                                                                                        </tr>
                                                                                                                        <tr>
                                                                                                                            <td align="right">
                                                                                                                                <asp:Literal ID="Literal26" runat="server" Text="Duración"></asp:Literal>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <dx:ASPxSpinEdit ID="spDuracionEdit" runat="server" 
                                                                                                                                    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                    CssPostfix="Office2010Silver" Height="21px" Number="0" ReadOnly="True" 
                                                                                                                                    Spacing="0" SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
                                                                                                                                    Width="100%">
                                                                                                                                    <spinbuttons horizontalspacing="0">
                                                                                                                                    </spinbuttons>
                                                                                                                                </dx:ASPxSpinEdit>
                                                                                                                            </td>
                                                                                                                        </tr>
                                                                                                                        <tr>
                                                                                                                            <td align="right">
                                                                                                                                <asp:Literal ID="LitModifSalida" runat="server" Text="Número de Salida"></asp:Literal>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <dx:ASPxSpinEdit ID="spAvisoModifiSalidas" runat="server" 
                                                                                                                                    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
                                                                                                                                    CssPostfix="Office2010Silver" Height="21px" Number="0" Spacing="0" 
                                                                                                                                    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%">
                                                                                                                                    <spinbuttons horizontalspacing="0">
                                                                                                                                    </spinbuttons>
                                                                                                                                </dx:ASPxSpinEdit>
                                                                                                                            </td>
                                                                                                                        </tr>
                                                                                                                        <tr>
                                                                                                                            <td align="right" colspan="2">
                                                                                                                                <table>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <dx:ASPxButton ID="btnUpdateEdit" runat="server" OnClick="btnUpdateEdit_Click" Text="Aceptar"
                                                                                                                                                Width="150px">
                                                                                                                                                <Image Url="~/Images/Crud/16_save.gif">
                                                                                                                                                </Image>
                                                                                                                                            </dx:ASPxButton>
                                                                                                                                        </td>
                                                                                                                                        <td>
                                                                                                                                            <dx:ASPxButton ID="btnCancelEdit" runat="server" OnClick="btnCancelEdit_Click" Text="Cancelar"
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
                                                                                                                </dx:PanelContent>
                                                                                                            </PanelCollection>
                                                                                                        </dx:ASPxRoundPanel>
                                                                                                        &nbsp;
                                                                                                    </td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        <dx:ASPxGridView ID="gv" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
                                                                                                            CssPostfix="Office2010Silver" Width="100%" OnRowUpdating="gv_RowUpdating" OnStartRowEditing="gv_StartRowEditing" OnCustomColumnDisplayText="gv_CustomColumnDisplayText">
                                                                                                            <ClientSideEvents RowDblClick="function(s, e) { }" EndCallback="function(s, e) {

}" />
                                                                                                            <Columns>
                                                                                                                <dx:GridViewCommandColumn ShowSelectCheckbox="True" VisibleIndex="0" ButtonType="Image"
                                                                                                                    Width="40px">
                                                                                                                </dx:GridViewCommandColumn>
                                                                                                                <dx:GridViewDataTextColumn FieldName="Dia" VisibleIndex="0" Caption="Día" ReadOnly="True"></dx:GridViewDataTextColumn>
                                                                                                                <dx:GridViewDataTextColumn FieldName="DiaSemana" VisibleIndex="1" Caption="Día Semana" ReadOnly="True"></dx:GridViewDataTextColumn>
                                                                                                                <dx:GridViewDataTextColumn FieldName="Hora" VisibleIndex="2" ReadOnly="True" SortIndex="1" SortOrder="Ascending"></dx:GridViewDataTextColumn>
                                                                                                                <dx:GridViewDataTextColumn FieldName="Salida" VisibleIndex="3" ReadOnly="True"></dx:GridViewDataTextColumn>
                                                                                                                <dx:GridViewDataComboBoxColumn Caption="Aviso" FieldName="IdentifAviso" VisibleIndex="4"></dx:GridViewDataComboBoxColumn>
                                                                                                                <dx:GridViewDataTextColumn FieldName="Duracion" VisibleIndex="5" Caption="Duración" ReadOnly="True"></dx:GridViewDataTextColumn>
<%--                                                                                                                <dx:GridViewDataDateColumn Caption="Fecha" VisibleIndex="6" FieldName="Fecha" ReadOnly="True" SortIndex="0" SortOrder="Ascending"></dx:GridViewDataDateColumn>--%>
<%--                                                                                                                <dx:GridViewCommandColumn ShowSelectCheckbox="False" VisibleIndex="7" ButtonType="Image" Width="50px">
                                                                                                                    <EditButton><Image Url="~/Images/Crud/EditProperties_16.png"></Image></EditButton>
                                                                                                                    <CancelButton><Image Url="~/Images/Crud/Delete_16.png"></Image></CancelButton>
                                                                                                                    <UpdateButton><Image Url="~/Images/Crud/Save_16.png"></Image></UpdateButton>
                                                                                                                </dx:GridViewCommandColumn>--%>
                                                                                                            </Columns>
                                                                                                            <SettingsBehavior AutoExpandAllGroups="True" />
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
                                                                    <paddings padding="2px" paddingleft="5px" 
                                                                        paddingright="5px" />
                                                                    <ContentStyle>
                                                                        <Paddings Padding="12px" />
                                                                        <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px" />
                                                                        <Paddings Padding="12px" />
                                                                        <paddings padding="12px" />
                                                                        <paddings padding="12px" />
                                                                        <Border BorderColor="#868B91" BorderStyle="Solid" BorderWidth="1px"></Border>
                                                                    </ContentStyle>
                                                                </dx:ASPxPageControl>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr runat="server" id="trButtons">
                                                <td align="right">
                                                    <table align="right">
                                                        <tr>
                                                            <td runat="server" id ="tdValidar" align="left">
                                                            <dx:ASPxButton ID="btnValidar" runat="server" Text="Validar" OnClick="btnValidar_Click"
                                                            Width="150px" ToolTip="Validar">
                                                                <ClientSideEvents Click="function(s, e) {
                                                                
                                                                    e.processOnServer = confirm('Validación \n\n\n Confirma la Certificación para este Origen de certificado?');
                                                                
                                                                }" />
                                                            <Image Url="~/Images/Crud/ActivateQuote_16.png">
                                                            </Image>
                                                            </dx:ASPxButton>
                                                            </td>
                                                            <td runat="server" id ="tdCancel">
                                                            <dx:ASPxButton ID="btnBack2" runat="server" OnClick="btnBack_Click" Text="Volver"
                                                            Width="150px" ToolTip="Cancelar">
                                                            </dx:ASPxButton>
                                                            </td>
                                                            <td runat="server" id="tdSave">
                                                                <dx:ASPxButton ID="btnSave" runat="server" Text="Guardar Certificado" Width="150px"
                                                                    OnClick="btnSave_Click">
                                                                    <Image Url="~/Images/Crud/16_save.gif">
                                                                    </Image>
                                                                </dx:ASPxButton>
                                                            </td>
                                                            <td runat="server" id="tdAdd">
                                                                <dx:ASPxButton ID="btnAdd" runat="server" Text="Nuevo Certificado" 
                                                                    Width="150px" OnClick="btnAdd_Click" Visible="False">
                                                                    <Image Url="~/Images/Crud/cmd_add.png">
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
