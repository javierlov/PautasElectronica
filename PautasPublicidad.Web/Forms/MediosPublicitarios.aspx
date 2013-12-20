<%@ Page Title="" Language="C#" MasterPageFile="~/Accendo.Master" AutoEventWireup="true"
    CodeBehind="MediosPublicitarios.aspx.cs" Inherits="PautasPublicidad.Web.MediosPublicitarios" %>

<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxGridView.v11.2.Export, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxMenu" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Src="../Controls/ucComboBox.ascx" TagName="ucComboBox" TagPrefix="uc1" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallbackPanel" TagPrefix="dx" %>
<%@ Register src="../Controls/ucABM.ascx" tagname="ucABM" tagprefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <div align="center" style="vertical-align: top; height: 100%; overflow: auto;">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <dx1:ASPxGridViewExporter ID="ASPxGridViewExporter1" runat="server" GridViewID="gv">
        </dx1:ASPxGridViewExporter>
        <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItems.xml"
            XPath="/MenuItems/*"></asp:XmlDataSource>
        <dx:ASPxMenu ID="ASPxMenu1" runat="server" AutoSeparators="RootOnly" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver" DataSourceID="XmlDataSource1" ShowPopOutImages="True"
            SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" Width="100%"
            OnItemClick="ASPxMenu1_ItemClick" AutoPostBack="True">
            <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif">
            </LoadingPanelImage>
            <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
            <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
            <LoadingPanelStyle ImageSpacing="5px">
            </LoadingPanelStyle>
            <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
        </dx:ASPxMenu>
        <uc2:ucABM ID="ucABM1" runat="server" Visible="False" />
        <dx:ASPxGridView ID="gv" runat="server" AutoGenerateColumns="False" CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css"
            CssPostfix="Office2010Silver"
            Width="100%">
            <Columns>
                <dx:GridViewCommandColumn VisibleIndex="0" ButtonType="Image">
                    <EditButton Text="Modificar">
                        <Image Url="~/Images/Crud/EditProperties_16.png" Width="16px">
                        </Image>
                    </EditButton>
                    <NewButton Text="Nuevo">
                        <Image Url="~/Images/Crud/cmd_add.png" Width="16px">
                        </Image>
                    </NewButton>
                    <DeleteButton Text="Eliminar">
                        <Image Url="~/Images/Crud/Delete_16.png" Width="16px">
                        </Image>
                    </DeleteButton>
                    <SelectButton Visible="True">
                        <Image Url="~/Images/Crud/AddActivity_16.png">
                        </Image>
                    </SelectButton>
                    <CancelButton Text="Cancelar">
                        <Image Url="~/Images/Crud/Delete_16.png">
                        </Image>
                    </CancelButton>
                    <UpdateButton Text="Guardar">
                        <Image Url="~/Images/Crud/Save_16.png">
                        </Image>
                    </UpdateButton>
                    <ClearFilterButton Visible="True" Text="Limpiar">
                        <Image Url="~/Images/Crud/Delete_16.png">
                        </Image>
                    </ClearFilterButton>
                </dx:GridViewCommandColumn>
                <dx:GridViewDataTextColumn Caption="Medio Publicitario" FieldName="IdentifMedio"
                    VisibleIndex="3">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="Descripción" FieldName="Name" VisibleIndex="4">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="RecId" FieldName="RecId" ReadOnly="True" Visible="False"
                    VisibleIndex="1">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="DatareaId" FieldName="DatareaId" ReadOnly="True"
                    Visible="False" VisibleIndex="2">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataComboBoxColumn Caption="Grupo de Medios" FieldName="IdentifGrupo"
                    VisibleIndex="5">
                    <PropertiesComboBox IncrementalFilteringMode="Contains">
                    </PropertiesComboBox>
                </dx:GridViewDataComboBoxColumn>
                <dx:GridViewDataComboBoxColumn Caption="Tipo de Medios" FieldName="IdentifTipo" VisibleIndex="6">
                    <PropertiesComboBox IncrementalFilteringMode="Contains">
                    </PropertiesComboBox>
                </dx:GridViewDataComboBoxColumn>
            </Columns>
            <SettingsBehavior ConfirmDelete="True" />
            <Settings ShowFilterRow="True" ShowGroupPanel="True" />
            <SettingsBehavior ConfirmDelete="True"></SettingsBehavior>
            <SettingsPager PageSize="20">
            </SettingsPager>
            <SettingsEditing Mode="EditForm"></SettingsEditing>
            <Settings ShowFilterRow="True"></Settings>
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
