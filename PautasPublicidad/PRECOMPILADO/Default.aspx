<%@ Page Title="" Language="C#" MasterPageFile="~/PautasPublicidad.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="PautasPublicidad.Web.Default" %>
<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxMenu" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<table width="100%">
<tr>
<td>
<div width="100%">
    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/MenuItems.xml"
            XPath="/MenuItems/*"></asp:XmlDataSource>
    <dx:ASPxMenu ID="ASPxMenu3" runat="server" AutoSeparators="RootOnly" 
        CssFilePath="~/App_Themes/Office2010Blue/{0}/styles.css" 
        CssPostfix="Office2010Blue" DataSourceID="XmlDataSource1" 
        ShowPopOutImages="True" 
        SpriteCssFilePath="~/App_Themes/Office2010Blue/{0}/sprite.css" 
        Width="100%">
        <LoadingPanelImage Url="~/App_Themes/Office2010Blue/Web/Loading.gif">
        </LoadingPanelImage>
        <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
        <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
        <LoadingPanelStyle ImageSpacing="5px">
        </LoadingPanelStyle>
        <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
    </dx:ASPxMenu>
    
    <dx:ASPxGridView ID="ASPxGridView1" runat="server" 
        CssFilePath="~/App_Themes/Office2010Blue/{0}/styles.css" 
        CssPostfix="Office2010Blue" Width="100%">
        <SettingsPager AlwaysShowPager="True">
        </SettingsPager>
        <Images SpriteCssFilePath="~/App_Themes/Office2010Blue/{0}/sprite.css">
            <LoadingPanelOnStatusBar Url="~/App_Themes/Office2010Blue/GridView/Loading.gif">
            </LoadingPanelOnStatusBar>
            <LoadingPanel Url="~/App_Themes/Office2010Blue/GridView/Loading.gif">
            </LoadingPanel>
        </Images>
        <ImagesFilterControl>
            <LoadingPanel Url="~/App_Themes/Office2010Blue/GridView/Loading.gif">
            </LoadingPanel>
        </ImagesFilterControl>
        <Styles CssFilePath="~/App_Themes/Office2010Blue/{0}/styles.css" 
            CssPostfix="Office2010Blue">
            <Header ImageSpacing="5px" SortingImageSpacing="5px">
            </Header>
            <LoadingPanel ImageSpacing="5px">
            </LoadingPanel>
        </Styles>
        <StylesPager>
            <PageNumber ForeColor="#3E4846">
            </PageNumber>
            <Summary ForeColor="#1E395B">
            </Summary>
        </StylesPager>
        <StylesEditors ButtonEditCellSpacing="0">
            <ProgressBar Height="21px">
            </ProgressBar>
        </StylesEditors>
    </dx:ASPxGridView>
    <br />
    <dx:ASPxMenu ID="ASPxMenu2" runat="server" AutoSeparators="RootOnly" 
        CssFilePath="~/App_Themes/Aqua/{0}/styles.css" 
        CssPostfix="Aqua" DataSourceID="XmlDataSource1" 
        ShowPopOutImages="True" 
        SpriteCssFilePath="~/App_Themes/Aqua/{0}/sprite.css" Width="100%" 
        GutterImageSpacing="7px">
        <LoadingPanelImage Url="~/App_Themes/Aqua/Web/Loading.gif">
        </LoadingPanelImage>
        <RootItemSubMenuOffset FirstItemX="-1" FirstItemY="-1" X="-1" Y="-1" />
        <ItemStyle DropDownButtonSpacing="12px" PopOutImageSpacing="18px" 
            ToolbarDropDownButtonSpacing="5px" ToolbarPopOutImageSpacing="5px" 
            VerticalAlign="Middle" />
        <SubMenuStyle GutterWidth="0px" />
    </dx:ASPxMenu>
    
    <dx:ASPxGridView ID="ASPxGridView2" runat="server" 
        CssFilePath="~/App_Themes/Aqua/{0}/styles.css" CssPostfix="Aqua" 
        Width="100%">
        <SettingsPager AlwaysShowPager="True">
        </SettingsPager>
        <SettingsLoadingPanel ImagePosition="Top" />
        <Images SpriteCssFilePath="~/App_Themes/Aqua/{0}/sprite.css">
            <LoadingPanelOnStatusBar Url="~/App_Themes/Aqua/GridView/gvLoadingOnStatusBar.gif">
            </LoadingPanelOnStatusBar>
            <LoadingPanel Url="~/App_Themes/Aqua/GridView/Loading.gif">
            </LoadingPanel>
        </Images>
        <ImagesEditors>
            <DropDownEditDropDown>
                <SpriteProperties HottrackedCssClass="dxEditors_edtDropDownHover_Aqua" 
                    PressedCssClass="dxEditors_edtDropDownPressed_Aqua" />
            </DropDownEditDropDown>
            <SpinEditIncrement>
                <SpriteProperties HottrackedCssClass="dxEditors_edtSpinEditIncrementImageHover_Aqua" 
                    PressedCssClass="dxEditors_edtSpinEditIncrementImagePressed_Aqua" />
            </SpinEditIncrement>
            <SpinEditDecrement>
                <SpriteProperties HottrackedCssClass="dxEditors_edtSpinEditDecrementImageHover_Aqua" 
                    PressedCssClass="dxEditors_edtSpinEditDecrementImagePressed_Aqua" />
            </SpinEditDecrement>
            <SpinEditLargeIncrement>
                <SpriteProperties HottrackedCssClass="dxEditors_edtSpinEditLargeIncImageHover_Aqua" 
                    PressedCssClass="dxEditors_edtSpinEditLargeIncImagePressed_Aqua" />
            </SpinEditLargeIncrement>
            <SpinEditLargeDecrement>
                <SpriteProperties HottrackedCssClass="dxEditors_edtSpinEditLargeDecImageHover_Aqua" 
                    PressedCssClass="dxEditors_edtSpinEditLargeDecImagePressed_Aqua" />
            </SpinEditLargeDecrement>
        </ImagesEditors>
        <ImagesFilterControl>
            <LoadingPanel Url="~/App_Themes/Aqua/Editors/Loading.gif">
            </LoadingPanel>
        </ImagesFilterControl>
        <Styles CssFilePath="~/App_Themes/Aqua/{0}/styles.css" CssPostfix="Aqua">
            <LoadingPanel ImageSpacing="8px">
            </LoadingPanel>
        </Styles>
        <StylesEditors>
            <CalendarHeader Spacing="1px">
            </CalendarHeader>
            <ProgressBar Height="25px">
            </ProgressBar>
        </StylesEditors>
    </dx:ASPxGridView>
    <br />
    <dx:ASPxMenu ID="ASPxMenu1" runat="server" AutoSeparators="RootOnly" 
        CssFilePath="~/App_Themes/Office2010Black/{0}/styles.css" 
        CssPostfix="Office2010Black" DataSourceID="XmlDataSource1" 
        ShowPopOutImages="True" 
        SpriteCssFilePath="~/App_Themes/Office2010Black/{0}/sprite.css" Width="100%">
        <LoadingPanelImage Url="~/App_Themes/Office2010Black/Web/Loading.gif">
        </LoadingPanelImage>
        <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
        <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
        <LoadingPanelStyle ImageSpacing="5px">
        </LoadingPanelStyle>
        <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
    </dx:ASPxMenu>
    
    <dx:ASPxGridView ID="ASPxGridView4" runat="server" 
        CssFilePath="~/App_Themes/Office2010Black/{0}/styles.css" 
        CssPostfix="Office2010Black" Width="100%">
        <SettingsPager AlwaysShowPager="True">
        </SettingsPager>
        <Images SpriteCssFilePath="~/App_Themes/Office2010Black/{0}/sprite.css">
            <LoadingPanelOnStatusBar Url="~/App_Themes/Office2010Black/GridView/Loading.gif">
            </LoadingPanelOnStatusBar>
            <LoadingPanel Url="~/App_Themes/Office2010Black/GridView/Loading.gif">
            </LoadingPanel>
        </Images>
        <ImagesFilterControl>
            <LoadingPanel Url="~/App_Themes/Office2010Black/GridView/Loading.gif">
            </LoadingPanel>
        </ImagesFilterControl>
        <Styles CssFilePath="~/App_Themes/Office2010Black/{0}/styles.css" 
            CssPostfix="Office2010Black">
            <Header ImageSpacing="5px" SortingImageSpacing="5px">
            </Header>
            <LoadingPanel ImageSpacing="5px">
            </LoadingPanel>
        </Styles>
        <StylesPager>
            <CurrentPageNumber ForeColor="Black">
            </CurrentPageNumber>
            <PageNumber ForeColor="White">
            </PageNumber>
            <Summary ForeColor="White">
            </Summary>
            <Ellipsis ForeColor="White">
            </Ellipsis>
        </StylesPager>
        <StylesEditors ButtonEditCellSpacing="0">
            <ProgressBar Height="21px">
            </ProgressBar>
        </StylesEditors>
    </dx:ASPxGridView>
    <br />    
    <dx:ASPxMenu ID="ASPxMenu4" runat="server" AutoSeparators="RootOnly" 
        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
        CssPostfix="Office2010Silver" DataSourceID="XmlDataSource1" 
        ShowPopOutImages="True" 
        SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
        Width="100%">
        <LoadingPanelImage Url="~/App_Themes/Office2010Silver/Web/Loading.gif">
        </LoadingPanelImage>
        <ItemSubMenuOffset FirstItemX="2" LastItemX="2" X="2" />
        <ItemStyle DropDownButtonSpacing="10px" PopOutImageSpacing="10px" />
        <LoadingPanelStyle ImageSpacing="5px">
        </LoadingPanelStyle>
        <SubMenuStyle GutterImageSpacing="9px" GutterWidth="13px" />
    </dx:ASPxMenu>
    
    <dx:ASPxGridView ID="ASPxGridView3" runat="server" 
        CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
        CssPostfix="Office2010Silver" Width="100%">
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
    </div>
</td>
</tr>
</table>
</asp:Content>
