<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ucABM.ascx.cs" Inherits="PautasPublicidad.Web.Controls.ucABM" %>
<%@ Register assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" tagprefix="dx" %>
<%@ Register src="ucComboBox.ascx" tagname="ucComboBox" tagprefix="uc1" %>

<%@ Register assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>

<dx:ASPxRoundPanel ID="pnlMain" runat="server" 
    CssFilePath="~/App_Themes/Office2010Silver/{0}/styles.css" 
    CssPostfix="Office2010Silver" EnableDefaultAppearance="False" 
    GroupBoxCaptionOffsetX="6px" GroupBoxCaptionOffsetY="-19px" 
    SpriteCssFilePath="~/App_Themes/Office2010Silver/{0}/sprite.css" 
    Width="100%">
    <ContentPaddings PaddingBottom="10px" PaddingLeft="9px" PaddingRight="11px" 
        PaddingTop="10px" />
    <HeaderStyle>
    <Paddings PaddingBottom="6px" PaddingLeft="9px" PaddingRight="11px" 
        PaddingTop="3px" />
    </HeaderStyle>
    <PanelCollection>
<dx:PanelContent runat="server" SupportsDisabledAttribute="True">      <table runat="server" id="tblABM" width="100%" style="text-align: right;">
                        <tr runat="server" id="trMsg">
                        <td align="center">
                            <asp:Label ID="lblMsg" runat="server" 
                                Text="¿Está completamente seguro de que desea eliminar los registros seleccionados? Esta operación no puede deshacerse." 
                                Font-Bold="True" ForeColor="Blue"></asp:Label></td>
                        </tr>
                        <tr runat="server" id="trAbm">
                            <td>
                                <table runat="server" id="tblControls" width="100%">
                                    <tr>
                                        <td>
                                            XXX
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
                                            <dx:ASPxButton ID="btnCancel" runat="server" Text="Cancelar" Width="150px" 
                                                OnClick="btnCancel_Click">
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

                    </table></dx:PanelContent>
</PanelCollection>
</dx:ASPxRoundPanel>


