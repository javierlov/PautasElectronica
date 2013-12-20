<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ucComboBox.ascx.cs"
    Inherits="PautasPublicidad.Web.Controls.ucComboBox" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallbackPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <table width="100%" style="border-spacing: 0px; border-width: 0px; padding: 0px; margin: 0px 0px 0px 0px;">
            <tr>
         
                <td>
                    <dx:ASPxComboBox ID="cmbEntidad" runat="server" Width="100%" IncrementalFilteringMode="Contains" 
                        onselectedindexchanged="cmbEntidad_SelectedIndexChanged">
                    </dx:ASPxComboBox>
                </td>
                <td style="width: 20px;" align="center">
                    <asp:ImageButton ID="btnRefresh" runat="server" ImageUrl="~/Images/Crud/16_L_refresh.gif"
                        OnClick="btnRefresh_Click" ToolTip="Actualizar" />
                </td>       <td style="width: 20px;" align="center">
                    <asp:ImageButton ID="btnAdd" runat="server" ImageUrl="~/Images/Crud/cmd_add.png"
                        ToolTip="Agregar" OnClientClick="window.open('pagina1.aspx' ,'','height=300', 'width=300');" />
                </td>
            </tr>
        </table>
    </ContentTemplate>
</asp:UpdatePanel>
