<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="GridTest.aspx.cs" Inherits="PautasPublicidad.Web.GridTest" %>

<%@ Register assembly="DevExpress.Web.ASPxGridView.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.ASPxEditors.v11.2, Version=11.2.11.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    <dx:ASPxGridView ID="gv" runat="server" AutoGenerateColumns="False" 
            onrowdeleting="gv_RowDeleting" onrowinserting="gv_RowInserting" 
            onrowupdating="gv_RowUpdating">
        <Columns>
            <dx:GridViewCommandColumn VisibleIndex="0">
                <EditButton Visible="True">
                </EditButton>
                <NewButton Visible="True">
                </NewButton>
                <DeleteButton Visible="True">
                </DeleteButton>
                <ClearFilterButton Visible="True">
                </ClearFilterButton>
            </dx:GridViewCommandColumn>
            <dx:GridViewDataTextColumn Caption="Código del grupo " FieldName="IdentifGrupo" 
                VisibleIndex="3">
            </dx:GridViewDataTextColumn>
            <dx:GridViewDataTextColumn Caption="Nombre del Grupo" FieldName="Name" 
                VisibleIndex="4">
            </dx:GridViewDataTextColumn>
            <dx:GridViewDataTextColumn Caption="RecId" FieldName="RecId" ReadOnly="True" 
                Visible="False" VisibleIndex="1">
            </dx:GridViewDataTextColumn>
            <dx:GridViewDataTextColumn Caption="DatareaId" FieldName="DatareaId" 
                ReadOnly="True" Visible="False" VisibleIndex="2">
            </dx:GridViewDataTextColumn>
        </Columns>
        <SettingsBehavior ConfirmDelete="True" />
        <SettingsEditing Mode="EditForm" />
        <Settings ShowFilterRow="True" />
    </dx:ASPxGridView>
    
    </div>
    </form>
</body>
</html>
