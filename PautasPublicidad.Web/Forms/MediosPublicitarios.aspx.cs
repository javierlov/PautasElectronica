using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxEditors;
using PautasPublicidad.DTO;
using DevExpress.Web.ASPxGridView;
using PautasPublicidad.Business;
using PautasPublicidad.Web.Controls;
using DevExpress.Web.ASPxMenu;

namespace PautasPublicidad.Web
{
    public partial class MediosPublicitarios : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
                FormsHelper.Inicializar(gv);
            
            ucABM1.ActualizarGrilla += new ucABM.ABMEventHandler(ucABM1_ActualizarGrilla);
            ucABM1.Inicializar(BusinessMapper.eEntities.MediosPub);
            RefreshGrid(gv);
        }

        void ucABM1_ActualizarGrilla(object sender, ABMEventArgs e)
        {
            RefreshGrid(gv);
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            GridViewDataComboBoxColumn c;

            c = (GridViewDataComboBoxColumn)gv.Columns["IdentifTipo"];
            c.PropertiesComboBox.TextField = "Name";
            c.PropertiesComboBox.ValueField = "IdentifTipo";
            c.PropertiesComboBox.DataSource = Business.MediosPublicitarios.ReadAllTipo("");

            c = (GridViewDataComboBoxColumn)gv.Columns["IdentifGrupo"];

            c.PropertiesComboBox.TextField = "Name";
            c.PropertiesComboBox.ValueField = "IdentifGrupo";
            c.PropertiesComboBox.DataSource = Business.MediosPublicitarios.ReadAllGrupo("");

            gv.DataSource = Business.MediosPublicitarios.ReadAll("");
            gv.DataBind();
        }

        protected void ASPxMenu1_ItemClick(object source, DevExpress.Web.ASPxMenu.MenuItemEventArgs e)
        {
            FormsHelper.ToolBarClick(ucABM1, e.Item.Name, gv, ASPxGridViewExporter1);
        }
    }
}