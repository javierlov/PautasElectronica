using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxGridView;
using DevExpress.Web.ASPxEditors;
using System.Data;
using PautasPublicidad.DTO;
using PautasPublicidad.Web.Controls;
using DevExpress.Web.ASPxMenu;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web
{
    public partial class TiposMediosPublicitarios : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
                FormsHelper.Inicializar(gv);

            ucABM1.ActualizarGrilla += new ucABM.ABMEventHandler(ucABM1_ActualizarGrilla);
            ucABM1.Inicializar(BusinessMapper.eEntities.TipoMediosPub);
            RefreshGrid(gv);
        }

        void ucABM1_ActualizarGrilla(object sender, ABMEventArgs e)
        {
            RefreshGrid(gv);
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            GridViewDataComboBoxColumn c = (GridViewDataComboBoxColumn)gv.Columns["IdentifTecno"];

            c.PropertiesComboBox.TextField = "Name";
            c.PropertiesComboBox.ValueField = "IdentifTecno";
            c.PropertiesComboBox.DataSource = Business.TecnologiaSoporte.ReadAll("");

            gv.DataSource = Business.MediosPublicitarios.ReadAllTipo("");
            gv.DataBind();
        }

        protected void mnuPrincipal_ItemClick(object source, MenuItemEventArgs e)
        {
            FormsHelper.ToolBarClick(ucABM1, e.Item.Name, gv, null);
        }
    }
}