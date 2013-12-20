using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;
using PautasPublicidad.Web.Controls;

namespace PautasPublicidad.Web
{
    public partial class GruposMediosPublicitarios : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
                FormsHelper.Inicializar(gv);

            ucABM1.ActualizarGrilla += new ucABM.ABMEventHandler(ucABM1_ActualizarGrilla);
            ucABM1.Inicializar(BusinessMapper.eEntities.GrupoMediosPub );
            RefreshGrid(gv);
        }

        void ucABM1_ActualizarGrilla(object sender, ABMEventArgs e)
        {
            RefreshGrid(gv);
        }

        private void RefreshGrid(DevExpress.Web.ASPxGridView.ASPxGridView gv)
        {

            gv.DataSource = Business.MediosPublicitarios.ReadAllGrupo("");
            gv.DataBind();
        }

        protected void ASPxMenu1_ItemClick(object source, DevExpress.Web.ASPxMenu.MenuItemEventArgs e)
        {
            FormsHelper.ToolBarClick(ucABM1, e.Item.Name, gv, ASPxGridViewExporter1);
        }
    }
}