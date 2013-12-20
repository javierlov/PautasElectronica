using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web
{
    public partial class GridTest : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Conexion.SetConnectionString(@"Data Source=.;Initial Catalog=Publicidad;Integrated Security=True");

            gv.KeyFieldName = "RecId";
            RefreshGrid();
        }

        protected void gv_RowDeleting(object sender, DevExpress.Web.Data.ASPxDataDeletingEventArgs e)
        {
            e.Cancel = true;

            MediosPublicitarios.DeleteGrupo((int)e.Keys["RecId"]);
            RefreshGrid();
        }

        protected void gv_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            e.Cancel = true;

            var o = MediosPublicitarios.ReadGrupo((int)e.Keys["RecId"]);
            o.Name = (string)e.NewValues["Name"];
            o.IdentifGrupo = (string)e.NewValues["IdentifGrupo"];

            MediosPublicitarios.UpdateGrupo(o);
            RefreshGrid();
        }

        protected void gv_RowInserting(object sender, DevExpress.Web.Data.ASPxDataInsertingEventArgs e)
        {
            e.Cancel = true;

            MediosPublicitarios.CreateGrupo((string)e.NewValues["Name"], (string)e.NewValues["IdentifGrupo"]);
            RefreshGrid();
        }        
        
        private void RefreshGrid()
        {
            gv.DataSource = MediosPublicitarios.ReadAllGrupo("");
            gv.DataBind();
        }
    }
}