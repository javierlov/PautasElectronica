using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PautasPublicidad.Web
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ASPxGridView1.SettingsBehavior.AllowSelectByRowClick = true;
            ASPxGridView1.DataSource = Business.MediosPublicitarios.ReadAll("");
            ASPxGridView1.DataBind();

            ASPxGridView2.SettingsBehavior.AllowSelectByRowClick = true;
            ASPxGridView2.DataSource = Business.MediosPublicitarios.ReadAll("");
            ASPxGridView2.DataBind();

            ASPxGridView3.SettingsBehavior.AllowSelectByRowClick = true;
            ASPxGridView3.DataSource = Business.MediosPublicitarios.ReadAll("");
            ASPxGridView3.DataBind();

            ASPxGridView4.SettingsBehavior.AllowSelectByRowClick = true;
            ASPxGridView4.DataSource = Business.MediosPublicitarios.ReadAll("");
            ASPxGridView4.DataBind();

        }
    }
}