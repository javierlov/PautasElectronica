using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PautasPublicidad.Web.Forms
{
    public partial class EstimadoConsultaVersion : System.Web.UI.Page
    {
        public Business.Estimado Estimado
        {
            get
            {
                if (Session["EstimadoAnulacionReemplazo.Estimado"] != null && Session["EstimadoAnulacionReemplazo.Estimado"] is Business.Estimado)
                {
                    return Session["EstimadoAnulacionReemplazo.Estimado"] as Business.Estimado;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                Session.Add("EstimadoAnulacionReemplazo.Estimado", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                gvCabecera.Settings.ShowHorizontalScrollBar = true;
                gvCabecera.SettingsBehavior.AllowSelectSingleRowOnly = true;
                gvCabecera.SettingsBehavior.AllowSelectByRowClick = true;

                gvDetalle.Settings.ShowHorizontalScrollBar = true;
                gvDetalle.SettingsBehavior.AllowSelectSingleRowOnly = true;
                gvDetalle.SettingsBehavior.AllowSelectByRowClick = true;

                gvSKUs.Settings.ShowHorizontalScrollBar = true;
                gvSKUs.SettingsBehavior.AllowSelectSingleRowOnly = true;
                gvSKUs.SettingsBehavior.AllowSelectByRowClick = true;
            }

            List<DTO.EstimadoCabVersionDTO> versiones = Business.Estimados.GetVersiones(Estimado.Cabecera.PautaId);

            gvCabecera.DataSource = versiones;
            gvCabecera.DataBind();

            btnRefreshSKU_Click(sender, e);

        }

        protected void btnRefreshSKU_Click(object sender, EventArgs e)
        {
            var selectedVersion = gvCabecera.GetSelectedFieldValues(new string[] { "Version" });

            if (selectedVersion != null && selectedVersion.Count > 0)
            {
                List<DTO.EstimadoDetVersionDTO> detalle = Business.Estimados.GetDetalles(Estimado.Cabecera.PautaId, Convert.ToInt32(selectedVersion[0]));
                List<DTO.EstimadoSKUVersionDTO> skus = Business.Estimados.GetSKUs(Estimado.Cabecera.PautaId, Convert.ToInt32(selectedVersion[0]));

                gvDetalle.DataSource = detalle;
                gvDetalle.DataBind();

                gvSKUs.DataSource = skus;
                gvSKUs.DataBind();
            }
        }
    }
}