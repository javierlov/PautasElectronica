using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PautasPublicidad.Web.Forms
{
    public partial class CertificadoConsulta : System.Web.UI.Page
    {
        public Business.Certificado Certificado
        {
            get
            {
                if (Session["Certificado.Certificado"] != null && Session["Certificado.Certificado"] is Business.Certificado)
                    return Session["Certificado.Certificado"] as Business.Certificado;
                else
                    return null;
            }
            set
            {
                Session.Add("Certificado.Certificado", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                gvCabecera.Settings.ShowHorizontalScrollBar          = true;
                gvCabecera.SettingsBehavior.AllowSelectSingleRowOnly = true;
                gvCabecera.SettingsBehavior.AllowSelectByRowClick    = true;

                gvDetalle.Settings.ShowHorizontalScrollBar           = true;
                gvDetalle.SettingsBehavior.AllowSelectSingleRowOnly  = true;
                gvDetalle.SettingsBehavior.AllowSelectByRowClick     = true;

                gvSKUs.Settings.ShowHorizontalScrollBar              = true;
                gvSKUs.SettingsBehavior.AllowSelectSingleRowOnly     = true;
                gvSKUs.SettingsBehavior.AllowSelectByRowClick        = true;
            }

            List<DTO.CertificadoCabDTO> cabeceras = Business.Certificados.GetOrigenes(Certificado.Cabecera.PautaId);

            gvCabecera.DataSource = cabeceras;
            gvCabecera.DataBind();

        }

        protected void btnRefreshSKU_Click(object sender, EventArgs e)
        {
            var selectedOrigen = gvCabecera.GetSelectedFieldValues(new string[] { "IdentifOrigen" });

            if (selectedOrigen != null && selectedOrigen.Count > 0)
            {
                List<DTO.CertificadoDetDTO> detalle = Business.Certificados.GetDetalles(Certificado.Cabecera.PautaId, Convert.ToString(selectedOrigen[0]));
                List<DTO.CertificadoSKUDTO> skus    = Business.Certificados.GetSKUs(Certificado.Cabecera.PautaId, Convert.ToString(selectedOrigen[0]));

                gvDetalle.DataSource = detalle;
                gvDetalle.DataBind();

                gvSKUs.DataSource = skus;
                gvSKUs.DataBind();
            }
        }
    }
}