using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web.Forms
{
    public partial class CertificadoBusqueda : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            ucIdentifEspacio.Inicializar(BusinessMapper.eEntities.EspacioCont);
            ucIdentifOrigen1.Inicializar(BusinessMapper.eEntities.Origen);
            ucIdentifOrigen2.Inicializar(BusinessMapper.eEntities.Origen);
        }

        protected void btnBuscarPauta_Click(object sender, EventArgs e)
        {
            try
            {
                var certificadoCab = Certificados.Buscar(txNroPauta.Text.Trim(), Convert.ToString(ucIdentifOrigen2.SelectedValue));

                if (certificadoCab != null)
                    Response.Redirect("Certificado.aspx?Certificado.RecId=" + certificadoCab.RecId.ToString(), true);
                else
                    lblMsg.Text = "No existe Pauta Certificada.";

            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void btnBuscarEspacioPeriodo_Click(object sender, EventArgs e)
        {
            try
            {
                var certificadoCab = Certificados.Buscar((string)ucIdentifEspacio.SelectedValue, Convert.ToInt32(seAño.Value), Convert.ToInt32(seMes.Value), Convert.ToString(ucIdentifOrigen2.SelectedValue));

                if (certificadoCab != null)
                    Response.Redirect("Certificado.aspx?Certificado.RecId=" + certificadoCab.RecId.ToString(), true);
                else
                    lblMsg.Text = "No existe Pauta Certificada.";
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }
    }
}