using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web.Forms
{
    public partial class EstimadoBusqueda : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ucIdentifEspacio.Inicializar(BusinessMapper.eEntities.EspacioCont);
        }

        protected void btnBuscarEspacioPeriodo_Click(object sender, EventArgs e)
        {
            try
            {
                var estimadoCab = Estimados.Buscar((string)ucIdentifEspacio.SelectedValue, Convert.ToInt32(seAño.Value), Convert.ToInt32(seMes.Value));

                if (estimadoCab != null)
                    Response.Redirect("EstimadoAnulacionReemplazo.aspx?Estimado.RecId=" + estimadoCab.RecId.ToString(), true);
                else
                    lblMsg.Text = "No existe Pauta Estimada.";

            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void btnBuscarPauta_Click(object sender, EventArgs e)
        {
            try
            {
                var estimadoCab = Estimados.Buscar(txNroPauta.Text.Trim());

                if (estimadoCab != null)
                    Response.Redirect("EstimadoAnulacionReemplazo.aspx?Estimado.RecId=" + estimadoCab.RecId.ToString(), true);
                else
                    lblMsg.Text = "No existe Pauta Estimada.";
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void txNroPauta_TextChanged(object sender, EventArgs e)
        {
            btnBuscarPauta_Click(sender, e);
        }

        protected void seAño_NumberChanged(object sender, EventArgs e)
        {

        }

        protected void seMes_NumberChanged(object sender, EventArgs e)
        {

        }
    }
}