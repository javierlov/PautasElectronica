using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web.Forms
{
    public partial class EstimadoCierre : System.Web.UI.Page
    {
        private System.Data.SqlClient.SqlTransaction tran;

        protected void Page_Load(object sender, EventArgs e)
        {
            seMes.Enabled     = false;
            seAño.Enabled     = false;
            int ano           = Ordenados.GetAñoMesCierreEst();
            DateTime anomes   = new DateTime(Convert.ToInt32(ano.ToString().Substring(0, 4)), Convert.ToInt32(ano.ToString().Substring(4, 2)), 1);
            anomes            = anomes.AddMonths(1);
            seAño.Value       = Convert.ToInt32(anomes.ToString("yyyy"));
            seMes.Value       = Convert.ToInt32(anomes.ToString("MM"));
            tblCerrar.Visible = false;
        }

        protected void btn_ShowCierre(object sender, EventArgs e)
        {
            lblMsg.Text       = string.Empty;
            tblCerrar.Visible = true;
        }

        protected void btnCerrarEstimado(object sender, EventArgs e)
        {
            Business.Estimado estimado;

            //preparo parametro añomes de referencia para guardar en setUp
            int ano = Ordenados.GetAñoMesCierreEst();
            ano     = Convert.ToInt32((seAño.Value).ToString() + (seMes.Value).ToString().PadLeft(2, '0'));

            try
            {
                lblMsg.Text = string.Empty;

                var estimadosCab = Estimados.ReadAll("AnoMes = " + seAño.Number.ToString("0000") + seMes.Number.ToString("00"));

                if (estimadosCab.Count > 0)
                {
                    foreach (var estimadoCab in estimadosCab)
                    {
                        if (estimadoCab.UsuCierre == string.Empty || estimadoCab.UsuCierre == null)
                        {
                            estimado = new Business.Estimado(estimadoCab.PautaId);

                            estimado.ProcesoDeCierreEstimado(((Accendo)this.Master).Usuario.UserName);

                            Ordenados.SetAñoMesCierreEst(ano, tran);

                            lblMsg.Text += "Estimado PautaId: " + estimadoCab.PautaId + ": El cierre se realizó correctamente." + "<br />";
                        }
                        else
                        {
                            lblMsg.Text += "Estimado PautaId: " + estimadoCab.PautaId + ": Ya se encontraba cerrado." + "<br />";
                        }
                    }
                }
                else
                {
                    lblMsg.Text = "No se encontraron Estimados para cerrar.";
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void btn_CancelarCierreEstimado(object sender, EventArgs e)
        {
            tblCerrar.Visible = false;
            lblMsg.Text       = string.Empty;
        }

    }
}