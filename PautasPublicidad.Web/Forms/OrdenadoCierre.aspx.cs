using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web.Forms
{
    public partial class OrdenadoCierre : System.Web.UI.Page
    {
        private System.Data.SqlClient.SqlTransaction tran;
        
        protected void Page_Load(object sender, EventArgs e)
        {
            seMes.Enabled = false;
            seAño.Enabled = false;
            int ano       = Ordenados.GetAñoMesCierreOrd();
            
            DateTime anomes   = new DateTime ( Convert.ToInt32(ano.ToString().Substring(0, 4)), Convert.ToInt32(ano.ToString().Substring(4, 2)), 1);
            anomes            = anomes.AddMonths(1);
            seAño.Value       = Convert.ToInt32(anomes.ToString("yyyy"));
            seMes.Value       = Convert.ToInt32(anomes.ToString("MM"));
            tblCerrar.Visible = false;
        }

        protected void seAñoMes_NumberChanged(object sender, EventArgs e)
        {
            deFechaCierre.Date = new DateTime(Convert.ToInt32(seAño.Number), Convert.ToInt32(seMes.Number), 1);
        }

        protected void btn_ShowCierre(object sender, EventArgs e)
        {
            lblMsg.Text = string.Empty;
            tblCerrar.Visible = true;
        }

        protected void btn_CerrarOrdenado(object sender, EventArgs e)
        {
            Business.Ordenado ordenado ;

            //preparo parametro añomes de referencia para guardar en setUp
            int ano = Ordenados.GetAñoMesCierreOrd();
            ano = Convert.ToInt32((seAño.Value).ToString() + (seMes.Value).ToString().PadLeft(2, '0'));


            try
            {
                lblMsg.Text = string.Empty;

                var ordenadosCab = Ordenados.ReadAll("AnoMes = " + seAño.Number.ToString("0000") + seMes.Number.ToString("00"));
                if (ordenadosCab.Count > 0)
                {
                    foreach (var ordenadoCab in ordenadosCab)
                    {
                        if (ordenadoCab.UsuCierre == string.Empty || ordenadoCab.UsuCierre == null)
                        {
                            ordenado = new Business.Ordenado(ordenadoCab.PautaId);
                            ordenado.ProcesoDeCierreOrdenado(((Accendo)this.Master).Usuario.UserName);
                            Ordenados.SetAñoMesCierreOrd(ano, tran);
                            Page_Load(sender, e);
                            lblMsg.Text += "Ordenado PautaId: " + ordenadoCab.PautaId + ": El cierre se realizó correctamente." + "<br />";
                            
                        }
                        else
                        {
                            Ordenados.SetAñoMesCierreOrd(ano, tran);
                            Page_Load(sender, e);
                            lblMsg.Text += "Ordenado PautaId: " + ordenadoCab.PautaId + ": Ya se encontraba cerrado." + "<br />";
                        }
                    }
                }
                else
                {
                    Ordenados.SetAñoMesCierreOrd(ano, tran);
                    Page_Load( sender,  e);
                    lblMsg.Text = "No se encontraron Ordenados para cerrar.";
                }


                tblCerrar.Visible = false;
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void btn_CancelarCierreOrdenado(object sender, EventArgs e)
        {
            tblCerrar.Visible = false;
            lblMsg.Text = string.Empty;
        }

     
    }
}