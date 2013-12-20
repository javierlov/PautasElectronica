//PLANILLA DE PAUTAS MENSUAL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;
using System.Xml;
using DevExpress.XtraPrinting;
using DevExpress.Web.ASPxGridView.Export;
using DevExpress.Web.ASPxGridView;
using DevExpress.Web.ASPxMenu;
using PautasPublicidad.DTO;
using System.Globalization;
using System.Drawing;
using System.Web.UI.HtmlControls;
using System.Reflection;
using DevExpress.Web.ASPxEditors;
using PautasPublicidad.Web.Controls;
using System.IO;
using System.Data;
using PautasPublicidad.DAO;
using PautasPublicidad.Web;

namespace PautasPublicidad.Web.Forms
{
    public partial class BusquedaPautasMensuales2 : System.Web.UI.Page
    {
        #region Variables

        public static List<CertificadoCabDTO> certificados;
        public static List<EstimadoCabDTO> estimados;
        public static List<OrdenadoCabDTO> ordenados;

        public static string AnioMes = string.Empty;
        public static System.Data.DataTable Tabla = new DataTable();

        #endregion

        #region Eventos

        protected void Page_Load()
        {
            if (Page.IsPostBack)
            {
                if (ucEstado.Text == "")
                {
                    AnioMes = deAnoMes.Text.Replace("-", "");

                    ucOrigen.Visible = false;

                    CargarTiposDePautas();
                }
            }

        }

        protected void deAnoMes_DateChanged(object sender, EventArgs e)
        {
            try {

                if (!Page.IsPostBack)
                {
                    if (ucEstado.Text != "")
                    {
                        AnioMes = deAnoMes.Text.Replace("-", "");

                        ucOrigen.Visible = false;

                        CargarTiposDePautas();

                    }
                }

            }
            catch (Exception ex) 
            { 
            }
        }

        protected void ucEstado_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch (ucEstado.Text)
            {
                case "Ordenado":case "Estimado":
                    {
                        lblOrigen.Visible = ucOrigen.Visible = false;
                        btnBuscar.Enabled = true;
                        break;
                    }
                case "Certificado":                    
                    {
                        CargarIdentificadores();
                        ucOrigen.Text = "";
                        lblOrigen.Visible = ucOrigen.Visible = true;
                        btnBuscar.Enabled = false;
                        break;
                    }
            }
        }

        protected void ucOrigen_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                switch (ucOrigen.Text)
                {
                    case "": break;
                    default:
                        {
                            btnBuscar.Enabled = true;

                            break;
                        }
                }
            }
            catch (Exception ex)
            {
            }

        }

        #endregion

        #region Botones

        protected void btnBuscar_Click(object sender, EventArgs e)
        {
            try {

                Response.Redirect("Copy of PlanillaPautasMensuales.aspx?Estado0=" + ucEstado.Text + "&Origen0=" + ucOrigen.Text + "&AnioMes0=" + deAnoMes.Text + "&Mostrado0=0");
            
            }
            catch(Exception ex)
            {
                string a = string.Empty;
                    a = ex.ToString();
            }
        }


        #endregion

        #region Procesos

        private void CargarTiposDePautas()
        {
            try 
            {
                certificados = Certificados.ReadAll("ANOMES = '" + AnioMes + "' AND NOT IDENTIFORIGEN IS NULL ");

                ucEstado.Items.Clear();

                if (certificados.Count > 0)
                {
                    ucEstado.Items.Add("Certificado");
                }

                estimados = Estimados.ReadAll(" ANOMES = '" + AnioMes + "'");

                if (estimados.Count > 0)
                {
                    ucEstado.Items.Add("Estimado");
                }

                ordenados = Ordenados.ReadAll(" ANOMES = '" + AnioMes + "' ");

                if (ordenados.Count > 0)
                {
                    ucEstado.Items.Add("Ordenado");
                }
            
            }
            catch (Exception ex) 
            { 
            }
        }

        private void CargarIdentificadores()
        {
            List<string> algo = new List<string>();

            ucOrigen.Items.Clear();

            for (int x = 0; x <= certificados.Count - 1; x++)
            {
                if (algo.Contains(certificados[x].IdentifOrigen) == false)
                {
                    algo.Add(certificados[x].IdentifOrigen);
                }
            }

            algo.Sort();

            for (int j = 0; j <= algo.Count - 1; j++)
            {
                ucOrigen.Items.Add(algo[j].ToString().ToUpper());
            }

        }

        #endregion

        protected void btnRefresh_Click(object sender, ImageClickEventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (ucEstado.Text != "")
                {
                    AnioMes = deAnoMes.Text.Replace("-", "");

                    ucOrigen.Visible = false;

                    CargarTiposDePautas();

                }
            }
        }
    }
}