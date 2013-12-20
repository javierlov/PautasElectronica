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
    public partial class PlanillaPautasMensual2 : System.Web.UI.Page
    {

        static System.Data.DataTable Tabla = null;
        public static string AnioMes = string.Empty;
        public static string Estado = string.Empty;
        public static string Origen = string.Empty;
        public static int Mostrado = 0;

        #region Eventos

        protected void Page_Load()
        {
            try {

                if (Page.IsPostBack)
                {
                    Mostrado = 1;
                    gv.DataSource = Tabla;
                    gv.DataBind();
                }

                if (!Page.IsPostBack)
                {
                    Mostrado = 0;
                }

                if (Mostrado == 0)
                {
                    AnioMes = Request.QueryString["AnioMes0"];
                    Estado = Request.QueryString["Estado0"];
                    Origen = Request.QueryString["Origen0"];
                    Mostrado = Convert.ToInt32(Request.QueryString["Mostrado0"]);

                    ArmarCabecera();

                    DataRow f1 = Tabla.NewRow();
                    Tabla.Rows.Add(f1);
                    AgregarTotales();

                    DataRow f2 = Tabla.NewRow();
                    Tabla.Rows.Add(f2);
                    AgregarPorcentajes();

                    Mostrado = 0;

                }

                ASPxGridViewExporter1.GridViewID = "gv";

                ASPxGridViewExporter1.DataBind();

            
            }
            catch (Exception ex)
            {
            }

        }

        protected void mnuPrincipal_ItemClick(object source, MenuItemEventArgs e)
        {

            try
            {
                switch (e.Item.Name)
                {
                    case "btnExport":
                        {
                            ASPxGridViewExporter1.GridViewID = "gv";
                            
                            XlsExportOptions xlsOpciones = new XlsExportOptions(TextExportMode.Text, true, false);

                            ASPxGridViewExporter1.DataBind();

                            ASPxGridViewExporter1.WriteXlsToResponse(xlsOpciones);

                            break;
                        }
                    case "btnLimpiar":
                        {
                            gv.DataSource = null;
                            gv.DataBind();

                            mnuPrincipal.Items[0].Enabled = false;

                            break;
                        }

                    default: break;
                }

            }
            catch (Exception ex)
            {
            }
        }

        #endregion


        #region Procesos

        protected void ArmarCabecera()
        {

            Tabla = new System.Data.DataTable();

            AnioMes = AnioMes.Replace("-", "");

            switch (Estado)
            {
                case "Ordenado":
                    {

                        OrdenadoCabDAO ocd = new OrdenadoCabDAO();

                        ocd.LlenarTabla(Tabla, "execute pa_pmpo1 '" + AnioMes + "'");
                        ocd.LlenarTabla(Tabla, "execute pa_pmpo2 '" + AnioMes + "'");
                        ocd.LlenarTabla(Tabla, "execute pa_pmpo3 '" + AnioMes + "'");

                        break;
                    }

                case "Estimado":
                    {

                        EstimadoCabDAO ecd = new EstimadoCabDAO();

                        ecd.LlenarTabla(Tabla, "execute pa_pmpe1 '" + AnioMes + "'");
                        ecd.LlenarTabla(Tabla, "execute pa_pmpe2 '" + AnioMes + "'");
                        ecd.LlenarTabla(Tabla, "execute pa_pmpe3 '" + AnioMes + "'");

                        break;
                    }

                case "Certificado":
                    {

                        CertificadoCabDAO ccd = new CertificadoCabDAO();

                        ccd.LlenarTabla(Tabla, "execute pa_pmpc1 '" + AnioMes + "','" + Origen + "'");
                        ccd.LlenarTabla(Tabla, "execute pa_pmpc2 '" + AnioMes + "','" + Origen + "'");
                        ccd.LlenarTabla(Tabla, "execute pa_pmpc3 '" + AnioMes + "','" + Origen + "'");

                        break;
                    }

                default: { return; }
            }

            Tabla.Columns.Remove("PautaId1");
            Tabla.Columns.Remove("PautaId2");

        }

        protected void AgregarTotales()
        {
            try {

                int suma = 0;
                decimal suma2 = 0;
                decimal suma3 = 0;

                string Titulo = string.Empty;

                Tabla.Rows[Tabla.Rows.Count - 1][0] = "TOTALES";

                gv.DataSource = Tabla;
                gv.DataBind();

                for (int x = 5; x <= Tabla.Columns.Count - 1; x++)
                {

                    Titulo = gv.Columns[x].ToString();

                    if (Titulo.Substring(Titulo.Length - 1) == "1")
                    {
                        Titulo = Titulo.Substring(0, Titulo.Length - 1);

                        Titulo = Titulo.Replace(" - Salidas", "");

                        Titulo += " - Costo";

                        gv.Columns[x].Caption = Titulo;

                    }

                    if (gv.Columns[x].Index > 6 && 
                        gv.Columns[x].ToString().Substring(gv.Columns[x].ToString().Length - 1) != "1" 
                        && gv.Columns[x].ToString().Substring(gv.Columns[x].ToString().Length - 6) != " Costo"
                        && gv.Columns[x].ToString().Substring(gv.Columns[x].ToString().Length - 7) != "Salidas")
                    {

                        Titulo = gv.Columns[x].ToString() + " - Salidas";

                        gv.Columns[gv.Columns[x].Index].Caption = Titulo;

                    }

                    gv.Columns[x].HeaderStyle.Wrap = DevExpress.Utils.DefaultBoolean.True;

                    suma = 0;
                    suma2 = 0;
                    suma3 = 0;

                    for (int y = 0; y <= Tabla.Rows.Count - 2; y++)
                    {
                        if (Tabla.Rows[y][x] != System.DBNull.Value)
                        {

                            if (Titulo.Substring(Titulo.Length - 7) == "Salidas")
                            {
                                suma += Convert.ToInt32(Tabla.Rows[y][x]);
                            }
                            else
                            {
                                if (Titulo.Substring(Titulo.Length - 5) == "Costo")
                                {
                                    suma2 += Convert.ToDecimal(Tabla.Rows[y][x]);
                                }
                                else
                                {
                                    suma3 += Convert.ToDecimal(Tabla.Rows[y][x]);
                                }
                            }
                        }
                    }

                    if (Titulo.Substring(Titulo.Length - 7) == "Salidas")
                    {
                        Tabla.Rows[Tabla.Rows.Count - 1][x] = suma;
                    }
                    else
                    {
                        if (Titulo.Substring(Titulo.Length - 5) == "Costo")
                        {
                            Tabla.Rows[Tabla.Rows.Count - 1][x] = suma2;
                        }
                        else
                        {
                            if (Titulo.Substring(Titulo.Length - 9) == "INVERSION")
                            {
                                Tabla.Rows[Tabla.Rows.Count - 1][x] = suma3;
                            }
                        }
                    }
                }

            }
            catch(Exception ex)
            {
            }

        }

        protected void AgregarPorcentajes()
        {
            string Titulo = string.Empty;

            decimal TotalSalidas = 0;
            decimal TotalSKUS = 0;
            decimal ParcialSalidas = 0;
            decimal ParcialSKUS = 0;

            //Me paro en la fila de los TOTALES
            int Fila = Convert.ToInt32(Tabla.Rows.Count - 2);

            //A PARTIR DE LA COLUMNA 7
            for (int xSalidas = 7; xSalidas <= Tabla.Columns.Count - 1; xSalidas++ )
            {
                if (Tabla.Rows[Fila][xSalidas] != System.DBNull.Value)
                {
                    if (Tabla.Rows[Fila][0].ToString() == "TOTALES")
                    {
                        Titulo = gv.Columns[xSalidas].ToString();

                        if (Titulo.Substring(Titulo.Length - 7) == "Salidas")
                        {
                            TotalSalidas += Convert.ToDecimal(Tabla.Rows[Fila][xSalidas]);
                        }
                        else
                        {
                            TotalSKUS += Convert.ToDecimal(Tabla.Rows[Fila][xSalidas]);
                        }
                    }
                }
            }

            Tabla.Rows[Tabla.Rows.Count - 1][0] = "PORCENTAJES";

            for (int x = 7; x <= Tabla.Columns.Count - 1; x++)
            {
                if (gv.Columns.Count == 0)
                {
                    Titulo += Tabla.Columns[x].ToString().Substring(Tabla.Columns[x].ToString().Length - 1) == "1" ? Tabla.Columns[x].ToString() + " - Costo" : Tabla.Columns[x].ToString() + " - Salidas";
                }
                else
                {
                    Titulo = gv.Columns[x].ToString();
                }

                ParcialSalidas = 0;

                ParcialSKUS = 0;

                if (Titulo.Substring(Titulo.Length - 7) == "Salidas")
                {

                    if (Tabla.Rows[Fila][x] != System.DBNull.Value)
                    {
                        if (Tabla.Rows[Fila][0].ToString() == "TOTALES")
                        {
                            ParcialSalidas += Convert.ToDecimal(Tabla.Rows[Fila][x]);
                        }
                    }
                }
                else
                {
                    if (Tabla.Rows[Fila][x] != System.DBNull.Value)
                    {
                        if (Tabla.Rows[Fila][0].ToString() == "TOTALES")
                        {
                            ParcialSKUS += Convert.ToDecimal(Tabla.Rows[Fila][x]);
                        }
                    }
                }

                if (Titulo.Substring(Titulo.Length - 7) == "Salidas")
                {
                    Tabla.Rows[Tabla.Rows.Count - 1][x] = System.Math.Round(ParcialSalidas / TotalSalidas * 100, 3);
                }
                else
                {
                    Tabla.Rows[Tabla.Rows.Count - 1][x] = System.Math.Round(ParcialSKUS / TotalSKUS * 100, 3);
                }
            }
        }

        #endregion

        protected void btnVolver_Click(object sender, EventArgs e)
        {
            Response.Redirect("Copy of Copy of EstimadoBusqueda.aspx");
        }

    }
}