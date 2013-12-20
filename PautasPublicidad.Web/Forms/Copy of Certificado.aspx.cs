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
    public partial class PlanillaPautasMensual : System.Web.UI.Page
    {
        #region Variables

        static List<CertificadoCabDTO> certificados;
        static List<EstimadoCabDTO> estimados;
        static List<OrdenadoCabDTO> ordenados;

        static string AnioMes = string.Empty;
        static System.Data.DataTable Tabla = new DataTable();

        #endregion

        #region Eventos

        protected void Page_Load()
        {
        }


        protected void deAnoMes_DateChanged(object sender, EventArgs e)
        {
            AnioMes = deAnoMes.Text.Replace("-", "");

            gv.Visible = false;

            ucOrigen.Visible = false;

            CargarTiposDePautas();
        }

        protected void gv_CustomColumnDisplayText(object sender, ASPxGridViewColumnDisplayTextEventArgs e)
        {
            //if (e.Column.FieldName == "PautaId1" || e.Column.FieldName == "PautaId2")
            //{
            //    e.Column.Visible = false;

            //    return;
            //}

            //string Titulo = e.Column.FieldName;

            //if (Titulo.Substring(Titulo.Length - 1) == "1")
            //{
            //    Titulo = Titulo.Substring(0, Titulo.Length - 1);

            //    gv.Columns[Titulo + "1"].Caption = Titulo + " - Costo";

            //    gv.Columns[e.Column.Index].HeaderStyle.Wrap = DevExpress.Utils.DefaultBoolean.True;

            //    gv.DataBind();
            //}

            //if (e.Column.Index > 6 && e.Column.FieldName.Substring(e.Column.FieldName.Length - 1) != "1" && e.Column.FieldName.Substring(e.Column.FieldName.Length - 7) != "Salidas")
            //{
            //    Titulo = e.Column.FieldName + " - Salidas";

            //    gv.Columns[e.Column.Index].Caption = Titulo;

            //    gv.Columns[e.Column.Index].HeaderStyle.Wrap = DevExpress.Utils.DefaultBoolean.True;

            //    gv.DataBind();

            //}
        }

        protected void ucEstado_SelectedIndexChanged(object sender, EventArgs e)
        {
            gv.DataSource = null;
            gv.DataBind();
            gv.Visible = false;

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
            gv.Visible = false;

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

        protected void mnuPrincipal_ItemClick(object source, MenuItemEventArgs e)
        {

            try
            {
                switch (e.Item.Name)
                {
                    case "btnExport":
                        {
                            ASPxGridViewExporter1.GridViewID = "gv";

                            gv.DataSource = Tabla;

                            ASPxGridViewExporter1.DataBind();

                            if (ASPxGridViewExporter1 != null)
                            {
                                XlsExportOptions xlsExportOptions = new XlsExportOptions(TextExportMode.Text, true, false);

                                string Titulo = ucEstado.Text.Trim() + " - " + ucOrigen.Text.Trim() + " " + deAnoMes.Text;

                                xlsExportOptions.SheetName = Titulo;

                                ASPxGridViewExporter1.WriteXlsToResponse(xlsExportOptions);
                            }
                            break;
                        }
                    case "btnLimpiar":
                        {
                            gv.DataSource = null;
                            gv.DataBind();

                            break;
                        }

                    default: break;
                }

            }
            catch(Exception ex)
            { 
            }
        }

        #endregion

        #region Botones

        protected void btnBuscar_Click(object sender, EventArgs e)
        {
            try
            {
                gv.DataSource = null;
                gv.DataBind();

                gv.Visible = true;

                if (Tabla != null)
                {
                    Tabla.Reset();
                }

                ArmarCabecera(ucEstado.Text);

                //Tabla.Rows.Add(Tabla.NewRow());
                //AgregarTotales();

                //Tabla.Rows.Add(Tabla.NewRow());
                //AgregarPorcentajes();

            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void btnRefresh_Click(object sender, ImageClickEventArgs e)
        {
            gv.Visible = false;

            AnioMes = deAnoMes.Text.Replace("-", "");

            CargarTiposDePautas();
        }

        #endregion

        #region Procesos

        private void CargarTiposDePautas()
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

        protected void ArmarCabecera(string Estado)
        {

            Tabla = new System.Data.DataTable();

            AnioMes = deAnoMes.Text.Replace("-", "");

            gv.DataSource = null;
            gv.DataBind();
            gv.Columns.Clear();

            gv.Visible = false;

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

                        ccd.LlenarTabla(Tabla, "execute pa_pmpc1 '" + AnioMes + "','" + ucOrigen.Text + "'");
                        ccd.LlenarTabla(Tabla, "execute pa_pmpc2 '" + AnioMes + "','" + ucOrigen.Text + "'");
                        ccd.LlenarTabla(Tabla, "execute pa_pmpc3 '" + AnioMes + "','" + ucOrigen.Text + "'");

                        break;
                    }

                default: { return; }
            }

            Tabla.Columns.Remove("PautaId1");
            Tabla.Columns.Remove("PautaId2");

            gv.DataSource = Tabla;
            gv.DataBind();
            gv.Visible = true;

        }

        protected void AgregarTotales()
        {
            try {

                int suma = 0;
                decimal suma2 = 0;
                decimal suma3 = 0;


                ASPxGridView aux = new ASPxGridView();

                string Titulo = string.Empty;

                Tabla.Rows[Tabla.Rows.Count - 1][0] = "TOTALES";

                aux.DataSource = Tabla;

                aux.DataBind();
               
                gv.DataSource = Tabla;

                gv.DataBind();

                for (int x = 5; x <= Tabla.Columns.Count - 1; x++)
                {

                    Titulo = aux.Columns[x].ToString();

                    if (Titulo.Substring(Titulo.Length - 1) == "1")
                    {
                        Titulo = Titulo.Substring(0, Titulo.Length - 1);

                        aux.Columns[Titulo + "1"].Caption = Titulo + " - Costo";

                    }

                    if (aux.Columns[x].Index > 6 && aux.Columns[x].ToString().Substring(aux.Columns[x].ToString().Length - 1) != "1" && aux.Columns[x].ToString().Substring(aux.Columns[x].ToString().Length - 7) != "Salidas")
                    {
                        Titulo = aux.Columns[x].ToString() + " - Salidas";

                        aux.Columns[aux.Columns[x].Index].Caption = Titulo;

                    }

                    aux.Columns[x].HeaderStyle.Wrap = DevExpress.Utils.DefaultBoolean.True;

                    gv = aux;

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
                            else
                            {
                                Tabla.Rows[Tabla.Rows.Count - 1][x] = Convert.ToInt32(suma3);
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

            int Fila = Convert.ToInt32(Tabla.Rows.Count - 2);

            decimal TotalSalidas = 0;
            decimal TotalSKUS = 0;
            decimal ParcialSalidas = 0;
            decimal ParcialSKUS = 0;

            for (int xSalidas = 7; xSalidas <= Tabla.Columns.Count - 1;xSalidas++ )
            {
                if (Tabla.Rows[Fila][xSalidas] != System.DBNull.Value)
                {
                    if (Tabla.Rows[Fila][0].ToString() == "TOTALES")
                    {
                        Titulo = Tabla.Columns[xSalidas].ToString();

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
                Titulo = Tabla.Columns[x].ToString();

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
                    Tabla.Rows[Tabla.Rows.Count - 1][x] = System.Math.Round(ParcialSalidas / TotalSalidas * 100, 2);
                }
                else
                {
                    Tabla.Rows[Tabla.Rows.Count - 1][x] = System.Math.Round(ParcialSKUS / TotalSKUS * 100, 2);
                }
            }
        }

        #endregion

        protected void deAnoMes_DateChanged1(object sender, EventArgs e)
        {
            gv.Visible = false;
        }

    }
}