using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;
using System.Xml;
using DevExpress.Web.ASPxGridView.Export;
using DevExpress.Web.ASPxGridView;
using DevExpress.XtraPrinting;
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
using PautasPublicidad.Web.Classes;

namespace PautasPublicidad.Web.Forms
{
    public partial class Certificado : System.Web.UI.Page
    {
        int ProxRecId = 0;

        string botonbuscarpauta   = string.Empty;
        string botonbuscarperiodo = string.Empty;

        public static List<CertificadoDetDTO> mycert  =new List<CertificadoDetDTO>();

        CertificadoCabDTO certificado;

        public List<CertificadoDetDTO> Lineas
        {
            get
            {
                if (Session["Certificado.Lineas" + Session.SessionID] != null && Session["Ordenado.Lineas" + Session.SessionID] is List<CertificadoDetDTO>)
                    return Session["Certificado.Lineas" + Session.SessionID] as List<CertificadoDetDTO>;
                else
                    return new List<CertificadoDetDTO>();
            }
            set
            {
                Session.Add("Certificado.Lineas" + Session.SessionID, value);
            }
        }

        private void SortDDL(ref DropDownList ddl)
        {
            ListItem[] items = new ListItem[ddl.Items.Count];
            ddl.Items.CopyTo(items, 0);
            ddl.Items.Clear();
            Array.Sort(items, (x, y) => { return x.Text.CompareTo(y.Text); });
            ddl.Items.AddRange(items);
        } 

        private void RefreshHomeGrid(ASPxGridView gvHome)
        {
            //cargar con pautas 
            gvHome.DataSource = mycert;
            gvHome.DataBind();
        }
        
        void IdentifEspacio_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucEspacioChanged();
        }

        void ddlNroPauta_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucNumPautaChanged();
        }
        
        private EspacioContDTO GetEspacioContenidoEnPauta()
        {
            if (ddlNroPauta.SelectedValue != null)
                return CRUDHelper.Read(string.Format("RecId = '{0}'", ddlNroPauta.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
            else
                return null;
        }
        
        private void ucNumPautaChanged()
        {
            CertificadoCabDTO ord = Certificados.Read(Convert.ToInt32(ddlNroPauta.SelectedValue));
            if (ord != null) //funcion para completar Medio
            {
                ucIdentifEspacio.SelectedValue = ord.IdentifEspacio;

                DateTime d     = new DateTime( Convert.ToInt32(ord.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(ord.AnoMes.ToString().Substring(4, 2)), 01);
                deAnoMes.Value = d;
                ucEspacioChanged();
            }
            else
            {
                ucIdentifEspacio.SelectedValue = null;
                deAnoMes.Value = new DateTime();
            }
        }
        
        protected void btnRefresh_Click(object sender, ImageClickEventArgs e)
        {
            lblErrorLineas.Text = string.Empty;
            Labelx.Text         = lblErrorLineas.Text;
            string setup        = Certificados.AnoMesCierreOrd();
            decimal anomes      = Convert.ToDecimal(deAnoMes.Date.ToString("yyyyMM"));

            if (ucIdentifEspacio.SelectedValue != null)
                if (deAnoMes.Value != null)
                    if (anomes < Convert.ToDecimal(setup))
                        CargarCertificado();
                    else
                        lblMsg.Text = "No puede crear Certificado en el periodo seleccionado. Ya se encuentra cerrado.";
                else
                    lblMsg.Text = "Debe ingresar una fecha";
            else
                lblMsg.Text = "Debe seleccionar un Espacio de contenido para continuar";
                        
        }

        private void CargarPautas()
        {
            RefreshHomeGrid(gvHome); 
        }

        private bool ValidarSalida()
        {
            bool bresul = false;

            if (teHoraInicio.Text == "00:00" && teHoraFin.Text == "00:00")
                if( spSalidasInsertar.Text == "0")
                    lblErrorLineas.Text = "Debe ingresar un numero de salida mayor que cero";
            else
                bresul = true;

            return bresul;

        }

        private bool ValidarIntervalo(ref string mjeErr)
        {
            bool bresul = false;

            try
            {
                if (teHoraFin.DateTime != teHoraInicio.DateTime && ucIdentifIntervalo.SelectedValue == null)
                        mjeErr = ("Debe seleccionar un Intervalo");
                else
                    bresul = true;

                if (!bresul)
                    ASPxPageControl1.Visible = true;
                    trPauta.Visible          = true;
                    trFind.Visible           = false;

            }
            catch (Exception ex)
            {
                ASPxPageControl1.Visible = true;
                trPauta.Visible          = true;
                trFind.Visible           = false;

                MsgErrorLinas(ex);
                mjeErr = ex.Message;
            }
            return bresul;
        }

        private bool ValidarFrecuencia(ref string mjeErr)
        {
            bool bresul = false;

          try
            {
                if (ucIdentifFrecuencia.SelectedValue == null)
                {
                    mjeErr = ("Debe seleccionar una Frecuencia.");
                }
                else
                {
                    bresul = true;
                }
                if (!bresul)
                {
                    ASPxPageControl1.Visible = true;
                    trPauta.Visible = true;
                    trFind.Visible = false;
                }
          
            }
            catch (Exception ex)
            {
                ASPxPageControl1.Visible = true;
                trPauta.Visible          = true;
                trFind.Visible           = false;
                MsgErrorLinas(ex);
                mjeErr = ex.Message;

            } return bresul;
        }
        
        private void VaciarCamposPauta()
        { 
            spPautaID.Value                   = "";
            ucIdentifFrecuencia.SelectedValue = null;
            teHoraInicio.Value                = "";
            teHoraFin.Value                   = "";
            ucIdentifIntervalo.SelectedValue  = null;
            spVersionCosto.Value              = "";
            txUsuCosto.Value                  = "";
            txUsuCierre.Value                 = "";
            deFecCosto.Value                  = "";
            deFecCierre.Value                 = "";
            spCantSalidas.Value               = "";
            ucIdentifOrigen1.SelectedValue    = null;
            ucIdentifOrigen2.SelectedValue    = null;
            deAnoMes.Value = null;
            
            litCambiarPauta.Text = "Presione aqui para seleccionar una pauta ";


        }

        private void VaciarCamposDetalle()
        {
            //insertar lineas
            teHoraInicioInsertar.Value          = "";
            teHoraFinInsertar.Value             = "";
            spSalidasInsertar.Value             = "";
            ucIdentifAviso.SelectedValue        = "";
            spDuracionInsertar.Value            = "";
            
            //copiar periodos
            deFechaDesdeOrigenCopiar.Value      = "";
            deFechaHastaOrigenCopiar.Value      = "";
            deFechaDesdeDestinoCopiar.Value     = "";
            ucIdentifAvisoEdit.SelectedValue    = "";
            spDuracionEdit.Value                = "";
            spAvisoModifiSalidas.Value          = spSalidasInsertar.Value;

            //reemplazar
            deFechaDesdeReemplazar.Value        = "";
            deHoraDesdeOrigenReemplazar.Value   = "";
            deFechaHastaReemplazar.Value        = "";
            deHoraHastaOrigenReemplazar.Value   = "";

            ucIdentifAvisoDestinoReemplazar.SelectedValue = "";
            ucIdentifAvisoOrigenReemplazar.SelectedValue = "";
        }

        private void RefreshAbmGridUnsorted(ASPxGridView gv)
        {
            gv.DataSource = mycert;
            gv.DataBind();
        }

        protected void gvHome_CustomCallback(object sender, DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs e)
        {
        }

          protected void btn_ConsultarGvHome(object sender, EventArgs e) 
        {
            if (gvHome.Selection.Count > 0)
            {
                int recId = Convert.ToInt32(gvHome.GetSelectedFieldValues(new string[] { "RecId" })[0]);
                
                if (recId != 0) //funcion para completar Medio
                {
                    ucIdentifEspacio.SelectedValue = Convert.ToString((gvHome.GetSelectedFieldValues(new string[] { "IdentifEspacio" })[0]));

                    deAnoMes.Date = new DateTime(Convert.ToInt32(gvHome.GetSelectedFieldValues(new string[] { "año" })[0]), Convert.ToInt32(gvHome.GetSelectedFieldValues(new string[] { "mes" })[0]), 1);
                }
                else
                {
                    ucIdentifEspacio.SelectedValue = null;

                    //txMedio.Text = "";

                    deAnoMes.Value = new DateTime();
                }

                CargarCertificado();

                gv.SortBy(gv.Columns["Dia"], DevExpress.Data.ColumnSortOrder.Ascending);

                gv.SortBy(gv.Columns["Hora"], -1);

                gv.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);

            }
            else
            {
                lblErrorHome.Text = "Debe Seleccionar un registro para proceder";
            }
        }

        protected void btn_CancelDelete(object sender, EventArgs e)
        {
            tblDelete.Visible    = false;
            lblErrorHome.Visible = false;

            gvHome.Selection.UnselectAll();
        }

        protected void btn_ShowDelete(object sender, EventArgs e)
        {
            if (gvHome.Selection.Count > 0)
            {
                tblDelete.Visible   = true;
                lblErrorLineas.Text = string.Empty;
                lblErrorHome.Text   = string.Empty;
                lblMsg.Text         = string.Empty;
            }
            else
            {
                lblErrorHome.Text = "Debe seleccionar un certificado de la lista";
            }
        }
        
        protected void btn_EliminarLineaGvHome(object sender, EventArgs e)
        {
            if (gvHome.Selection.Count == 1)
            {
                int recId        = Convert.ToInt32(gvHome.GetSelectedFieldValues(new string[] { "RecId" })[0]);
                string UsuCierre = Convert.ToString(gvHome.GetSelectedFieldValues(new string[] { "CertValido" })[0]);

                if (recId > 0)
                {
                    if (UsuCierre == "")
                    {
                        try
                        {
                            Certificados.Delete(Certificados.Read(recId));

                            gvHome.DataSource = Certificados.VistaCertificados();

                            gvHome.DataBind();

                            tblDelete.Visible = false;

                            lblErrorHome.Text = "Certificado borrado correctamente.";
                        }
                        catch (Exception ex)
                        {
                            MsgErrorLinas(ex);
                        }

                    }
                    else
                    {
                        lblErrorHome.Text = "No puede borrar la pauta seleccionada, ya se encuentra cerrada.";
                    }
                }

                gvHome.Selection.UnselectAll();
            }
            else
            {
                lblErrorHome.Text = "Debe Seleccionar un registro para proceder.";
            }
        }

        protected void txMedio_ValueChanged(object sender, EventArgs e)
        {
            CertificadoCabDTO ord = Certificados.Read(Convert.ToInt32(ddlNroPauta.SelectedValue));

            if (ord != null) //funcion para completar Medio
            {
                ucIdentifEspacio.SelectedValue = ord.IdentifEspacio;
                
                DateTime d = new DateTime(Convert.ToInt32(ord.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(ord.AnoMes.ToString().Substring(4, 2)), 01);

                deAnoMes.Value = d;

                ucEspacioChanged();
            }
        }

        protected void deFecCierre_DateChanged(object sender, EventArgs e)
        {

        }

        protected void ASPxCallback1_Callback(object sender, DevExpress.Web.ASPxCallback.CallbackEventArgs e)
        {
        }

        public CostosDTO Costos
        {
            get
            {
                if (ViewState["Costos"] != null && ViewState["Costos"] is CostosDTO)
                    return ViewState["Costos"] as CostosDTO;
                else
                    return new CostosDTO();
            }
            set
            {
                ViewState.Add("Costos", value);
            }
        }

        public CertificadoCabDTO CertificadoCab
        {
            get
            {
                if (ViewState["CertificadoCab"] != null && ViewState["CertificadoCab"] is CertificadoCabDTO)
                    return ViewState["CertificadoCab"] as CertificadoCabDTO;
                else
                    return new CertificadoCabDTO();
            }
            set
            {
                ViewState.Add("CertificadoCab", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (this.Request.Params.Get("__EVENTTARGET") == "nula")
            {
                string js = "alert('El certificado no existe.')";

                ClientScript.RegisterStartupScript(GetType(), "Message", js, true);
            }

            ucIdentifEspacio.Inicializar(BusinessMapper.eEntities.EspacioCont);
            ucIdentifOrigen1.Inicializar(BusinessMapper.eEntities.Origen);
            ucIdentifOrigen2.Inicializar(BusinessMapper.eEntities.Origen);

            btnSave.Enabled                      = true;
            btnValidar.Enabled                   = true;
            ucIdentifFrecuencia.ComboBox.Enabled = true;

            ASPxGridViewExporter1.GridViewID = "gv";

            if (!Page.IsPostBack && !Page.IsCallback)
            {
                Lineas = null; //Limpio las lineas de mi session.

                mycert.Clear();

                spSalidasInsertar.Enabled   = true;
                trButtons.Visible           = true;
                trFind.Visible              = true;
                trPauta.Visible             = false;
                trAccion.Visible            = false;
                trEditLine.Visible          = false;
                trQuerySKU.Visible          = false;
                opEditPeriodo.Checked       = true;
                tblDelete.Visible           = false;

                FormsHelper.InicializarPropsGrilla(gv);

                gv.SettingsEditing.Mode                      = GridViewEditingMode.Inline;
                gv.SettingsBehavior.AllowSelectByRowClick    = true;
                gv.SettingsBehavior.AllowSelectSingleRowOnly = false;

                FormsHelper.InicializarPropsGrilla(gvHome);

                gvHome.SettingsEditing.Mode                      = GridViewEditingMode.Inline;
                gvHome.SettingsBehavior.AllowSelectByRowClick    = true;
                gvHome.SettingsBehavior.AllowSelectSingleRowOnly = false;

                ASPxGridViewExporter1.GridViewID = "gvHome";
                
                var ord = Certificados.ReadAll("PautaId IS NOT NULL");

                //Invisibilizo los controles btnRefresh y btnAdd de los combos
                ucIdentifEspacio.Controls[0].Controls[0].Controls[3].Visible    = false;
                ucIdentifEspacio.Controls[0].Controls[0].Controls[5].Visible    = false;
                ucIdentifOrigen1.Controls[0].Controls[0].Controls[3].Visible    = false;
                ucIdentifOrigen1.Controls[0].Controls[0].Controls[5].Visible    = false;
                ucIdentifOrigen2.Controls[0].Controls[0].Controls[3].Visible    = false;
                ucIdentifOrigen2.Controls[0].Controls[0].Controls[5].Visible    = false;
                ucIdentifFrecuencia.Controls[0].Controls[0].Controls[3].Visible = false;
                ucIdentifFrecuencia.Controls[0].Controls[0].Controls[5].Visible = false;
                ucIdentifIntervalo.Controls[0].Controls[0].Controls[3].Visible  = false;
                ucIdentifIntervalo.Controls[0].Controls[0].Controls[5].Visible  = false;

            }

            GridViewDataComboBoxColumn gvc = gv.Columns["IdentifAviso"] as GridViewDataComboBoxColumn;

            gvc.Name                          = "IdentifAviso";
            gvc.Caption                       = "Aviso";
            gvc.FieldName                     = "IdentifAviso";
            gvc.PropertiesComboBox.TextField  = "Name";                         //mapInfo.EntityTextField;
            gvc.PropertiesComboBox.ValueField = "IdentifAviso";                 // mapInfo.EntityValueField;
            gvc.PropertiesComboBox.DataSource = Business.Avisos.ReadAll("");    //mapInfo.DAOHandler.ReadAll("");

            lblErrorLineas.Text = string.Empty;

            ucIdentifEspacio.Inicializar(BusinessMapper.eEntities.EspacioCont);
            ucIdentifFrecuencia.Inicializar(BusinessMapper.eEntities.Frecuencia);
            ucIdentifIntervalo.Inicializar(BusinessMapper.eEntities.Intervalo);

            if (ucIdentifEspacio.SelectedValue != null)
            {
                ucIdentifAviso.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue));
                ucIdentifAviso.WhereFilter = string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue);
            }
            else
            {
                ucIdentifAviso.Inicializar(BusinessMapper.eEntities.Avisos);
            }

            if (ucIdentifEspacio.SelectedValue != null)
            {
                ucIdentifAvisoOrigenReemplazar.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue)); ucIdentifAvisoOrigenReemplazar.WhereFilter = string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue);
                ucIdentifAvisoDestinoReemplazar.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue)); ucIdentifAvisoDestinoReemplazar.WhereFilter = string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue);
                ucIdentifAvisoEdit.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue)); ucIdentifAvisoEdit.WhereFilter = string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue);
            }
            else
            {
                ucIdentifAvisoOrigenReemplazar.Inicializar(BusinessMapper.eEntities.Avisos);
                ucIdentifAvisoDestinoReemplazar.Inicializar(BusinessMapper.eEntities.Avisos);
                ucIdentifAvisoEdit.Inicializar(BusinessMapper.eEntities.Avisos);
            }

            ucIdentifEspacio.ComboBox.AutoPostBack = true;
            ucIdentifEspacio.ComboBox.SelectedIndexChanged += new EventHandler(IdentifEspacio_SelectedIndexChanged);

            ASPxCallback1.Callback += new DevExpress.Web.ASPxCallback.CallbackEventHandler(ASPxCallback1_Callback);

            ddlNroPauta.AutoPostBack = true;
            ddlNroPauta.SelectedIndexChanged += new EventHandler(ddlNroPauta_SelectedIndexChanged);

            ucIdentifFrecuencia.ComboBox.AutoPostBack = true;
            ucIdentifFrecuencia.ComboBox.SelectedIndexChanged += new EventHandler(IdentifFrecuencia_SelectedIndexChanged);

            ucIdentifAviso.ComboBox.AutoPostBack = true;
            ucIdentifAviso.ComboBox.SelectedIndexChanged += new EventHandler(IdentifAviso_SelectedIndexChanged);

            ucIdentifAvisoDestinoReemplazar.ComboBox.AutoPostBack = true;
            ucIdentifAvisoDestinoReemplazar.ComboBox.SelectedIndexChanged += new EventHandler(AvisoDestinoReemplazar_SelectedIndexChanged);

            ucIdentifAvisoEdit.ComboBox.AutoPostBack = true;
            ucIdentifAvisoEdit.ComboBox.SelectedIndexChanged += new EventHandler(IdentifAvisoEdit_SelectedIndexChanged);

            gv.KeyFieldName = "RecId";

            RefreshAbmGrid(gv);

            if (trQuerySKU.Visible) { RefreshSKUGrid(gvSKU); }

            gvHome.DataSource = Certificados.VistaCertificados();
            gvHome.DataBind();

            spEspacio.Value = ucIdentifEspacio.SelectedValue;

            if (gvHome.VisibleRowCount > 0)
            {
                spAnioMes.Value = gvHome.GetCurrentPageRowValues("año")[0].ToString() + "-" + gvHome.GetCurrentPageRowValues("mes")[0].ToString(); ;
                spOrigenCertificado.Value = gvHome.GetCurrentPageRowValues("IdentifOrigen")[0].ToString(); 
            }
        }

        void AvisoDestinoReemplazar_SelectedIndexChanged(object sender, EventArgs e)
        {
            AvisoDestinoReemplazarChanged();
        }

        private void AvisoDestinoReemplazarChanged()
        {
            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoDestinoReemplazar.SelectedValue));

            if (aviso != null)
            {
                spAvisoReempDuracion.Value = aviso.Duracion;
                string msg                 = string.Empty;

                if (!ValidarFranjaHoraria(ref msg))
                    lblErrorLineas.Text = msg;

            }
        }

        private void RefreshSKUGrid(ASPxGridView gvSKU)
        {
            decimal total = 0;
            var dt = Certificados.BuildAllSKU(mycert);

            gvSKU.DataSource = dt;
            gvSKU.DataBind();

            foreach (System.Data.DataRow dr in dt.Rows)
                total += Convert.ToDecimal(dr["CantSalidas"]);

            lblSKUTotalSalidas.Text = "Total de Salidas: " + total.ToString();
        }

        void IdentifAvisoEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoEdit.SelectedValue));

            if (aviso != null)
                spDuracionEdit.Value = aviso.Duracion;

            spAvisoModifiSalidas.Value = gv.GetSelectedFieldValues("Salida")[0].ToString();
        }

        void IdentifAviso_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucAvisoChanged();
        }

        void IdentifFrecuencia_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucFrecuenciaChanged();
        }

        protected void gv_CustomColumnDisplayText(object sender, ASPxGridViewColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName != "Hora") return;

            if (e.Value != null)
                e.DisplayText = e.Value.ToString().Substring(0, 5);
        }

        protected void gvHome_CustomColumnDisplayText(object sender, ASPxGridViewColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName != "FecCertValido") return;

            if (e.Value != null)
                if (e.Value.ToString() == "")
                    e.DisplayText = "";
                else
                   if(e.Value.ToString().Substring(0,10) == "01/01/1900" || e.Value.ToString() == "") 
                     e.DisplayText = "";
        }
        private void RefreshAbmGrid(ASPxGridView gv)
        {
            var lineas = new List<CertificadoDetDTO>();

            if (mycert != null)
            {
                mycert = mycert.OrderBy(t => t.Dia).ThenBy(u => u.Hora).ThenBy(v => v.Salida).ToList();
            }

            lineas = mycert;

            lineas = lineas.OrderBy(t => t.Dia).ThenBy(u => u.Hora).ThenBy(v => v.Salida).ToList();

            gv.DataSource = lineas;
            gv.DataBind();
        }

        private void ReCargarControles(CertificadoCabDTO certificado)
        {
            EspacioContDTO espacio = GetEspacioContenido();

            lblErrorLineas.Text = string.Empty;
            Labelx.Text = string.Empty;

            //Controles de la pauta.
            spPautaID.Number                  = Convert.ToInt32(certificado.PautaId);
            ucIdentifFrecuencia.SelectedValue = certificado.IdentifFrecuencia;
            ucIdentifIntervalo.SelectedValue  = certificado.IdentifIntervalo;

            if (FormsHelper.ConvertToDateTime(certificado.HoraInicio) != Convert.ToDateTime("00:00:00"))
                teHoraInicio.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraInicio);

            teHoraInicioInsertar.DateTime = teHoraInicio.DateTime;

            if (FormsHelper.ConvertToDateTime(certificado.HoraFin) != Convert.ToDateTime("00:00:00"))
                teHoraFin.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraFin);

            teHoraFinInsertar.DateTime = teHoraFin.DateTime;

            //Controles de solo lectura.
            spVersionCosto.Value = certificado.VersionCosto;
            txUsuCosto.Text      = certificado.UsuCosto;
            deFecCosto.Date      = certificado.FecCosto;
            deFecCosto.Value     = certificado.FecCosto;
            txUsuCierre.Text     = certificado.CertValido;

            if (certificado.FecCertValido.ToShortDateString() == "01/01/1900" || certificado.FecCertValido == DateTime.MinValue)
                deFecCierre.Value = "";
            else
                deFecCierre.Value = certificado.FecCertValido;

            spCantSalidas.Value = certificado.CantSalidas;

            string Espacio   = string.Empty;
            string DeAnioMes = string.Empty;

            if (botonbuscarpauta == "true")
            {
                if (gvHome.Selection.Count > 0)
                {
                    Espacio   = gvHome.GetSelectedFieldValues("Espacio")[0].ToString();
                    DeAnioMes = gvHome.GetSelectedFieldValues("Año")[0].ToString() + "-" + gvHome.GetSelectedFieldValues("Mes")[0].ToString().PadLeft(2, '0');
                }
                else
                {
                    Espacio   = certificado.IdentifEspacio;
                    DeAnioMes = certificado.VigDesde.Date.Year.ToString() + "-" + certificado.VigDesde.Date.Month.ToString().PadLeft(2, '0');
                }

            }
            else
            {
                if (botonbuscarperiodo == "true")
                {
                    Espacio   = ucIdentifEspacio.SelectedText;
                    DeAnioMes = deAnoMes.Date.ToString("yyyy-MM");
                }
                else
                {
                    //tocó la grilla
                    Espacio = gvHome.GetCurrentPageRowValues("IdentifEspacio")[0].ToString();

                    switch (botonbuscarpauta)
                    {
                        case "true": break;
                        case "": 
                            {
                                DeAnioMes = certificado.VigDesde.Date.Year.ToString() + "-" + certificado.VigDesde.Date.Month.ToString().PadLeft(2, '0');
                                break; 
                            }
                        case "false":
                            {
                                switch (botonbuscarperiodo)
                                {
                                    case "true": { DeAnioMes = certificado.VigDesde.Date.Year.ToString() + "-" + certificado.VigDesde.Date.Month.ToString().PadLeft(2, '0'); break; }
                                    case "": { break; }
                                    case "false": { DeAnioMes = gvHome.GetSelectedFieldValues("año")[0].ToString() + "-" + gvHome.GetSelectedFieldValues("mes")[0].ToString().PadLeft(2, '0'); break; }
                                }
                                break;
                            }
                    }
                }

            }

            litCambiarPauta.Text = string.Format("Espacio: {0} | Período: {1} | Origen: {2}", Espacio, DeAnioMes, certificado.IdentifOrigen);

            //Actualizo controles.
            ucFrecuenciaChanged();

            ASPxPageControl1.Visible = true;
            trPauta.Visible          = true;
            trFind.Visible           = false;

            //Solo se pueden modificar certificados NO cerrados.
            if (certificado.FecCertValido.Year == 1 || certificado.FecCertValido.Year == 1900)
            {
                btnSave.Enabled    = true;
                btnValidar.Enabled = true;
                ucIdentifFrecuencia.ComboBox.Enabled = true;
            }
            else
            {
                btnSave.Enabled    = false;
                btnValidar.Enabled = false;
                ucIdentifFrecuencia.ComboBox.Enabled = false;
            }

            //Inicializo Fechas...

            deHoraDesdeOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraInicio);
            deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraFin);

            DateTime date = new DateTime();
            date = Convert.ToDateTime(deAnoMes.Date);

            deFechaDesdeDestinoCopiar.Date = new DateTime(date.Year, date.Month, 1);
            deFechaDesdeOrigenCopiar.Date  = new DateTime(date.Year, date.Month, 1);
            deFechaHastaOrigenCopiar.Date  = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
            deFechaDesdeReemplazar.Date    = new DateTime(date.Year, date.Month, 1);
            deFechaHastaReemplazar.Date    = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
        }        
        
        private void ucEspacioChanged()
        {
            EspacioContDTO espacio = GetEspacioContenido();

            if (espacio != null)
            {
                ucIdentifFrecuencia.SelectedValue = espacio.IdentifFrecuencia;
                if (espacio.HoraInicio.HasValue && espacio.HoraFin.HasValue)
                {
                    teHoraInicio.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraInicio.Value);
                    teHoraFin.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraFin.Value);
                }
                else
                {
                    teHoraInicio.Value = null;
                    teHoraFin.Value = null;
                }
                ucIdentifIntervalo.SelectedValue = espacio.IdentifIntervalo;
            }
        }

        private void ucFrecuenciaChanged()
        {
            FrecuenciaDTO frecuencia = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));

            if (frecuencia != null)
            {
                List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                ceDiasInsertar.Items.Clear();

                foreach (FrecuenciaDetDTO frecuenciaDetalle in frecuenciaDetalles)
                {
                    if (frecuencia.SemMes == "MES")
                    {
                        ceDiasInsertar.Items.Add(frecuenciaDetalle.Dia.Value.ToString(), frecuenciaDetalle.Dia.Value.ToString());
                    }
                    else
                    {
                        ceDiasInsertar.Items.Add(frecuenciaDetalle.DiaSemana, frecuenciaDetalle.DiaSemana.ToUpper().Trim());
                    }
             
                    ceDiasInsertar.Items[ceDiasInsertar.Items.Count - 1].Selected = true;
                }
            }
        }

        private void ucAvisoChanged()
        {
            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAviso.SelectedValue));

            if (aviso != null)
            {
                spDuracionInsertar.Value = aviso.Duracion;
                string msg               = string.Empty;

                if (!ValidarFranjaHoraria(ref msg))
                    lblErrorLineas.Text = msg;
            }
        }

        private EspacioContDTO GetEspacioContenido()
        {
            if (ucIdentifEspacio.SelectedValue != null)
            {
                return CRUDHelper.Read(string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

            }
            else

                if (botonbuscarpauta == "true")
                {
                    return CRUDHelper.Read(string.Format("IdentifEspacio = '{0}'", certificado.IdentifEspacio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
                }
                else
                {
                    return null;
                }
        }

        public int NextTempRecId()
        {
            if (ViewState["TempRecId"] != null)
            {
                int RecId = Convert.ToInt32(ViewState["TempRecId"]) + 1;
                ViewState.Add("TempRecId", RecId);
                return RecId;
            }
            else
            {
                ViewState.Add("TempRecId", 1);
                return 1;
            }
        }

        protected void btnGenerarLineas_Click(object sender, EventArgs e)
        {
            try
            {

                string msg = string.Empty;

                lblErrorLineas.Text = string.Empty;

                if (!ValidarFranjaHoraria(ref msg))
                {
                    lblErrorLineas.Visible = true;
                    lblErrorLineas.Text += " " + msg + "\r\n";
                }

                if (!ValidarFrecuencia(ref msg))
                {
                    lblErrorLineas.Visible = true;
                    lblErrorLineas.Text += " " + msg + "\r\n";
                }

                if (ceDiasInsertar.SelectedItems.Count == 0)
                {
                    lblErrorLineas.Visible = true;
                    lblErrorLineas.Text += " No se han seleccionado días de la semana." + "\r\n";
                }

                if (ucIdentifAviso.ComboBox.Text == "")
                {
                    lblErrorLineas.Visible = true;
                    lblErrorLineas.Text += " No se ha seleccionado ningún aviso." + "\r\n";
                }

                if (lblErrorLineas.Text == "")
                {
                    GenerarLineas(teHoraInicioInsertar.DateTime, teHoraFinInsertar.DateTime);

                    if (lblErrorLineas.Text == "")
                    { lblErrorLineas.Text = "Lineas insertadas correctamente. No olvide GRABAR antes de continuar."; }

                    RefreshAbmGrid(gv);

                    trAccion.Visible = false;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void GenerarLineas(DateTime horaInicio, DateTime horaFin)
        {
            try
            {
                IntervaloDTO intervalo                    = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));
                FrecuenciaDTO frecuencia                  = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAviso.SelectedValue));

                //Genero las nuevas líneas.
                mycert = GenerarLineas(FormsHelper.ConvertToTimeSpan(horaInicio), FormsHelper.ConvertToTimeSpan(horaFin), intervalo, frecuencia, aviso, frecuenciaDetalles);
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }

        private List<string> GetDiasSeleccionados()
        {
            List<string> diasSeleccionados = new List<string>();

            foreach (var item in ceDiasInsertar.SelectedValues)
                diasSeleccionados.Add(Convert.ToString(item));

            return diasSeleccionados;
        }

        private List<CertificadoDetDTO> GenerarLineas(TimeSpan horaInicio, TimeSpan horaFin, IntervaloDTO intervalo, FrecuenciaDTO frecuenciaCab, AvisosDTO aviso, List<FrecuenciaDetDTO> frecuenciaDetalles)
        {
            List<CertificadoDetDTO> lineas        = mycert;
            List<CertificadoDetDTO> preExistentes = new List<CertificadoDetDTO>();

            if (intervalo == null)
            {
                intervalo             = new IntervaloDTO();
                intervalo.CantMinutos = Convert.ToDecimal(0);
            }

            if (Validaciones(horaInicio, horaFin, intervalo, frecuenciaCab, aviso, frecuenciaDetalles) == false)
            {
                RefreshAbmGrid(gv);
                return lineas;
            }

            TimeSpan incremento = TimeSpan.FromMinutes(Convert.ToDouble(intervalo.CantMinutos));
            TimeSpan horaTemp;
            CertificadoDetDTO linea;
            List<DateTime> periodo;

            try
            {
                //Obtengo la lista de los días partiedo de fechas o nombres de dia.
                if (frecuenciaCab.SemMes == "SEMANA")
                    periodo = Certificados.GetDatesByDayNames(deAnoMes.Date.Year, deAnoMes.Date.Month, GetDiasSeleccionados());
                else
                    periodo = Certificados.GetDatesByDayNumbers(deAnoMes.Date.Year, deAnoMes.Date.Month, GetDiasSeleccionados());

                foreach (DateTime fecha in periodo)
                {
                    horaTemp = horaInicio;

                    if (aviso.Duracion == 0)
                    {
                        incremento = new TimeSpan(0);
                    }
                    else if (aviso.Duracion == null)
                    {
                        incremento = new TimeSpan(0);
                    }

                    if (horaTemp.CompareTo(horaFin) == 0)
                    {
                        horaFin = horaFin.Add(TimeSpan.FromHours(1));
                        incremento = incremento.Add(TimeSpan.FromHours(1));
                    }

                    if (incremento.Minutes == 0 && ((horaTemp.CompareTo(horaFin) == 0)))
                    {
                        horaFin = horaFin.Add(TimeSpan.FromHours(1));
                        incremento = incremento.Add(TimeSpan.FromHours(1));
                    }
                    else if (incremento.Minutes == 0 && ((horaTemp.CompareTo(horaFin) == 1)))
                    {
                        horaFin = horaFin.Add(TimeSpan.FromHours(-1));
                    }
                    
                    //Mientras no supere la hora hasta...
                    while (horaTemp.CompareTo(horaFin) < 0)
                    {
                        DateTime fechaTmp = new DateTime(fecha.Year, fecha.Month, fecha.Day, horaTemp.Hours, horaTemp.Minutes, horaTemp.Seconds);

                        linea           = new CertificadoDetDTO();
                        linea.RecId     = lineas.Count;
                        linea.Fecha     = fechaTmp;
                        linea.Hora      = horaTemp;
                        linea.Dia       = fecha.Day;
                        linea.DiaSemana = fecha.ToString("dddd", new CultureInfo("es-ES")).ToUpper().Trim();

                        if (aviso != null)
                        {
                            linea.IdentifAviso = aviso.IdentifAviso;
                            linea.Duracion     = aviso.Duracion;
                        }
                        else
                        {
                            linea.IdentifAviso = string.Empty;
                            linea.Duracion     = null;
                        }

                        linea.Salida = spSalidasInsertar.Number;

                        if (!lineas.Exists(
                            (x) => (x.Fecha == fechaTmp && x.Hora == horaTemp && x.Salida == linea.Salida && x.IdentifAviso == linea.IdentifAviso)))
                        {
                            lineas.Add(linea);
                        }
                        else
                        {
                            preExistentes.Add(linea);
                        }

                        horaTemp = horaTemp.Add(incremento);

                        if (aviso.Duracion == null || aviso.Duracion == 0)
                        {
                            horaTemp = horaFin;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }

            if (preExistentes.Count > 0)
            {
                lblErrorLineas.Text = "No se pudieron grabar todas las lineas. No olvide GRABAR antes de continuar.";
            }

            RefreshAbmGrid(gv);

            return lineas;
        }

        protected void ASPxPageControl2_ActiveTabChanged(object source, DevExpress.Web.ASPxTabControl.TabControlEventArgs e)
        {
            //RecargarDiaHora();

            string msg = string.Empty;

            if (!ValidarFranjaHoraria(ref msg))
                lblErrorLineas.Text = msg;

            if (e.Tab.Text == "Copiar Períodos")
            {
                CopyItems();
            }
            if (e.Tab.Text == "Reemplazar Avisos")
            {
                ReplaceItems();
            }
            if (e.Tab.Text == "Insertar Líneas")
            {
                InsertItems();
            }
        }

        private bool ValidarFranjaHoraria(ref string mjeErr)
        {
            bool retVal = false;

            try
            {
                if (teHoraInicio.Text == teHoraFin.Text && teHoraInicio.Text == "")
                    return true;

                if (Convert.ToDateTime(teHoraInicio.Text) > Convert.ToDateTime(teHoraFin.Text))
                {
                    mjeErr = ("La hora de inicio no puede ser mayor a la final.");
                }
                else
                {
                    retVal = true;
                }

                if (!retVal)
                {
                    ASPxPageControl1.Visible = true;
                    trPauta.Visible = true;
                    trFind.Visible = false;
                }
            }
            catch (Exception ex)
            {
                ASPxPageControl1.Visible = true;
                trPauta.Visible = true;
                trFind.Visible = false;

                MsgErrorLinas(ex);
                mjeErr = ex.Message;
            }

            return retVal;
        }

        private void CargarCertificado()
        {
            lblMsg.Text = string.Empty;
            lblErrorHome.Text = string.Empty;

            string sIdentifOrigen = string.Empty;

            if (botonbuscarperiodo != "")
            {
                if (botonbuscarperiodo == "true")
                {
                    sIdentifOrigen = (string)ucIdentifOrigen1.SelectedValue;
                    certificado = Business.Certificados.Read((string)ucIdentifEspacio.SelectedValue, deAnoMes.Date.Year, deAnoMes.Date.Month, sIdentifOrigen);
                }
                else
                {
                    sIdentifOrigen = gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString();
                    certificado = Business.Certificados.Read((string)ucIdentifEspacio.SelectedValue, deAnoMes.Date.Year, deAnoMes.Date.Month, sIdentifOrigen);
                }
            }
            else
            {
                if (botonbuscarpauta == "")
                {
                    if (botonbuscarpauta == "true")
                    {
                        sIdentifOrigen = gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString();
                        certificado    = Business.Certificados.Read(txNroPauta.Text, sIdentifOrigen);
                    }
                    else
                    {
                        sIdentifOrigen = gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString();
                        var miPauta    = gvHome.GetSelectedFieldValues("PautaId")[0].ToString();
                        certificado    = Business.Certificados.Read(miPauta, sIdentifOrigen);
                    }
                }
                else
                {
                    sIdentifOrigen = (string)ucIdentifOrigen2.SelectedValue;
                    certificado    = Business.Certificados.Read(txNroPauta.Text, sIdentifOrigen);
                }

            }

            ASPxPageControl1.Visible = true;
            trPauta.Visible = true;
            trFind.Visible = false;
            EspacioContDTO espacio = GetEspacioContenido();

            if (botonbuscarpauta == "false" || botonbuscarpauta == "")
            {
                if (botonbuscarperiodo == "false" || botonbuscarperiodo == "")
                    spOrigenCertificado.Value = gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString(); 
                else
                    spOrigenCertificado.Value = certificado.IdentifOrigen;
            }
            else
            {
                spOrigenCertificado.Value = certificado.IdentifOrigen;
            }

            if (certificado != null)
                ReCargarControles(certificado);
            else
            {
                if (gvHome.Selection.Count > 0)
                {
                    string EspacioX = string.Empty;
                    string DeAnioMes = string.Empty;
                    string IDOrigen = string.Empty;

                    EspacioX  = gvHome.GetSelectedFieldValues("IdentifEspacio")[0].ToString();
                    DeAnioMes = gvHome.GetSelectedFieldValues("año")[0].ToString() + "-" + gvHome.GetSelectedFieldValues("mes")[0].ToString().PadLeft(2, '0');
                    IDOrigen  = gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString();

                    litCambiarPauta.Text = string.Format("Espacio: {0} | Período: {1} | Origen: {2}", EspacioX, DeAnioMes, IDOrigen);

                    EspacioContDTO espacio0 = GetEspacioContenido();

                    lblErrorLineas.Text = string.Empty;
                    Labelx.Text = string.Empty;

                    //Controles de la pauta.
                    spPautaID.Number = Convert.ToDecimal(gvHome.GetSelectedFieldValues("PautaId")[0]);

                    certificado = Business.Certificados.Read(spPautaID.Number.ToString(), IDOrigen);

                    ucIdentifFrecuencia.SelectedValue = certificado.IdentifFrecuencia;
                    ucIdentifIntervalo.SelectedValue  = certificado.IdentifIntervalo;

                    if (FormsHelper.ConvertToDateTime(certificado.HoraInicio) != Convert.ToDateTime("00:00:00"))
                        teHoraInicio.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraInicio);

                    teHoraInicioInsertar.DateTime = teHoraInicio.DateTime;

                    if (FormsHelper.ConvertToDateTime(certificado.HoraFin) != Convert.ToDateTime("00:00:00"))
                        teHoraFin.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraFin);

                    teHoraFinInsertar.DateTime = teHoraFin.DateTime;

                    //Controles de solo lectura.
                    spVersionCosto.Value = certificado.VersionCosto;
                    txUsuCosto.Text      = certificado.UsuCosto;
                    deFecCosto.Date      = certificado.FecCosto;
                    deFecCosto.Value     = certificado.FecCosto;
                    txUsuCierre.Text     = certificado.CertValido;
                    deFecCierre.Value    = certificado.FecCertValido;
                    spCantSalidas.Value  = certificado.CantSalidas;

                    //Actualizo controles.
                    ucFrecuenciaChanged();

                    ASPxPageControl1.Visible = true;
                    trPauta.Visible          = true;
                    trFind.Visible           = false;

                    //Solo se pueden modificar certificados NO cerrados.
                    if (certificado.FecCertValido.Year == 1)
                    {
                        btnSave.Enabled     = true;
                        btnValidar.Enabled  = true;
                        ucIdentifFrecuencia.ComboBox.Enabled = true;
                    }
                    else
                    {
                        btnSave.Enabled     = false;
                        btnValidar.Enabled  = false;
                        ucIdentifFrecuencia.ComboBox.Enabled = false;

                    }

                    deHoraDesdeOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraInicio);
                    deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraFin);

                    DateTime date = new DateTime();
                    date = Convert.ToDateTime(deAnoMes.Date);

                    deFechaDesdeDestinoCopiar.Date  = new DateTime(date.Year, date.Month, 1);
                    deFechaDesdeOrigenCopiar.Date   = new DateTime(date.Year, date.Month, 1);
                    deFechaHastaOrigenCopiar.Date   = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
                    deFechaDesdeReemplazar.Date     = new DateTime(date.Year, date.Month, 1);
                    deFechaHastaReemplazar.Date     = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
                }

            }

            RecargarDiaHora();

            spEspacio.Value           = GetEspacioContenido().IdentifEspacio;
            spMedio.Value             = GetEspacioContenido().IdentifMedio;
            spAnioMes.Value           = CertificadoCab.AnoMes.ToString().Substring(0, 4) + "-" + CertificadoCab.AnoMes.ToString().Substring(4);
            spOrigenCertificado.Value = certificado.IdentifOrigen;

        }

        private void RecargarDiaHora()
        {

            lblMsg.Text       = string.Empty;
            lblErrorHome.Text = string.Empty;

            string sIdentifOrigen  = string.Empty;
            EspacioContDTO espacio = GetEspacioContenido();

            if (botonbuscarpauta == "false" || botonbuscarpauta == "")
            {
                if (botonbuscarperiodo == "false" || botonbuscarperiodo == "")
                {
                    //selecciono directamente de la grilla
                    sIdentifOrigen = gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString();
                    var miPauta    = gvHome.GetSelectedFieldValues("PautaId")[0].ToString()      ;
                    certificado    = Business.Certificados.Read(miPauta, sIdentifOrigen)         ;
                }
                else
                {
                    //botonbuscarperiodo = "true"
                    sIdentifOrigen = (string)ucIdentifOrigen1.SelectedValue;

                    certificado = Business.Certificados.Read((string)ucIdentifEspacio.SelectedValue, deAnoMes.Date.Year, deAnoMes.Date.Month, sIdentifOrigen);
                }
            }
            else
            {
                //botonbuscarpauta = "true"
                sIdentifOrigen = (string)ucIdentifOrigen2.SelectedValue;

                certificado = Business.Certificados.Read(txNroPauta.Text, sIdentifOrigen);
            }

            CostosDTO costos = Business.Certificados.FindCosto(Convert.ToString(ucIdentifEspacio.SelectedValue), deAnoMes.Date.Year, deAnoMes.Date.Month);

            if (certificado != null)
            {
                if (CertificadoCab.CertValido != null)
                {
                    certificado.CertValido = CertificadoCab.CertValido;
                    certificado.FecCertValido = CertificadoCab.FecCertValido;
                }

                CertificadoCab = certificado;
                mycert = Certificados.ReadAllLineas(certificado);

                Costos = costos;
                RefreshAbmGrid(gv);

                /// si el certificado esta cerrado oculto paneles de edicion. 
                if (CertificadoCab.CertValido != "" && CertificadoCab.CertValido != null)
                {
                    ASPxPageControl2.Visible = false;
                    mnuDetalle.Visible = false;
                }
                else
                {
                    ASPxPageControl2.Visible = true;
                    mnuDetalle.Visible = true;
                }

                DateTime date = new DateTime();
                date = Convert.ToDateTime(deAnoMes.Date);

                deFechaDesdeOrigenCopiar.Date = new DateTime(date.Year, date.Month, 1);
                deFechaHastaOrigenCopiar.Date = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
                deFechaDesdeReemplazar.Date   = new DateTime(date.Year, date.Month, 1);
                deFechaHastaReemplazar.Date   = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);

                if (espacio.HoraInicio != null)
                    deHoraDesdeOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraInicio.Value);
                else
                    deHoraDesdeOrigenReemplazar.DateTime = certificado.VigDesde.Date;

                if (espacio.HoraFin != null)
                    deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraFin.Value);
                else if (certificado.HoraFin != null)
                    deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraFin);
                else
                    deHoraHastaOrigenReemplazar.DateTime = certificado.VigHasta.Date.AddHours(23.9999);//origen

                deFechaDesdeDestinoCopiar.Date = deFechaDesdeOrigenCopiar.Date;

            }
            else
            {
                if (costos != null)
                {
                    Costos = costos;
                    //Cargo controles.
                    spPautaID.Number = Business.Certificados.GetNextPautaId();

                    if (ucIdentifFrecuencia.SelectedValue == null)
                        ucIdentifFrecuencia.SelectedValue = espacio.IdentifFrecuencia;

                    if (ucIdentifIntervalo.SelectedValue == null)
                        ucIdentifIntervalo.SelectedValue = espacio.IdentifIntervalo;

                    if (espacio.HoraInicio.HasValue)
                    {
                        teHoraInicio.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraInicio.Value);
                        teHoraInicioInsertar.DateTime = teHoraInicio.DateTime;
                    }

                    if (espacio.HoraFin.HasValue)
                    {
                        teHoraFin.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraFin.Value);
                        teHoraFinInsertar.DateTime = teHoraFin.DateTime;
                    }

                    //Actualizo controles.
                    ucFrecuenciaChanged();

                    ASPxPageControl1.Visible = true;
                    trPauta.Visible          = true;
                    trFind.Visible           = false;

                    //Inicializo Fechas origen para copia
                    deFechaDesdeOrigenCopiar.Date = costos.VigDesde;  //origen
                    deFechaHastaOrigenCopiar.Date = costos.VigHasta; //origen
                    deFechaDesdeReemplazar.Date   = costos.VigDesde;//origen
                    deFechaHastaReemplazar.Date   = costos.VigHasta;//origen

                    //Inicializo Fechas origen para copia
                    DateTime date = new DateTime();
                    date = Convert.ToDateTime(deAnoMes.Date);
                    deFechaDesdeOrigenCopiar.Date = new DateTime(date.Year, date.Month, 1);
                    deFechaHastaOrigenCopiar.Date = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
                    deFechaDesdeReemplazar.Date   = new DateTime(date.Year, date.Month, 1);
                    deFechaHastaReemplazar.Date   = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);

                    //inicializo horas origen
                    if (espacio.HoraInicio != null)
                        deHoraDesdeOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraInicio.Value);
                    else
                        deHoraDesdeOrigenReemplazar.DateTime = costos.VigDesde.Date;

                    if (espacio.HoraFin != null)
                        deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraFin.Value);
                    else if (certificado != null)
                        if (certificado.HoraFin != null)
                            deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(certificado.HoraFin);

                    //fecha inicio de destino
                    deFechaDesdeDestinoCopiar.Date = costos.VigDesde; //destino

                    certificado                     = new CertificadoCabDTO();
                    certificado.AnoMes              = Convert.ToInt32(deAnoMes.Date.Year.ToString() + deAnoMes.Date.Month.ToString("00"));
                    certificado.CantSalidas         = 0;                      //Cantidad de salidas total
                    certificado.Costo               = 0;                      //Costo Total de la Pauta
                    certificado.CostoOp             = 0;                      //Costo de la Pauta para la Orden de Publicidad
                    certificado.CostoOpUni          = 0;                      //Costo de la Pauta para la Orden de Publicidad por unidad (segundos, página)
                    certificado.CostoUni            = 0;                      //Costo Total por unidad (segundos, página)
                    certificado.DatareaId           = 0;
                    certificado.DuracionTot         = 0;                      //Total Duración o Cantidad
                    certificado.FecCosto            = DateTime.Now;           //Fecha en qué se calculo el costo por última vez
                    certificado.HoraInicio          = FormsHelper.ConvertToTimeSpan(teHoraInicio.DateTime);
                    certificado.HoraFin             = FormsHelper.ConvertToTimeSpan(teHoraFin.DateTime);
                    certificado.IdentifEspacio      = Convert.ToString(ucIdentifEspacio.SelectedValue);
                    certificado.IdentifFrecuencia   = Convert.ToString(ucIdentifFrecuencia.SelectedValue);
                    certificado.IdentifIntervalo    = Convert.ToString(ucIdentifIntervalo.SelectedValue);
                    certificado.PautaId             = Convert.ToString(spPautaID.Number);
                    certificado.RecId               = 0;
                    certificado.CertValido          = "";                     //Usuario que cerró
                    certificado.UsuCosto            = costos.Confirmado;      //Usuario qué calculo el costo
                    certificado.VersionCosto        = costos.Version.Value;   //Versión del registro
                    certificado.VigDesde            = costos.VigDesde;        //Fecha desde la cual está vigente el Costo
                    certificado.VigHasta            = costos.VigHasta;        //Fecha hasta la cual estará vigente el Costo

                    CertificadoCab = certificado;

                    RefreshHomeGrid(gvHome);

                    string msg = string.Empty;
                    
                    if (!ValidarFranjaHoraria(ref msg))
                        lblErrorLineas.Text = msg;

                    msg = string.Empty;

                    if (!ValidarFrecuencia(ref msg))
                        lblErrorLineas.Text = msg;

                    msg = string.Empty;

                    if (!ValidarIntervalo(ref msg))
                        lblErrorLineas.Text = msg;
                }
                else
                {
                    lblMsg.Text = "No existen costos Confirmados para este Espacio – Vigencia.";
                    ASPxPageControl1.Visible = true;
                    trPauta.Visible = false;
                    trFind.Visible = true;

                }
            }
        }

        //HTTD
        protected void btnDelete_Click(object sender, EventArgs e)
        {
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            gv.Selection.UnselectAll();

            try
            {
                decimal duracionTot = 0;

                CertificadoCabDTO certificado = CertificadoCab;

                certificado.IdentifFrecuencia = Convert.ToString(ucIdentifFrecuencia.SelectedValue);
                certificado.HoraInicio        = FormsHelper.ConvertToTimeSpan(teHoraInicio.DateTime);
                certificado.HoraFin           = FormsHelper.ConvertToTimeSpan(teHoraFin.DateTime);
                certificado.IdentifIntervalo  = Convert.ToString(ucIdentifIntervalo.SelectedValue);

                //SOLO PARA LA VALIDACION

                List<CertificadoCabDTO> Certis = Certificados.ReadAll("PAUTAID = '" + certificado.PautaId + "' AND NOT IDENTIFORIGEN IS NULL");

                if(Certis.Count > 0)
                {

                    //salvo el registro actual
                    Certificados.Update(certificado, "PAUTAID = '" + certificado.PautaId + "' AND IDENTIFORIGEN = '" + certificado.IdentifOrigen + "'");

                    for (int x = 0; x <= Certis.Count - 1; x++)
                    {
                        Certis[x].CertValido    = string.Empty;
                        Certis[x].FecCertValido = (DateTime)System.Data.SqlTypes.SqlDateTime.Null;

                        Certificados.Update(Certis[x], "PAUTAID = '" + certificado.PautaId + "' AND NOT IDENTIFORIGEN IS NULL AND NOT IDENTIFORIGEN = '" + certificado.IdentifOrigen + "'");
                    }
                }

                //
                //Sumatoria de duracion en registros de la tabla CertificadoDet.
                {
                    mycert.ForEach(x => { duracionTot += (x.Duracion != null) ? x.Duracion.Value : 0; });
                    certificado.DuracionTot = duracionTot;
                }

                //Cantidad de registros tabla CertificadoDet cuyo campo IdentifAviso <> “”
                certificado.CantSalidas = mycert.FindAll(x => (x.IdentifAviso != null && x.IdentifAviso != string.Empty)).Count;

                if (certificado.CantSalidas != 0)
                {
                    if (certificado.RecId == 0)
                    {
                        //Es nuevo...
                        certificado.PautaId = Business.Certificados.GetNextPautaId().ToString();

                        if (certificado.IdentifIntervalo == "") certificado.IdentifIntervalo = null;

                        Certificados.Create(certificado, mycert);
                    }
                    else
                    {
                        //Es modificacion...
                        if (mycert.Count > 0)
                        {
                            if (mycert[0].Salida > 0)
                            {
                                Labelx.Text = string.Empty;

                                lblErrorLineas.Text = Labelx.Text;

                                if (teHoraInicio.Text == mycert[0].Hora.ToString().Substring(0, 5) && teHoraFin.Text == mycert[0].Hora.ToString().Substring(0, 5))
                                {
                                    Labelx.Text = string.Empty;

                                    lblErrorLineas.Text = Labelx.Text;
                                }
                                else
                                {
                                    if (Convert.ToDateTime(teHoraInicio.Text) > (Convert.ToDateTime(teHoraFin.Text)))
                                        Labelx.Text = "Hora de inicio < Hora de Fin. \r\n";

                                    Labelx.Text += " La hora de inicio/ fin no se corresponde con registros previamente grabados. NO SE HA MODIFICADO LA CABECERA";

                                    lblErrorLineas.Text = Labelx.Text;

                                    return;
                                }
                            }
                            else
                            {
                                Labelx.Text = string.Empty;

                                lblErrorLineas.Text = Labelx.Text;
                            }
                        }

                        if (certificado.IdentifIntervalo == "")
                            certificado.IdentifIntervalo = null;

                        Certificados.Update(certificado, mycert);
                    }

                    CertificadoCab = certificado;

                    Business.Certificados.CalcularCosto(CertificadoCab, Costos, mycert, ((Accendo)this.Master).Usuario.UserName);

                    ReCargarControles(CertificadoCab);

                }
                else if (certificado.CantSalidas == 0)
                {
                    throw new Exception("Debe Insertar lineas en Detalle para poder Guardar.");
                }

                lblErrorLineas.Text = "Se Grabo correctamente";

                Labelx.Text = lblErrorLineas.Text;

            }
            catch (Exception ex)
            {
                lblErrorLineas.Text = ex.Message;
            }
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            trAccion.Visible = false;
        }

        protected void mnuDetalle_ItemClick(object source, MenuItemEventArgs e)
        {
            string msg = string.Empty;

            if (!ValidarFranjaHoraria(ref msg))
                lblErrorLineas.Text = msg;

            switch (e.Item.Name)
            {
                case "btnInsert": InsertItems(); break;
                case "btnEdit": EditItem(); break;
                case "btnDelete": DeleteItems(); break;
                case "btnSelectAll":
                    {
                        int PosicionActual = gv.VisibleStartIndex;

                        int Maximo = PosicionActual + gv.SettingsPager.PageSize - 1;

                        for (int i = PosicionActual; i <= Maximo; i++)
                        {

                            gv.Selection.SelectRow(i);
                        }

                        break;
                    }
                case "btnSelect0":
                    {
                        for (int i = 0; i <= gv.SettingsPager.PageSize - 1; i++)
                        {
                            gv.Selection.SelectRow(i);
                        }

                        break;
                    }

                case "btnCopy": CopyItems(); break;
                case "btnReplace": ReplaceItems(); break;
                case "btnSKU": QuerySKU(); break;
                default: break;
            }
        }

        private void QuerySKU()
        {
            trQuerySKU.Visible = true;

            RefreshSKUGrid(gvSKU);
        }

        #region "Actions Buttons"
        private void EditItem()
        {
            if (FormsHelper.GetSelectedId(gv) == null)
            {
                lblErrorLineas.Text = "Debe Seleccionar una linea para Modificar";
            }
            else if (gv.Selection.Count != 1)
            {
                lblErrorLineas.Text = "Solo puede modificar una linea por vez. Quite la marca a las que no modificara en este momento.";
            }
            else
            {
                try
                {
                    var id = FormsHelper.GetSelectedId(gv);

                    if (id.HasValue)
                    {
                        var linea = mycert.Find(x => x.RecId == (int)gv.GetSelectedFieldValues("RecId")[0]);

                        //Para tener el ID en el momento en que vaya a guardar.
                        pnlEditLine.Attributes.Add("RecId", linea.RecId.ToString());

                        ucIdentifAvisoEdit.SelectedValue = linea.IdentifAviso;

                        AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoEdit.SelectedValue));

                        if (aviso != null)
                        {
                            spDuracionEdit.Value            = aviso.Duracion;
                            ASPxTimeEdit miTimeEdit         = new ASPxTimeEdit();
                            miTimeEdit.DateTime             = Convert.ToDateTime(gv.GetSelectedFieldValues("Hora")[0].ToString());
                            teHoraInicioModificar.DateTime  = miTimeEdit.DateTime;
                            spAvisoModifiSalidas.Value      = gv.GetSelectedFieldValues("Salida")[0].ToString();
                            teHoraInicioModificar.Enabled   = spAvisoModifiSalidas.Text == "0";
                            spAvisoModifiSalidas.Enabled    = teHoraInicioModificar.Enabled == false;
                        }

                        gv.CancelEdit();

                        trEditLine.Visible = true;

                        trAccion.Visible = false;
                    }
                }
                catch (Exception ex)
                {
                    MsgErrorLinas(ex);
                }
            }
        }

        private void ReplaceItems()
        {
            ASPxPageControl2.ActiveTabPage = ASPxPageControl2.TabPages[2];
            trAccion.Visible = true;
            trEditLine.Visible = false;
        }

        private void CopyItems()
        {
            ASPxPageControl2.ActiveTabPage = ASPxPageControl2.TabPages[1];
            trAccion.Visible = true;
            trEditLine.Visible = false;
        }

        private void InsertItems()
        {
            teHoraInicioInsertar.DateTime   = teHoraInicio.DateTime;
            teHoraFinInsertar.DateTime      = teHoraFin.DateTime;
            ASPxPageControl2.ActiveTabPage  = ASPxPageControl2.TabPages[0];
            spSalidasInsertar.Enabled       = (teHoraInicio.Text == teHoraFin.Text && teHoraInicio.Text == "00:00");
            spSalidasInsertar.Value         = spSalidasInsertar.Enabled ? 1 : 0;
            teHoraInicioInsertar.Enabled    = (spSalidasInsertar.Enabled == false);
            teHoraFinInsertar.Enabled       = (spSalidasInsertar.Enabled == false);
            trAccion.Visible                = true;
            trEditLine.Visible              = false;
        }

        private void DeleteItems()
        {
            try
            {
                if (FormsHelper.GetSelectedId(gv) != null)
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.CertificadoDetDTO> aux = new List<DTO.CertificadoDetDTO>();

                    //Creo una nueva coleccion con todos los productos menos los seleccionados, y la guardo en el Viewstate.
                    foreach (var linea in mycert)

                        if (!FormsHelper.IsSelectedRecId(linea.RecId, gv))
                            aux.Add(linea);

                    mycert        = aux;
                    gv.DataSource = mycert;

                    gv.Selection.UnselectAll();

                    RefreshAbmGrid(gv);

                    gv.Selection.UnselectAll();

                    lblErrorLineas.Text = "Linea eliminada correctamente. No olvide Guardar antes de salir";
                }
                else
                {
                    lblErrorLineas.Text = "Debe seleccionar una linea para eliminar.";
                }
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }
        #endregion

        private void MsgErrorLinas(Exception ex)
        {
            lblErrorLineas.Text = ex.Message.ToLower().Contains("Violation of UNIQUE KEY constraint".ToLower()) == true ? "No puede ingresar un nuevo registro con su clave duplicada." :
                                  ex.Message.ToLower().Contains("FK".ToLower()) == true ? "No puede modificar/eliminar la clave de este registro, ya que se encuentra relacionado a otra entidad." :
                                  ex.Message;
        }

        #region "Solapas internas"
        protected void btnCopiarPeriodos_Click(object sender, EventArgs e)
        {
            try
            {
                CertificadoCabDTO certificado = CertificadoCab;

                DateTime fechaTopeMes = new DateTime(Convert.ToInt32(certificado.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(certificado.AnoMes.ToString().Substring(4, 2)), 1).AddMonths(1).AddDays(-1);

                //Validaciones de Origen
                
                //FECHA DESDE ORIGEN NO NULA
                if (deFechaDesdeOrigenCopiar.Value == null)
                    throw new Exception("Debe ingresar una fecha DESDE ORIGEN.");

                //FECHA DESDE ORIGEN DENTRO DEL PERIODO DEL ORDENADO
                if (deFechaDesdeOrigenCopiar.Date.Year.ToString() + deFechaDesdeOrigenCopiar.Date.Month.ToString().PadLeft(2, '0') != certificado.AnoMes.ToString())
                    throw new Exception("El mes/año en DESDE ORIGEN no corresponde al ordenado seleccionado.");

                //FECHA HASTA ORIGEN NO NULA
                if (deFechaHastaOrigenCopiar.Value == null)
                    throw new Exception("Debe ingresar una fecha HASTA ORIGEN.");

                //FECHA HASTA ORIGEN DENTRO DEL PERIODO DEL ORDENADO
                if (deFechaHastaOrigenCopiar.Date.Year.ToString() + deFechaHastaOrigenCopiar.Date.Month.ToString().PadLeft(2, '0') != certificado.AnoMes.ToString())
                    throw new Exception("El mes/año en HASTA ORIGEN no corresponde al ordenado seleccionado.");


                //FECHA HASTA ORIGEN >=  DESDE ORIGEN
                if (Convert.ToDateTime(deFechaHastaOrigenCopiar.Value) < Convert.ToDateTime(deFechaDesdeOrigenCopiar.Value))
                    throw new Exception("La fecha HASTA ORIGEN debe ser superior o igual a la fecha DESDE ORIGEN.");

                //CALCULO DE DIAS POR COPIAR
                int DiasPorCopiar = (deFechaHastaOrigenCopiar.Date.Day - deFechaDesdeOrigenCopiar.Date.Day) + 1;

                //Validaciones de Destino

                //FECHA DESDE DESTINO NO NULA
                if (deFechaDesdeDestinoCopiar.Value == null)
                    throw new Exception("Debe ingresar una fecha DESDE DESTINO.");

                //FECHA DESDE DESTINO DENTRO DEL PERIODO DEL ORDENADO
                if (deFechaDesdeDestinoCopiar.Date.Year.ToString() + deFechaDesdeDestinoCopiar.Date.Month.ToString().PadLeft(2, '0') != certificado.AnoMes.ToString())
                    throw new Exception("El mes/año en DESDE DESTINO no corresponde al ordenado seleccionado.");

                bool EntraArriba = false;

                EntraArriba = (deFechaDesdeDestinoCopiar.Date.Day + DiasPorCopiar <= deFechaDesdeOrigenCopiar.Date.Day);

                if (deFechaDesdeOrigenCopiar.Date.Day < deFechaDesdeDestinoCopiar.Date.Day)
                    EntraArriba = true;

                /// Validaciones Generales ///
                /// 
                if (!EntraArriba)
                    throw new Exception("El espacio donde copiar es inferior al espacio seleccionado.");

                if ((deFechaDesdeDestinoCopiar.Date.Day + DiasPorCopiar -1) > Convert.ToInt32(fechaTopeMes.Day))
                    throw new Exception("LOS DIAS A COPIAR SUPERAN LA FECHA MAXIMA DE ESTA PAUTA");

                /// AHORA HABRIA QUE BORRAR EN LA LISTA ORIGEN LA CANTIDAD DE DIAS POR COPIAR A PARTIR DE LA FECHA DESDE ORIGEN PARA PODER INSERTAR EL PERIODO SELECCIONADO
                /// 

                ///.. PROBEMOS A VER SI SALE
                ///
                for (int w = 0; w <= Lineas.Count - 1; w++)
                {
                    for (int x = 0; x <= Lineas.Count - 1; x++)
                    {
                        if (Lineas[x].Dia >= deFechaDesdeDestinoCopiar.Date.Day &&
                            Lineas[x].Dia <= deFechaDesdeDestinoCopiar.Date.Day + DiasPorCopiar -1)
                        {
                            Lineas.RemoveAt(x);

                            break;
                        }
                    }

                }
               
                ///
                //3. COPIAR PERIODO: PERMITE COPIAR UN RANGO ORIGEN EN UN RANGO DESTINO QUE NO PERTENECE A LA FRECUENCIA INDICADA POR LA PAUTA
                //VALIDAR QUE EL RANGO DESTINO ESTE DENTRO DE LA FRECUENCIA INDICADA POR LA PAUTA

                Lineas = PeriodosCopiar(deFechaDesdeOrigenCopiar.Date, deFechaHastaOrigenCopiar.Date, deFechaDesdeDestinoCopiar.Date,deFechaHastaDestinoCopiar.Date, Lineas);

                RefreshAbmGrid(gv);
                
                trAccion.Visible = false;
                
                lblErrorLineas.Text = "No olvide GRABAR antes de continuar.";
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }


        private List<CertificadoDetDTO> PeriodosCopiar(DateTime fOrigenDesde, DateTime fOrigenHasta, DateTime fDestinoDesde, DateTime fDestinoHasta, List<CertificadoDetDTO> lineas)
        {
            List<CertificadoDetDTO> LineasOrigen = lineas;
            List<CertificadoDetDTO> LineasSeleccionadas = new List<CertificadoDetDTO>();
            List<CertificadoDetDTO> LineasDestino = lineas;

            FrecuenciaDTO frecuencia = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));

            List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

            lineas = mycert.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

            EspacioContDTO espacio = GetEspacioContenido();

            string[,] DiasSemana = { { "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO" }, { "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY" } };

            var lineasACopiar = lineas.FindAll((x) => (x.Fecha.DayOfYear >= fOrigenDesde.DayOfYear  && x.Fecha.DayOfYear <= fOrigenHasta.DayOfYear ));

            DateTime origenM;
            // ARMADO DE LINEAS SELECCIONADAS //
            for (int i = 0; i <= LineasOrigen.Count - 1; i++)
            {
                origenM = new DateTime(LineasOrigen[i].Fecha.Year, LineasOrigen[i].Fecha.Month, LineasOrigen[i].Fecha.Day);

                if (origenM >= fOrigenDesde && origenM <= fOrigenHasta)
                {
                    LineasSeleccionadas.Add(lineas[i]);
                }
            }

            //ORDENA LAS LINEAS SELECCIONADAS
            LineasSeleccionadas = LineasSeleccionadas.OrderBy(o => o.Dia).ThenBy(p => p.Hora).ThenBy(q => q.Salida).ToList();

            //CALCULO CANTIDAD DE DIAS EN EL MES
            int DiasEnMes = System.DateTime.DaysInMonth(LineasOrigen[0].Fecha.Year, LineasOrigen[0].Fecha.Month);

            //CALCULO CUAL ES EL DIA DE LA SEMANA DEL 1 DEL MES
            int DiaSemana = Convert.ToInt32(Convert.ToDateTime( LineasOrigen[0].Fecha.Year.ToString() + "-" + LineasOrigen[0].Fecha.Month.ToString("00") + "-" + "01").DayOfWeek);

            // RECORRO UNO A UNO LOS DIAS DEL MES
            for (int i = 1; i <= DiasEnMes; i++)
            {
                // PREGUNTO SI EL DIA DEL MES ESTA DENTRO DEL RANGO DE DIAS SELECCIONADOS COMO DESTINO
                if (i >= fDestinoDesde.Day && i <= fDestinoHasta.Day)
                {
                    DateTime muleto = new DateTime(LineasSeleccionadas[0].Fecha.Year, LineasSeleccionadas[0].Fecha.Month, i);
                    string diamuleto = muleto.DayOfWeek.ToString().Trim().ToUpper();
                    
                    for (int x = 0; x <= DiasSemana.Length - 1; x++)
                    {
                        if (DiasSemana[1, x].ToUpper() == diamuleto)
                        {
                            diamuleto = DiasSemana[0, x].ToUpper();

                            break;
                        }
                    }

                    var listamuleto = frecuenciaDetalles.FindAll(q => q.DiaSemana.Trim().ToUpper() == diamuleto);

                    if (listamuleto.Count > 0)
                    {
                        //BUSCO A VER SI EXISTEN LINEAS PREVIAS CON DATOS PARA ESA FECHA
                        var newList = LineasOrigen.FindAll(s => s.Fecha.Day == i);
                        
                        DAO.CertificadoDetDAO odd = new DAO.CertificadoDetDAO();
                        
                        int LastId = odd.GetLastRecId();

                        if (newList.Count > 0)
                            LineasDestino.RemoveAll(p => p.Fecha.Day == i);

                        //NO HAY DATOS PREEXISTENTES PARA ESE DIA DEL MES
                        DateTime diaActual = new DateTime(LineasOrigen[0].Fecha.Year, LineasOrigen[0].Fecha.Month, i);

                        // CREO UNA NUEVA LINEA CON LOS DATOS DE LA PRESELECCION PARA ESE DIA DE LA SEMANA
                        for (int j = 0; j <= LineasSeleccionadas.Count - 1; j++)
                        {
                            var miLineaSeleccionada = LineasSeleccionadas[j].DiaSemana.ToUpper().Trim().Replace("É", "E").Replace("Á", "A");
                            var lista = frecuenciaDetalles.FindAll(q => q.DiaSemana.Trim() == miLineaSeleccionada);

                            if (lista.Count == 1)
                            {
                                CertificadoDetDTO newLine = new CertificadoDetDTO();

                                LastId++;
                                newLine.DatareaId  = 0;
                                newLine.RecId      = LastId;
                                newLine.Costo      = LineasSeleccionadas[j].Costo;
                                newLine.CostoOp    = LineasSeleccionadas[j].CostoOp;
                                newLine.CostoOpUni = LineasSeleccionadas[j].CostoOpUni;
                                newLine.CostoUni   = LineasSeleccionadas[j].CostoUni;
                                newLine.Dia        = i;
                                newLine.Fecha      = new DateTime(LineasSeleccionadas[j].Fecha.Year, LineasSeleccionadas[j].Fecha.Month, i);

                                string dia = newLine.Fecha.DayOfWeek.ToString().ToUpper();

                                for (int k = 0; k <= DiasSemana.Length - 1; k++)
                                {
                                    if (DiasSemana[1, k].ToUpper() == dia)
                                    {
                                        newLine.DiaSemana = DiasSemana[0, k].ToUpper();
                                        break;
                                    }
                                }

                                newLine.Duracion     = LineasSeleccionadas[j].Duracion;
                                newLine.Hora         = LineasSeleccionadas[j].Hora;
                                newLine.IdentifAviso = LineasSeleccionadas[j].IdentifAviso;
                                newLine.PautaId      = LineasSeleccionadas[j].PautaId;
                                newLine.Salida       = LineasSeleccionadas[j].Salida;

                                LineasDestino.Add(newLine);
                            }
                        }
                    }
                }

                DiaSemana = DiaSemana == 7 ? 1 : DiaSemana++;
            }
            mycert = LineasDestino.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

            lineas = LineasDestino;

            return lineas;
        }

        private List<CertificadoDetDTO> CopiarPeriodos(DateTime fOrigenDesde, DateTime fOrigenHasta, DateTime fDestinoDesde, List<CertificadoDetDTO> lineas)
        {
            bool preExistentesFlag = false;
            EspacioContDTO espacio = GetEspacioContenido();

            //Armo lista de elemtnos que SI reemplazo.
            //Busco en la coleccion, todas las lineas en el periodo, y con el aviso seleccionado.
            var lineasACopiar = lineas.FindAll(
                (x) =>
                    (x.Fecha >= fOrigenDesde
                    && x.Fecha <= fOrigenHasta
                    ));


            //Si encontre líneas a copiar...
            if (lineasACopiar.Count > 0)
            {
                CertificadoDetDTO nuevaLinea;
                DateTime fechaTmp;
                List<CertificadoDetDTO> lineasTmp     = new List<CertificadoDetDTO>();
                List<CertificadoDetDTO> preExistentes = new List<CertificadoDetDTO>();

                DateTime diasEnElFuturo = fOrigenDesde;
                int cantDias;
                fechaTmp = fDestinoDesde;

                cantDias = 0;
                //Por cada linea que encontre, genero una nueva e igual, x dias en el futuro.
                foreach (var linea in lineasACopiar)
                {
                    FrecuenciaDTO frecuencia                  = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                    List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                    nuevaLinea = new CertificadoDetDTO();
                    nuevaLinea.RecId     = NextTempRecId();
                    nuevaLinea.DatareaId = linea.DatareaId;

                    //cargo lineas nuevas

                    if (fOrigenDesde != linea.Fecha)
                    {
                        cantDias++;
                        //cantDias = Convert.ToInt32(linea.Fecha.DayOfYear) - Convert.ToInt32(diasEnElFuturo.DayOfYear);
                        diasEnElFuturo = linea.Fecha;
                        fechaTmp = fechaTmp.AddDays(cantDias);
                    }

                    nuevaLinea.Fecha = fDestinoDesde.AddDays(cantDias);
                    nuevaLinea.Dia = nuevaLinea.Fecha.Day;

                    string[] sDias = { "DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO" };

                    for (int x = 0; x <= 6; x++)
                    {
                        if ((int)nuevaLinea.Fecha.DayOfWeek == x)
                        {
                            nuevaLinea.DiaSemana = sDias[x];
                        }
                    }

                    //evito que se carguen datos fuera de los dias pautados
                    bool retval = false;
                    for (int k = 0; k <= frecuenciaDetalles.Count - 1; k++)
                    {
                        if (nuevaLinea.DiaSemana.Trim() == frecuenciaDetalles[k].DiaSemana.Trim())
                        {
                            retval = true;
                            break;
                        }
                    }

                    if (retval == false)
                    {
                        throw new Exception("No se puede grabar en dias de semana distintos a los pautados.");
                    }
                    ////////////////////////////////////////////////////////
                    nuevaLinea.Hora         = linea.Hora        ;
                    nuevaLinea.Costo        = linea.Costo       ;
                    nuevaLinea.CostoOp      = linea.CostoOp     ;
                    nuevaLinea.CostoOpUni   = linea.CostoOpUni  ;
                    nuevaLinea.CostoUni     = linea.CostoUni    ;
                    nuevaLinea.Duracion     = linea.Duracion    ;
                    nuevaLinea.IdentifAviso = linea.IdentifAviso;
                    nuevaLinea.PautaId      = linea.PautaId     ;
                    nuevaLinea.Salida       = linea.Salida      ;

                    foreach (CertificadoDetDTO l in lineas)
                    {
                        if (l.Fecha == fechaTmp.Date)
                        {
                            if (l.Hora == linea.Hora)
                            {
                                if (l.Salida == nuevaLinea.Salida)
                                {
                                    //preExistentes.Add(nuevaLinea);
                                    preExistentes.Add(linea);
                                }
                            }
                        }
                    }

                    lineasTmp.Add(nuevaLinea);//Agrego la nueva linea.
                }

                //Junto las dos listas (temporal y la que ya tenia).
                lineasTmp.AddRange(lineas);

                //Ordeno por fecha.
                lineasTmp.Sort((x, y) => DateTime.Compare(x.Fecha, y.Fecha));

                if (preExistentesFlag == true)
                {
                    lblErrorLineas.Text = "No se pudieron grabar todas las lineas. No olvide GRABAR antes de continuar.";
                }

                //Guardo la lista en el Viewstate.
                gv.DataSource = lineasTmp;
                return lineasTmp;
            }
            else
            {
                throw new Exception("No hay avisos para copiar dentro del rango seleccionado");
            }
        }

        protected void btnReemplazarAvisos_Click(object sender, EventArgs e)
        {
            string avisoOrigen;

            string avisoDestino;

            decimal salidas = 0;

            try
            {
                //Validaciones generales...
                if (ucIdentifAvisoOrigenReemplazar.SelectedValue != null)
                {
                    avisoOrigen = ucIdentifAvisoOrigenReemplazar.SelectedValue.ToString();
                }
                else
                {
                    //avisoOrigen = string.Empty;
                    lblErrorLineas.Text = "El aviso de origen está en blanco. NO se grabo ningún registro.";

                    return;

                }

                if (ucIdentifAvisoDestinoReemplazar.SelectedValue != null)
                {
                    avisoDestino = ucIdentifAvisoDestinoReemplazar.SelectedValue.ToString();
                }
                else
                {
                    lblErrorLineas.Text = "No se puede reemplazar un aviso con espacios en blanco. NO se grabo ningún registro.";

                    return;
                }


                if (opEditPeriodo.Checked)
                {
                    //Validaciones...
                    if (deFechaDesdeReemplazar.Value == null)
                        throw new Exception("Debe seleccionar una 'Fecha Desde' (origen).");

                    if (deFechaHastaReemplazar.Value == null)
                        throw new Exception("Debe seleccionar una 'Fecha Hasta' (origen).");

                    mycert = ReemplazarAvisosPorPeriodo(deFechaDesdeReemplazar.Date.AddHours(deHoraDesdeOrigenReemplazar.DateTime.Hour).AddMinutes(deHoraDesdeOrigenReemplazar.DateTime.Minute),
                                                        deFechaHastaReemplazar.Date.AddHours(deHoraHastaOrigenReemplazar.DateTime.Hour).AddMinutes(deHoraHastaOrigenReemplazar.DateTime.Minute),
                                                        avisoOrigen,
                                                        avisoDestino,
                                                        mycert,
                                                        salidas);
                    RefreshAbmGrid(gv);

                    trAccion.Visible = false;
                }
                else if (opEditSeleccionados.Checked)
                {
                    mycert = ReemplazarAvisosSeleccionados(avisoOrigen, avisoDestino, mycert, salidas);
                }
                else if (opEditTodas.Checked)
                {
                    mycert = ReemplazarAvisosTodos(avisoOrigen, avisoDestino, mycert, salidas);
                }

                RefreshAbmGrid(gv);

                trAccion.Visible = false;

                gv.Selection.UnselectAll();

            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }

        private List<CertificadoDetDTO> ReemplazarAvisosTodos(string avisoOrigen, string avisoDestino, List<CertificadoDetDTO> lineas, decimal salidas)
        {

            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, DateTime.Now, DateTime.Now, "Todos");
        }

        private List<CertificadoDetDTO> ReemplazarAvisosSeleccionados(string avisoOrigen, string avisoDestino, List<CertificadoDetDTO> lineas, decimal salidas)
        {

            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, DateTime.Now, DateTime.Now, "Seleccionado");

        }

        private List<CertificadoDetDTO> ReemplazarAvisosPorPeriodo(DateTime fDesde, DateTime fHasta, string avisoOrigen, string avisoDestino, List<CertificadoDetDTO> lineas, decimal salidas)
        {

            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, fDesde, fHasta, "Ingresado");

        }

        private List<CertificadoDetDTO> ReemplazarAvisos(string avisoOrigen, string avisoDestino, List<CertificadoDetDTO> lineas, decimal salidas, DateTime fDesde, DateTime fHasta, string Tipo = "Todo")
        {
            switch (Tipo)
            {
                case "Ingresado":
                    {
                        //Busco las línas con el periodo seleccionado, y cuyo aviso, sea igual al 'avisoOrigen',
                        //Para cada una de las líneas encontradas, reemplazo el aviso.

                        string FechaInicio = string.Empty;

                        string FechaFinal = string.Empty;

                        FechaInicio = fDesde.ToShortDateString();

                        FechaFinal = fHasta.ToShortDateString();

                        lineas.FindAll(
                            (x) =>
                                (Convert.ToDateTime(x.Fecha.ToShortDateString()) >= Convert.ToDateTime(FechaInicio) && Convert.ToDateTime(x.Fecha.ToShortDateString()) <= Convert.ToDateTime(FechaFinal) && x.IdentifAviso == avisoOrigen)).ForEach(
                                (linea) =>
                                {
                                    if (Convert.ToInt32(spAvisoReempDuracion.Value) > Convert.ToInt32(linea.Duracion))
                                    {
                                        TimeSpan tsPautaInicio = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicio.Text));
                                        TimeSpan tsPautaFin    = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraFin.Text));
                                        TimeSpan tsAvisoViejoi = linea.Hora;
                                        TimeSpan tsAvisoViejod = new TimeSpan(0, 0, Convert.ToInt32(linea.Duracion));
                                        TimeSpan tsAvisoViejof = tsAvisoViejoi.Add(tsAvisoViejod);
                                        TimeSpan tsAvisoNuevoi = linea.Hora;
                                        TimeSpan tsAvisoNuevod = new TimeSpan(0, 0, Convert.ToInt32(spAvisoReempDuracion.Value));
                                        TimeSpan tsAvisoNuevof = tsAvisoNuevoi.Add(tsAvisoNuevod);

                                        for (int i = 0; i <= lineas.Count - 1; i++)
                                        {
                                            if (lineas[i].Fecha.DayOfYear == linea.Fecha.DayOfYear)
                                            {
                                                IntervaloDTO intervalo = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));

                                                if (tsAvisoNuevod.Minutes > intervalo.CantMinutos)
                                                {
                                                    lblErrorLineas.Text = "Error de solapamiento de avisos. Verifique";

                                                    break;
                                                }
                                                else
                                                {
                                                    if (linea.Fecha == lineas[i].Fecha)
                                                    {
                                                        if (lineas[i].Hora > linea.Hora)
                                                        {
                                                            if (lineas[i].Hora < tsAvisoNuevof)
                                                            {
                                                                lblErrorLineas.Text = "Error de solapamiento de avisos. Verifique";

                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (linea.Hora >= FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(deHoraDesdeOrigenReemplazar.Value)) && linea.Hora <= FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(deHoraHastaOrigenReemplazar.Value)))
                                        {
                                            linea.IdentifAviso  = avisoDestino;
                                            linea.Salida        = salidas;
                                            linea.Duracion      = Convert.ToDecimal(spAvisoReempDuracion.Value);
                                            lblErrorLineas.Text = "No olvide GRABAR antes de continuar.";
                                        }
                                    }
                                });
                        break;

                    }
                case "Seleccionado":
                    {

                        //Obtengo todos los Ids de los registros seleccionados.
                        List<object> ids = gv.GetSelectedFieldValues("RecId");

                        //Busco las línas con los RecId seleccionados, y cuyo aviso, sea igual al 'avisoOrigen',
                        //Para cada una de las líneas encontradas, reemplazo el aviso.
                        lineas.FindAll(
                            (x) => ids.Contains(x.RecId) && x.IdentifAviso == avisoOrigen).ForEach(
                            (linea) =>
                            {

                                //Si la duración del aviso nuevo es mayor que la del viejo
                                if (Convert.ToInt32(spAvisoReempDuracion.Value) > Convert.ToInt32(linea.Duracion))
                                {
                                    TimeSpan tsPautaInicio = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicio.Text));
                                    TimeSpan tsPautaFin    = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraFin.Text));
                                    TimeSpan tsAvisoViejoi = linea.Hora;
                                    TimeSpan tsAvisoViejod = new TimeSpan(0, 0, Convert.ToInt32(linea.Duracion));
                                    TimeSpan tsAvisoViejof = tsAvisoViejoi.Add(tsAvisoViejod);
                                    TimeSpan tsAvisoNuevoi = linea.Hora;
                                    TimeSpan tsAvisoNuevod = new TimeSpan(0, 0, Convert.ToInt32(spAvisoReempDuracion.Value));
                                    TimeSpan tsAvisoNuevof = tsAvisoNuevoi.Add(tsAvisoNuevod);

                                    for (int i = 0; i <= lineas.Count - 1; i++)
                                    {
                                        if (lineas[i].Fecha.DayOfYear == linea.Fecha.DayOfYear)
                                        {
                                            IntervaloDTO intervalo = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));

                                            if (tsAvisoNuevod.Minutes > intervalo.CantMinutos)
                                            {
                                                lblErrorLineas.Text = "Error de solapamiento de avisos. Verifique";

                                                break;
                                            }
                                            else
                                            {
                                                if (linea.Fecha == lineas[i].Fecha)
                                                {
                                                    if (lineas[i].Hora > linea.Hora)
                                                    {
                                                        if (lineas[i].Hora < tsAvisoNuevof)
                                                        {
                                                            lblErrorLineas.Text = "Error de solapamiento de avisos. Verifique";

                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    linea.IdentifAviso  = avisoDestino;
                                    linea.Salida        = salidas;
                                    linea.Duracion      = Convert.ToDecimal(spAvisoReempDuracion.Value);
                                    lblErrorLineas.Text = "No olvide GRABAR antes de continuar.";
                                }

                            });

                        break;
                    }
                case "Todos":
                    {

                        //Por cada linea cuto identifAviso sea igual al 'avisoOrigen', reemplazo el aviso.
                        lineas.ForEach(linea =>
                        {
                            if (linea.IdentifAviso == avisoOrigen)
                            {
                                if (Convert.ToInt32(spAvisoReempDuracion.Value) > Convert.ToInt32(linea.Duracion))
                                {

                                    TimeSpan tsPautaInicio = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicio.Text));
                                    TimeSpan tsPautaFin    = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraFin.Text));
                                    TimeSpan tsAvisoViejoi = linea.Hora;
                                    TimeSpan tsAvisoViejod = new TimeSpan(0, 0, Convert.ToInt32(linea.Duracion));
                                    TimeSpan tsAvisoViejof = tsAvisoViejoi.Add(tsAvisoViejod);
                                    TimeSpan tsAvisoNuevoi = linea.Hora;
                                    TimeSpan tsAvisoNuevod = new TimeSpan(0, 0, Convert.ToInt32(spAvisoReempDuracion.Value));
                                    TimeSpan tsAvisoNuevof = tsAvisoNuevoi.Add(tsAvisoNuevod);

                                    for (int i = 0; i <= lineas.Count - 1; i++)
                                    {
                                        if (lineas[i].Fecha.DayOfYear == linea.Fecha.DayOfYear)
                                        {
                                            IntervaloDTO intervalo = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));

                                            if (tsAvisoNuevod.Minutes > intervalo.CantMinutos)
                                            {
                                                lblErrorLineas.Text = "Error de solapamiento de avisos. Verifique";

                                                break;
                                            }
                                            else
                                            {
                                                if (linea.Fecha == lineas[i].Fecha)
                                                {
                                                    if (lineas[i].Hora >= linea.Hora)
                                                    {
                                                        if (lineas[i].Hora < tsAvisoNuevof)
                                                        {
                                                            lblErrorLineas.Text = "Error de solapamiento de avisos. Verifique";

                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    linea.IdentifAviso  = avisoDestino;
                                    linea.Salida        = salidas;
                                    linea.Duracion      = Convert.ToDecimal(spAvisoReempDuracion.Value);
                                    lblErrorLineas.Text = "No olvide GRABAR antes de continuar.";
                                }
                            }
                        });

                        break;
                    }
            }

            return lineas;   
        }


        protected void opEditPeriodo_CheckedChanged(object sender, EventArgs e)
        {
        }

        protected void opEditSeleccionados_CheckedChanged(object sender, EventArgs e)
        {
        }

        protected void opEditTodas_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void ReemplazarAviso()
        {
            string avisoOrigen;
            string avisoDestino;
            DateTime fDesde;
            DateTime fHasta;

            try
            {
                //Validaciones generales...
                if (ucIdentifAvisoOrigenReemplazar.SelectedValue != null)
                {
                    avisoOrigen = ucIdentifAvisoOrigenReemplazar.SelectedValue.ToString();
                }
                else
                {
                    avisoOrigen = string.Empty;
                    lblErrorLineas.Text = "El aviso de origen está en blanco. NO se grabo ningún registro.";
                    return;
                }

                if (ucIdentifAvisoDestinoReemplazar.SelectedValue != null)
                {
                    avisoDestino = ucIdentifAvisoDestinoReemplazar.SelectedValue.ToString();
                }
                else
                {
                    lblErrorLineas.Text = "No se puede reemplazar un aviso con espacios en blanco. NO se grabo ningún registro.";
                    return;
                }


                if (opEditPeriodo.Checked)
                {
                    //Validaciones...
                    if (deFechaDesdeReemplazar.Value == null)
                        throw new Exception("Debe seleccionar una 'Fecha Desde' (origen).");

                    if (deFechaHastaReemplazar.Value == null)
                        throw new Exception("Debe seleccionar una 'Fecha Hasta' (origen).");

                    fDesde = deFechaDesdeReemplazar.Date;
                    fHasta = deFechaHastaReemplazar.Date.AddHours(23.9999);

                    mycert = ProcesoHelper.ReemplazarAvisosPorPeriodo(fDesde, fHasta, avisoOrigen, avisoDestino, mycert);
                }
                else if (opEditSeleccionados.Checked)
                {
                    mycert = ProcesoHelper.ReemplazarAvisosSeleccionados(avisoOrigen, avisoDestino, gv.GetSelectedFieldValues("RecId"), mycert);
                }
                else if (opEditTodas.Checked)
                {
                    mycert = ProcesoHelper.ReemplazarAvisosTodos(avisoOrigen, avisoDestino, mycert);
                }

                RefreshAbmGrid(gv);
                trAccion.Visible = false;
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }
        #endregion

        protected void btnUpdateEdit_Click(object sender, EventArgs e)
        {
            try
            {
                //Guarde el RecId en un atributo del panel al momento de cargarlo.
                int RecId = Convert.ToInt32(pnlEditLine.Attributes["RecId"]);

                var lineas = mycert;

                lineas = lineas.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                decimal? duracion = 0;

                AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoEdit.SelectedValue));

                if (aviso != null)
                    duracion = aviso.Duracion;

                ///// RECALCULO DURACION DEL AVISO ///////////

                IntervaloDTO intervalo = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));

                var tiempo = intervalo.CantMinutos * 60;

                if (duracion > tiempo)
                {
                    throw new Exception("Solapamiento de horarios. Verifique.");
                }

                var LastRecX = lineas.Last();

                if (RecId != LastRecX.RecId)
                {
                    for (int x = 0; x <= lineas.Count - 1; x++)
                    {
                        if (lineas[x].RecId == RecId)
                        {
                            if (lineas[x].RecId != LastRecX.RecId)
                            {
                                ProxRecId = lineas[x + 1].RecId;

                                break;
                            }
                            else
                            {
                                ProxRecId = LastRecX.RecId;

                                break;
                            }
                        }
                    }

                    lineas.ForEach(y =>
                    {
                        if (y.RecId == ProxRecId)
                        {
                            var limiteSuperior = LastRecX.Hora;
                            TimeSpan tpHoraX   = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicioModificar.Text));
                            DateTime tpHoraY   = new DateTime(y.Fecha.Year, y.Fecha.Month, y.Fecha.Day, tpHoraX.Hours, tpHoraX.Minutes, 0);
                            decimal? tpX       = duracion;
                            tpHoraY            = tpHoraY.Add(new TimeSpan(0, 0, Convert.ToInt32(tpX)));
                            DateTime dt        = new DateTime(y.Fecha.Year, y.Fecha.Month, y.Fecha.Day, y.Hora.Hours, y.Hora.Minutes, 0);

                            if (tpHoraY >= dt && tpX > 0)
                            {
                                throw new Exception("Solapamiento de horarios. Verifique.");
                            }
                            else
                            {

                            }
                        }
                    }
                    );
                }

                ///// FIN DE FUNCION DE RECALCULAR DURACION //

                if (Convert.ToDateTime(teHoraInicioModificar.Text) < Convert.ToDateTime(teHoraInicio.Text) || Convert.ToDateTime(teHoraInicioModificar.Text) > Convert.ToDateTime(teHoraFin.Text))
                {
                    throw new Exception("La hora no esta dentro del horario pautado.");
                }

                decimal? tp = 0;

                TimeSpan tpHora;

                decimal tpSalida = 0;

                lineas = lineas.OrderBy(x => x.Dia).ThenBy(y => y.Hora).ThenBy(z => z.Salida).ToList();

                lineas.ForEach(x =>
                {
                    if (x.RecId == RecId)
                    {
                        tp = x.Duracion;

                        tpHora = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicioModificar.Text));

                        tpSalida = x.Salida;

                        tpHora = tpHora.Add(new TimeSpan(0, 0, Convert.ToInt32(tp)));

                        string HoraLimite = string.Empty;

                        if (x.Salida == 0)
                        {
                            var LastRec = lineas.Last();

                            if (x.Dia.ToString() == gv.GetSelectedFieldValues("Dia")[0].ToString())
                            {

                                if (RecId != LastRec.RecId)
                                {
                                    lineas.ForEach(y =>
                                    {
                                        if (y.RecId == ProxRecId)
                                        {
                                            HoraLimite = y.Hora.ToString();
                                        }
                                    }
                                    );

                                }
                                else
                                {
                                    HoraLimite = teHoraFin.Text;
                                }

                                //Deberia continuar por aca
                                if (Convert.ToDateTime(tpHora.ToString().Substring(0, 5) + ":00") >= Convert.ToDateTime(HoraLimite) && (int)Convert.ToChar(spDuracionEdit.Text) != 48)
                                {
                                    throw new Exception("Solapamiento de horarios. Verifique.");
                                }
                                else
                                {
                                    //Cuando encuentre el item que estaba editando, lo actualizo.
                                    x.IdentifAviso = Convert.ToString(ucIdentifAvisoEdit.SelectedValue);
                                    x.Duracion     = duracion;
                                    DateTime dt    = (DateTime)teHoraInicioModificar.DateTime;
                                    TimeSpan ts    = new TimeSpan(0, dt.Hour, dt.Minute, 0);
                                    x.Hora         = ts;
                                    x.Salida       = Convert.ToDecimal(spAvisoModifiSalidas.Value);
                                }
                            }
                        }
                    }
                }
                );

                mycert = lineas;

                trEditLine.Visible = false;

                RefreshAbmGrid(gv);

                gv.Selection.UnselectAll();
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }

        protected void btnCancelEdit_Click(object sender, EventArgs e)
        {
            trEditLine.Visible = false;

            gv.Selection.UnselectAll();
        }

        protected void btnCancelEdit2_Click(object sender, EventArgs e)
        {
            btnBack_Click(sender, null);
        }

        protected void gvHome_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {

            ASPxGridView gvHome = (ASPxGridView)sender;

            e.Cancel = true;

            //Do things...
            var aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", (string)e.NewValues["IdentifAviso"]));

            var lineas = mycert;

            lineas.FindAll(x => x.RecId == (int)e.Keys[0]).ForEach(
                (linea) =>
                {
                    linea.IdentifAviso = aviso.IdentifAviso;
                    linea.Duracion     = aviso != null ? aviso.Duracion : null;
                    linea.Salida       = (decimal)e.NewValues["Salida"];

                });

            mycert = lineas;

            gv.CancelEdit();

            RefreshAbmGrid(gv);

        }

        protected void gv_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            ASPxGridView gv = (ASPxGridView)sender;

            e.Cancel = true;

            //Do things...
            var aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", (string)e.NewValues["IdentifAviso"]));

            var lineas = mycert;

            lineas.FindAll(x => x.RecId == (int)e.Keys[0]).ForEach(
                (linea) =>
                {
                    linea.IdentifAviso = aviso.IdentifAviso;
                    linea.Duracion     = aviso != null ? aviso.Duracion : null;
                    linea.Salida       = (decimal)e.NewValues["Salida"];

                });

            mycert = lineas;

            gv.CancelEdit();

            RefreshAbmGrid(gv);
        }

        protected void gvHome_StartRowEditing(object sender, DevExpress.Web.Data.ASPxStartRowEditingEventArgs e)
        { 

        }

        protected void gv_StartRowEditing(object sender, DevExpress.Web.Data.ASPxStartRowEditingEventArgs e)
        {
            trEditLine.Visible = false;
        }

        protected void btnRefreshSKU_Click(object sender, EventArgs e)
        {
            RefreshSKUGrid(gvSKU);
        }

        #region "Orden de Publicidad"
        protected void EmitirOP()
        {
            EspacioContDTO myspace = GetEspacioContenido();
            string TipoOP = string.Empty;

            if (myspace != null)
                TipoOP = "OP_" + myspace.FormatoOP;
        }

        #endregion

        protected void mnuPrincipal_ItemClick(object source, MenuItemEventArgs e)
        {

            switch (e.Item.Name)
            {
                case "btnQuerySKU":

                    QuerySKU();

                    break;

                case "btnOP":

                    if (gvHome.Selection.Count == 1)
                    {

                        string Pauta = string.Empty;

                        Pauta = "OP CERTIFICADO " + gvHome.GetSelectedFieldValues("año")[0].ToString() + gvHome.GetSelectedFieldValues("mes")[0].ToString() + " " + gvHome.GetSelectedFieldValues("IdentifEspacio")[0].ToString() + " - " + gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString(); 

                        string filename = Pauta + System.DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                                  System.DateTime.Now.Minute.ToString().PadLeft(2, '0') +
                                                  System.DateTime.Now.Second.ToString().PadLeft(2, '0') + ".xlsx";

                        TextBox1.Text = Server.MapPath("~/Excel/");

                        switch (TextBox1.Text)
                        {
                            case "": { break; }

                            default:
                                {
                                    try
                                    {
                                        System.IO.FileInfo fileinfo = new FileInfo(@TextBox1.Text);
                                        bool dExiste = fileinfo.Directory.Exists;

                                        if (dExiste)
                                        {
                                            lblErrorLineas.Text = "";

                                            string[] dirs = Directory.GetFiles(@TextBox1.Text);

                                            if (dirs.Length > 0)
                                            {
                                                filename = @TextBox1.Text + filename;

                                                dExiste = false;

                                                for (int i = 0; i <= dirs.Length - 1; i++)
                                                {
                                                    if (dirs[i].ToString() == filename)
                                                    {
                                                        dExiste = true;

                                                        break;
                                                    }
                                                }
                                            }

                                            if (dExiste == false)
                                            {
                                                //ACA VA LA LLAMADA A LA FUNCION DEL HELPER
                                                // ES EL CORE, YA TENEMOS EL NOMBRE DE ARCHIVO Y LA RUTA

                                                CertificadoCabDTO Cabecera      = Certificados.Read( gvHome.GetSelectedFieldValues("PautaId")[0].ToString(),gvHome.GetSelectedFieldValues("IdentifOrigen")[0].ToString());
                                                List<CertificadoDetDTO> Detalle = Certificados.ReadAllLineas(Cabecera);
                                                List<CertificadoSKUDTO> SKUS    = Certificados.GetSKUs(Cabecera.PautaId,Cabecera.IdentifOrigen);
                                                EspacioContDTO Espacio          = CRUDHelper.Read(string.Format("IdentifEspacio = '{0}'", Cabecera.IdentifEspacio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

                                                csOP_Helper Helper = new csOP_Helper("CERTIFICADO", "", CertificadoCab.PautaId,Cabecera, Detalle, SKUS, Espacio,filename);

                                                System.IO.FileInfo toDownload = new FileInfo(filename);

                                                if (toDownload.Exists == true)
                                                {
                                                    Response.Clear();
                                                    Response.AddHeader("Content-Disposition", "attachment; filename=" + toDownload.Name);
                                                    Response.AddHeader("Content-Length", toDownload.Length.ToString());
                                                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
                                                    Response.WriteFile(filename);
                                                    Response.End();
                                                }

                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        string a = string.Empty;

                                        a = ex.Message;

                                        lblErrorLineas.Text = "Ruta no válida.";
                                    }

                                    break;
                                }
                        }
                    }
                    else
                    {
                        lblErrorHome.Text = "Debe seleccionar una línea";
                    }

                    break;

                case "btnCost":

                    Business.Certificados.CalcularCosto(CertificadoCab, Costos, mycert, ((Accendo)this.Master).Usuario.UserName);

                    lblErrorLineas.Text = "Se calculo correctamente";

                    break;

                case "btnExport":

                case "btnExportXls":
                    {
                        if (ASPxGridViewExporter1 != null)
                       {
                            XlsExportOptions xlsExportOptions = new XlsExportOptions(TextExportMode.Text, true, false);

                            ASPxGridViewExporter1.WriteXlsToResponse(xlsExportOptions);
                        }
                        break;
                    }
                case "btnExportPdf":

                    if (ASPxGridViewExporter1 != null)

                        ASPxGridViewExporter1.WritePdfToResponse();

                    break;

                default: break;
            }

        }

        protected void ASPxGridViewExporter1_RenderBrick(object sender, ASPxGridViewExportRenderingEventArgs e)
        {
            if (e.RowType == DevExpress.Web.ASPxGridView.GridViewRowType.Data && e.Column != null)
            {
                GridViewDataColumn dataColumn = e.Column as GridViewDataColumn;

                if (dataColumn.FieldName.ToUpper() == "HORA")
                {
                    DateTime dt1 = Convert.ToDateTime(e.Value.ToString());

                    e.Text = string.Format("{0:HH:mm}", dt1);
                }

                if (dataColumn.FieldName.ToUpper() == "FECHA")
                {
                    DateTime dt = Convert.ToDateTime(e.Value.ToString());

                    e.Text = string.Format("{0:dd/MM/yyyy}", dt);
                }
            }
        }

        #region "Botón Volver"
        protected void btnBack_Click(object sender, EventArgs e)
        {
            Back();
        }

        private void Back()
        {
            Response.Redirect("Certificado.aspx");

            VaciarCamposPauta();
            VaciarCamposDetalle();
            gvHome.Selection.UnselectAll();
            ASPxGridViewExporter1.GridViewID = "gvHome";
        }
        #endregion

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            Back();
        }

        protected void btnCancelSKU_Click(object sender, EventArgs e)
        {
            trQuerySKU.Visible = false;
        }

        protected void teHoraFinInsertar_ValueChanged(object sender, EventArgs e)
        {
            spSalidasInsertar.Enabled = (teHoraInicio.Text != teHoraFinInsertar.Text && teHoraFinInsertar.Text == "00:00");
        }

        protected void teHoraFinInsertar_DateChanged(object sender, EventArgs e){}

        /// <summary>
        ///  PROPIO DE CERTIFICADO
        /// </summary>
        /// <param name="horaInicio"></param>
        /// <param name="horaFin"></param>
        /// <param name="intervalo"></param>
        /// <param name="frecuenciaCab"></param>
        /// <param name="aviso"></param>
        /// <param name="frecuenciaDetalles"></param>
        /// <returns></returns>
        protected bool Validaciones(TimeSpan horaInicio, TimeSpan horaFin, IntervaloDTO intervalo, FrecuenciaDTO frecuenciaCab, AvisosDTO aviso, List<FrecuenciaDetDTO> frecuenciaDetalles)
        {
            bool retval = false;

            decimal CantMinutosAviso = 0;

            List<CertificadoDetDTO> lineas = mycert;

            List<CertificadoDetDTO> preExistentes = new List<CertificadoDetDTO>();

            //TODO: VALIDACION 1: Validar que la hora de inicio de la inserción, no sea superior a la hora fin de la pauta.
            if (Convert.ToDateTime(teHoraInicioInsertar.Text) > Convert.ToDateTime(teHoraFin.Text) || Convert.ToDateTime(teHoraInicioInsertar.Text) < Convert.ToDateTime(teHoraInicio.Text))
            {
                lblErrorLineas.Text = "Las horas ingresadas no están dentro del rango de la cabecera.";

                retval = false;

                return retval;
            }

            //TODO: VALIDACION 2: La duración del aviso excede el intervalo seleccionado para la Pauta. NO se generaron líneas.
            if (intervalo == null)
            {
                intervalo = new IntervaloDTO();

                intervalo.CantMinutos = Convert.ToDecimal(0);
            }

            if (Convert.ToInt32(spDuracionInsertar.Value) != 0)
            {
                CantMinutosAviso = Convert.ToDecimal(spDuracionInsertar.Value) / 60;

                if (CantMinutosAviso > intervalo.CantMinutos)
                {
                    if (spSalidasInsertar.Enabled == false)
                    {
                        lblErrorLineas.Text = "La duración del aviso excede el intervalo seleccionado para la Pauta. NO se generaron líneas.";

                        retval = false;

                        return retval;
                    }
                }
            }

            //TODO: VALIDACION 3: Si no hay lineas preexistentes,debe devolver verdadero..
            if (lineas.Count == 0)
            {
                retval = true;

                return retval;
            }

            //TODO: VALIDACION 4: Para aviso con duracion cero segundos y salidas deshabilitada.... (GRAFICA, PNT, ETC)
            if (Convert.ToInt32(spDuracionInsertar.Value) == 0 && spSalidasInsertar.Enabled == false)
            {
                    lblErrorLineas.Text = "No ingrese avisos con duracion cero segundos si salida está deshabilitada.";

                    retval = false;
            }

            lineas = lineas.OrderBy(p => p.Dia).ThenBy(p => p.Hora).ToList(); //Ordena la lista por dia y luego por hora
          
            //TODO: VALIDACION 5: Comparo que la linea insertada no pise valores preexistentes
            for (int j = 0; j <= lineas.Count - 1; j++)
            {
                //Hora de inicio dentro del array de preexistentes
                TimeSpan ts1i = new TimeSpan(0, lineas[j].Hora.Hours, lineas[j].Hora.Minutes, 0);

                //Duracion del aviso dentro del array de preexistentes
                TimeSpan ts1d = new TimeSpan(0, 0, Convert.ToInt32(lineas[j].Duracion));

                //Hora de finalizacion del aviso dentro del array de preexistentes midiendo su duracion
                TimeSpan ts1f = ts1i.Add(ts1d);

                //Hora de inicio dentro de la linea a insertar
                TimeSpan ts2i = new TimeSpan(0, Convert.ToDateTime(teHoraInicioInsertar.Text).Hour, Convert.ToDateTime(teHoraInicioInsertar.Text).Minute, 0);

                //Duracion del aviso dentro de la linea a insertar
                TimeSpan ts2d = new TimeSpan(0, 0, Convert.ToInt32(aviso.Duracion));

                //Hora de finalizacion del aviso dentro de la linea a insertar midiendo su duracion
                TimeSpan ts2f = ts2i.Add(ts2d);

                //TODO: VALIDACION 6: Si es la misma hora de inicio del aviso preexistente, sale con error
                if (ts2i.CompareTo(ts1i) == 0)
                {
                    lblErrorLineas.Text = "El aviso se superpone con otro";

                    retval = false;

                    break;
                }
                
                //Verificacion contra el registro siguiente
                //Quiere decir que hay un registro previo despues del actual
                if (j < lineas.Count)
                {
                    TimeSpan tsli                   = new TimeSpan(0, lineas[j + 1].Hora.Hours, lineas[j + 1].Hora.Minutes, 0);
                    TimeSpan tsld                   = new TimeSpan(0, 0, Convert.ToInt32(lineas[j + 1].Duracion));
                    TimeSpan tslf                   = tsli.Add(tsld);
                    TimeSpan RangoInicio            = ts1i;
                    TimeSpan RangoFinal             = ts1f;
                    TimeSpan HoraInicioAvisoNuevo   = ts2i;
                    TimeSpan HoraFinAvisoNuevo      = ts2f;

                    if (HoraFinAvisoNuevo.CompareTo(tsli) == 1)
                    {
                        lblErrorLineas.Text = "El aviso se superpone con otro";

                        retval = false;

                        break;
                    }

                }

                if (ts1i.Days == ts2i.Days)
                {
                    TimeSpan RangoInicio          = ts1i;
                    TimeSpan RangoFinal           = ts1f;
                    TimeSpan HoraInicioAvisoNuevo = ts2i;
                    TimeSpan HoraFinAvisoNuevo    = ts2f;

                    if (HoraInicioAvisoNuevo.CompareTo(RangoInicio) == 1 && HoraInicioAvisoNuevo.CompareTo(RangoFinal) == 2)
                    {
                        lblErrorLineas.Text = "El aviso se superpone con otro";

                        retval = false;

                        break;
                    }

                    if (ts2i.CompareTo(ts1i) == 1)
                    {
                        //Comparo que la hora de inicio del aviso sea mayor que la hora de finalizacion del preexistente
                        if (ts1f <= ts2i && aviso.Duracion != 0)
                        {
                            retval = true;

                            break;
                        }
                        else
                        {
                            lblErrorLineas.Text = "El aviso se superpone con otro";

                            retval = false;
                        }
                    }
                }
                else
                {
                    retval = false;
                }
            }
            return retval;
        }

        protected void spAvisoReempDuracion_NumberChanged(object sender, EventArgs e)
        {

        }

        protected void btnBuscarEspacioPeriodo_Click(object sender, EventArgs e)
        {
            botonbuscarperiodo = "true";
            try
            {
                string sEncontrado = Certificados.BuscarCertificado((string)ucIdentifEspacio.SelectedValue, deAnoMes.Date.Year.ToString(), deAnoMes.Date.Month.ToString().PadLeft(2,'0'), "", Convert.ToString(ucIdentifOrigen1.SelectedValue));

                List<CertificadoCabDTO> miPautaId = null;

                switch(sEncontrado)
                {
                    case "Certificado NULL":
                         string jscode = "function Devolver() { __doPostBack('nula', '')  } Devolver();";
                         ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "ajax", jscode, true);
                         break;

                    case "Origen NULL":
                         miPautaId = Certificados.ReadAll("IDENTIFORIGEN IS NULL AND IDENTIFESPACIO = '" + (string)ucIdentifEspacio.SelectedValue + "' AND ANOMES = '" + deAnoMes.Date.Year.ToString() + deAnoMes.Date.Month.ToString().PadLeft(2,'0') + "'");
                         Certificados.Crear(miPautaId[0].PautaId, ucIdentifOrigen1.SelectedValue.ToString());
                         CargarCertificado();

                          gv.SortBy(gv.Columns["Dia"], DevExpress.Data.ColumnSortOrder.Ascending);
                          gv.SortBy(gv.Columns["Hora"], -1);
                          gv.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);
                         break;

                    case "Certificado OK":
                         miPautaId = Certificados.ReadAll("IDENTIFORIGEN = '" + ucIdentifOrigen1.SelectedValue.ToString() + "' AND IDENTIFESPACIO = '" + (string)ucIdentifEspacio.SelectedValue + "' AND ANOMES = '" + deAnoMes.Date.Year.ToString() + deAnoMes.Date.Month.ToString().PadLeft(2, '0') + "'");
                         CargarCertificado();

                          gv.SortBy(gv.Columns["Dia"], DevExpress.Data.ColumnSortOrder.Ascending);
                          gv.SortBy(gv.Columns["Hora"], -1);
                          gv.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);

                          break;
                }

                botonbuscarperiodo = "false";
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }
        }

        protected void nula()
        { 
        }

        protected void btnBuscarPauta_Click(object sender, EventArgs e)
        {
            botonbuscarpauta = "true";

            try
            {
                string sEncontrado = Certificados.BuscarCertificado((string)ucIdentifEspacio.SelectedValue, deAnoMes.Date.Year.ToString(), deAnoMes.Date.Month.ToString().PadLeft(2,'0'), txNroPauta.Text, Convert.ToString(ucIdentifOrigen2.SelectedValue));

                List<CertificadoCabDTO> miPautaId = null;

                switch (sEncontrado)
                {
                    case "Certificado NULL":
                        string jscode = "function Devolver() { __doPostBack('nula', '')  } Devolver();";
                        ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "ajax", jscode, true);
                        break;

                    case "Origen NULL":

                        if (botonbuscarpauta == "true")
                        {
                            miPautaId = Certificados.ReadAll("IDENTIFORIGEN IS NULL AND PAUTAID = '" + txNroPauta.Text + "'");
                            Certificados.Crear(miPautaId[0].PautaId, ucIdentifOrigen2.SelectedValue.ToString());
                        }
                        else
                        {
                            miPautaId = Certificados.ReadAll("IDENTIFORIGEN IS NULL AND IDENTIFESPACIO = '" + (string)ucIdentifEspacio.SelectedValue + "' AND ANOMES = '" + deAnoMes.Date.Year.ToString() + deAnoMes.Date.Month.ToString().PadLeft(2,'0') + "'");
                            Certificados.Crear(miPautaId[0].PautaId, ucIdentifOrigen1.SelectedValue.ToString());
                        }

                        CargarCertificado();
                        gv.SortBy(gv.Columns["Dia"], DevExpress.Data.ColumnSortOrder.Ascending);
                        gv.SortBy(gv.Columns["Hora"], -1);
                        gv.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);
                        break;

                    case "Certificado OK":
                        miPautaId = Certificados.ReadAll("IDENTIFORIGEN = '" + ucIdentifOrigen2.SelectedValue.ToString() + "' AND PAUTAID = '" + txNroPauta.Text + "'");
                        CargarCertificado();
                        gv.SortBy(gv.Columns["Dia"], DevExpress.Data.ColumnSortOrder.Ascending);
                        gv.SortBy(gv.Columns["Hora"], -1);
                        gv.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);
                        break;
                }

                botonbuscarpauta = "false";
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsg, ex);
            }

        }

        protected void gvHome_RowDblClick()
        { 
        }
        protected void btnValidar_Click(object sender, EventArgs e)
        {
            //Si llegó hasta acá es porque quiero validar el certificado actual.

            Business.Certificados.CalcularCosto(CertificadoCab,Costos,mycert, ((Accendo)this.Master).Usuario.UserName);
            this.txUsuCierre.Text  = ((Accendo)this.Master).Usuario.UserName;;
            this.deFecCierre.Value = System.DateTime.Now;

            CertificadoCab.FecCertValido = Convert.ToDateTime(deFecCierre.Value);
            CertificadoCab.CertValido    =  ((Accendo)this.Master).Usuario.UserName;;

            RecargarDiaHora();

            spEspacio.Value           = GetEspacioContenido().IdentifEspacio;
            spMedio.Value             = GetEspacioContenido().IdentifMedio;
            spAnioMes.Value           = CertificadoCab.AnoMes.ToString().Substring(0, 4) + "-" + CertificadoCab.AnoMes.ToString().Substring(4);
            spOrigenCertificado.Value = Convert.ToString((gvHome.GetSelectedFieldValues(new string[] { "IdentifOrigen" })[0]));
            
        }
    }
}