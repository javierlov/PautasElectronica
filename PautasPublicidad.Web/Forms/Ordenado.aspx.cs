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
using PautasPublicidad.Web.Classes;
using System.IO;
using System.Threading;
using ClosedXML.Excel;

namespace PautasPublicidad.Web.Forms
{
    public partial class Ordenado : System.Web.UI.Page
    {
        int ProxRecId = 0;

        OrdenadoCabDTO ordenado;

        #region ViewState

         public List<OrdenadoDetDTO> Lineas
        {
             get
            {
                if (Session["Ordenado.Lineas" + Session.SessionID] != null && Session["Ordenado.Lineas" + Session.SessionID] is List<OrdenadoDetDTO>)
                    return Session["Ordenado.Lineas" + Session.SessionID] as List<OrdenadoDetDTO>;
                else
                    return new List<OrdenadoDetDTO>();
            }

             set
            {
                Session.Add("Ordenado.Lineas" + Session.SessionID, value);
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

        public OrdenadoCabDTO OrdenadoCab
        {
            get
            {
                if (ViewState["OrdenadoCab"] != null && ViewState["OrdenadoCab"] is OrdenadoCabDTO)
                    return ViewState["OrdenadoCab"] as OrdenadoCabDTO;
                else
                    return new OrdenadoCabDTO();
            }
            set
            {
                ViewState.Add("OrdenadoCab", value);
            }
        }

        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            btnSave.Enabled = true;

            if (!Page.IsPostBack && !Page.IsCallback)
            {
                Lineas                                              = null; //Limpio las lineas de mi session.
                spSalidasInsertar.Enabled                           = true;
                deAnoMes.Date                                       = DateTime.Now;
                trButtons.Visible                                   = true;
                trFind.Visible                                      = true;
                trPauta.Visible                                     = false;
                trAccion.Visible                                    = false;
                trEditLine.Visible                                  = false;
                trQuerySKU.Visible                                  = false;
                opEditPeriodo.Checked                               = true;
                tblDelete.Visible                                   = false;

                FormsHelper.InicializarPropsGrilla(gv);
                
                gv.SettingsEditing.Mode                             = GridViewEditingMode.Inline;
                gv.SettingsBehavior.AllowSelectByRowClick           = true;
                gv.SettingsBehavior.AllowSelectSingleRowOnly        = false;
                
                FormsHelper.InicializarPropsGrilla(gvHome);
                
                gvHome.SettingsEditing.Mode                         = GridViewEditingMode.Inline;
                gvHome.SettingsBehavior.AllowSelectByRowClick       = true;
                gvHome.SettingsBehavior.AllowSelectSingleRowOnly    = false;

                ASPxGridViewExporter1.GridViewID                    = "gvHome";

                var ord = Ordenados.ReadAll("PautaId IS NOT NULL");

            }

            GridViewDataComboBoxColumn gvc = gv.Columns["IdentifAviso"] as GridViewDataComboBoxColumn;
            gvc.Name                            = "IdentifAviso";
            gvc.Caption                         = "Aviso";
            gvc.FieldName                       = "IdentifAviso";
            gvc.PropertiesComboBox.TextField    = "Name"; //mapInfo.EntityTextField;
            gvc.PropertiesComboBox.ValueField   = "IdentifAviso"; // mapInfo.EntityValueField;
            gvc.PropertiesComboBox.DataSource   = Business.Avisos.ReadAll(""); //mapInfo.DAOHandler.ReadAll("");

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

            RefreshAbmGridUnsorted(gv);

            if (trQuerySKU.Visible){RefreshSKUGrid(gvSKU);}

            gvHome.DataSource = Ordenados.VistaOrdenados();
            gvHome.DataBind();
       
        }
        
        private void SortDDL(ref DropDownList ddl)
        {
            ListItem[] items = new ListItem[ddl.Items.Count];

            ddl.Items.CopyTo(items, 0);
            ddl.Items.Clear();

            Array.Sort(items, (x, y) => { return x.Text.CompareTo(y.Text); });

            ddl.Items.AddRange(items);
        } 

        private void RefreshSKUGrid(ASPxGridView gvSKU)
        {
            decimal total = 0;
            var dt = Ordenados.BuildAllSKU(Lineas);
            
            gvSKU.DataSource = dt;
            gvSKU.DataBind();

            foreach (System.Data.DataRow dr in dt.Rows)
                total += Convert.ToDecimal(dr["CantSalidas"]);
            
            lblSKUTotalSalidas.Text = "Total de Salidas: " + total.ToString();
        }

        private void RefreshHomeGrid(ASPxGridView gvHome)
        {
            //cargar con pautas 
            gvHome.DataSource = Lineas;
            gvHome.DataBind();

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
        void AvisoDestinoReemplazar_SelectedIndexChanged(object sender, EventArgs e)
        {
            AvisoDestinoReemplazarChanged();
        }

        void IdentifFrecuencia_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucFrecuenciaChanged();
        }

        void IdentifEspacio_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucEspacioChanged();
        }

        void ddlNroPauta_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucNumPautaChanged();

        }
        
        private EspacioContDTO GetEspacioContenido()
        {
            if (ucIdentifEspacio.SelectedValue != null)
            {
                return CRUDHelper.Read(string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
            }
            else
                return null;
        }

        private EspacioContDTO GetEspacioContenidoEnPauta()
        {
            if (ddlNroPauta.SelectedValue != null)
            {
                return CRUDHelper.Read(string.Format("RecId = '{0}'", ddlNroPauta.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
            }
            else
                return null;
        }
        
        private void ucNumPautaChanged()
        {
            OrdenadoCabDTO ord = Ordenados.Read(Convert.ToInt32(ddlNroPauta.SelectedValue));
            if (ord != null) //funcion para completar Medio
            {
                ucIdentifEspacio.SelectedValue = ord.IdentifEspacio;

                DateTime d = new DateTime(Convert.ToInt32(ord.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(ord.AnoMes.ToString().Substring(4, 2)), 01);

                deAnoMes.Value = d;
                ucEspacioChanged();
            }
            else
            {
                ucIdentifEspacio.SelectedValue = null;
                txMedio.Text = "";
                deAnoMes.Value = new DateTime();
            }
        }
        
        private void ucEspacioChanged()
        {
            EspacioContDTO espacio = GetEspacioContenido();

            if (espacio != null)
            {
                txMedio.Text = espacio.IdentifMedio;
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
            else
            {
                txMedio.Text = string.Empty;
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
            
                string msg = string.Empty;
                
                if (!ValidarFranjaHoraria(ref msg))
                    lblErrorLineas.Text = msg;
            }
        }
        
        private void AvisoDestinoReemplazarChanged()
        {
            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoDestinoReemplazar.SelectedValue));

            if (aviso != null)
            {
                spAvisoReempDuracion.Value = aviso.Duracion;

                string msg = string.Empty;
                
                if (!ValidarFranjaHoraria(ref msg))
                    lblErrorLineas.Text = msg;
                
            }
        }
        
        protected void btnRefresh_Click(object sender, ImageClickEventArgs e)
        {
            lblErrorLineas.Text = string.Empty;

            Labelx.Text = lblErrorLineas.Text;

            string setup = Ordenados.AnoMesCierreOrd();

            decimal anomes = Convert.ToDecimal(deAnoMes.Date.ToString("yyyyMM"));
            
            if (ucIdentifEspacio.SelectedValue != null)
            {
                if (deAnoMes.Value != null)
                {
                    if (anomes > Convert.ToDecimal(setup))
                        CargarOrdenado();
                    else
                        lblValidaAñoMes.Text = "No puede crear ordenado en el periodo seleccionado. Ya se encuentra cerrado.";
                }
                else
                {
                    lblValidaAñoMes.Text = "Debe ingresar una fecha";
                }
            }
            else
            {
                lblValidaAñoMes.Text = "Debe seleccionar un Espacio de contenido para continuar";
            }
                        
        }

        private void CargarPautas()
        {
            RefreshHomeGrid(gvHome); 
        }

        private bool ValidarFranjaHoraria(ref string mjeErr)
        {
            bool retVal = false;

            try
            {
                if (teHoraInicio.Text == teHoraFin.Text && teHoraInicio.Text == "")
                    return true;

                if (Convert.ToDateTime(teHoraInicio.Text) > Convert.ToDateTime(teHoraFin.Text))
                   mjeErr = ("La hora de inicio no puede ser mayor a la final.");
                else
                     retVal = true;

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

        private void CargarOrdenado()
        {
        
            lblValidaAñoMes.Text = string.Empty;
            lblErrorHome.Text = string.Empty;

            ordenado = Business.Ordenados.Read(Convert.ToString(ucIdentifEspacio.SelectedValue), deAnoMes.Date.Year, deAnoMes.Date.Month);
            
            ASPxPageControl1.Visible = true;
            trPauta.Visible = true;
            trFind.Visible = false;

            EspacioContDTO espacio = GetEspacioContenido();
            
            if (ordenado != null)
                ReCargarControles(ordenado);

            RecargarDiaHora();
        }

        private void RecargarDiaHora() 
        {
            EspacioContDTO espacio = GetEspacioContenido();
            
            OrdenadoCabDTO ordenado = Business.Ordenados.Read(Convert.ToString(ucIdentifEspacio.SelectedValue), deAnoMes.Date.Year, deAnoMes.Date.Month);

            CostosDTO costos = Business.Ordenados.FindCosto(Convert.ToString(ucIdentifEspacio.SelectedValue), deAnoMes.Date.Year, deAnoMes.Date.Month);

            if (ordenado != null)
            {

                //GenerarLineas(teHoraInicio.DateTime, teHoraFin.DateTime);
                OrdenadoCab = ordenado;
                Lineas = Ordenados.ReadAllLineas(ordenado);

                Costos = costos;
                //RefreshHomeGrid(gvHome);
                RefreshAbmGrid(gv);

                /// si el ordenado esta cerrado oculto paneles de edicion. 
                if (OrdenadoCab.UsuCierre != "" && OrdenadoCab.UsuCierre != null)
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
                    deHoraDesdeOrigenReemplazar.DateTime = ordenado.VigDesde.Date;
                
                if (espacio.HoraFin != null)
                    deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraFin.Value);
                else if(ordenado.HoraFin != null)
                    deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(ordenado.HoraFin);
                else 
                    deHoraHastaOrigenReemplazar.DateTime = ordenado.VigHasta.Date.AddHours(23.9999);//origen

                deFechaDesdeDestinoCopiar.Date = deFechaDesdeOrigenCopiar.Date;
                deFechaHastaOrigenCopiar.Date = deFechaHastaOrigenCopiar.Date;
                
            }
            else
            {
                if (costos != null)
                {
                    Costos = costos;
                    //Cargo controles.
                    spPautaID.Number = Business.Ordenados.GetNextPautaId();

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

                    litCambiarPauta.Text = string.Format("Espacio: {0} | Período: {1}", ucIdentifEspacio.SelectedText, deAnoMes.Date.ToString("yyyy-MM"));

                    //Actualizo controles.
                    ucFrecuenciaChanged();

                    ASPxPageControl1.Visible = true;
                    trPauta.Visible = true;
                    trFind.Visible = false;

                    //Inicializo Fechas origen para copia
                    deFechaDesdeOrigenCopiar.Date = costos.VigDesde;
                    deFechaHastaOrigenCopiar.Date = costos.VigHasta;
                    deFechaDesdeReemplazar.Date   = costos.VigDesde;
                    deFechaHastaReemplazar.Date   = costos.VigHasta;

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

                    else if (ordenado != null)
                        if(ordenado.HoraFin != null)
                           deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(ordenado.HoraFin);

                    //fecha inicio de destino
                    deFechaDesdeDestinoCopiar.Date = costos.VigDesde; //destino
                    
                    ordenado = new OrdenadoCabDTO();
                    ordenado.AnoMes            = Convert.ToInt32(deAnoMes.Date.Year.ToString() + 
                                                                 deAnoMes.Date.Month.ToString("00"));

                    ordenado.CantSalidas       = 0;                                                     //Cantidad de salidas total
                    ordenado.Costo             = 0;                                                     //Costo Total de la Pauta
                    ordenado.CostoOp           = 0;                                                     //Costo de la Pauta para la Orden de Publicidad
                    ordenado.CostoOpUni        = 0;                                                     //Costo de la Pauta para la Orden de Publicidad por unidad (segundos, página)
                    ordenado.CostoUni          = 0;                                                     //Costo Total por unidad (segundos, página)
                    ordenado.DatareaId         = 0;                                                     //Area de Trabajo
                    ordenado.DuracionTot       = 0;                                                     //Total Duración o Cantidad
                    ordenado.FecCierre         = null;                                                  //Fecha del cierre
                    ordenado.FecCosto          = DateTime.Now;                                          //Fecha en qué se calculo el costo por última vez
                    ordenado.HoraInicio        = FormsHelper.ConvertToTimeSpan(teHoraInicio.DateTime);  //Hora de Inicio
                    ordenado.HoraFin           = FormsHelper.ConvertToTimeSpan(teHoraFin.DateTime);     //Hora de Finalizacion
                    ordenado.IdentifEspacio    = Convert.ToString(ucIdentifEspacio.SelectedValue);      // Espacio de Contenido
                    ordenado.IdentifFrecuencia = Convert.ToString(ucIdentifFrecuencia.SelectedValue);   // Frecuencia
                    ordenado.IdentifIntervalo  = Convert.ToString(ucIdentifIntervalo.SelectedValue);    // Intervalo
                    ordenado.PautaId           = Convert.ToString(spPautaID.Number);                    // Pauta
                    ordenado.RecId             = 0;                                                     // Registro #
                    ordenado.UsuCierre         = "";                                                    //Usuario que cerró
                    ordenado.UsuCosto          = costos.Confirmado;                                     //Usuario qué calculo el costo
                    ordenado.VersionCosto      = costos.Version.Value;                                  //Versión del registro
                    ordenado.VigDesde          = costos.VigDesde;                                       //Fecha desde la cual está vigente el Costo
                    ordenado.VigHasta          = costos.VigHasta;                                       //Fecha hasta la cual estará vigente el Costo

                    OrdenadoCab = ordenado;

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
                    lblValidaAñoMes.Text = "No existen costos Confirmados para este Espacio – Vigencia.";
                    ASPxPageControl1.Visible = true;
                    trPauta.Visible = false;
                    trFind.Visible = true;

                }
            }
        }

        private bool ValidarSalida()
        {
            bool bResult = false;

            if (teHoraInicio.Text == "00:00" && teHoraFin.Text == "00:00")
            {
                if( spSalidasInsertar.Text == "0")
                {  
                    lblErrorLineas.Text = "Debe ingresar un numero de salida mayor que cero";
                }
            }
            else
            {
                bResult = true;
            }

            return bResult;
        }

        private bool ValidarIntervalo(ref string mjeErr)
        {
            bool bresul = false;

            try
            {
                if (teHoraFin.DateTime != teHoraInicio.DateTime && ucIdentifIntervalo.SelectedValue == null)
                {
                    mjeErr = ("Debe seleccionar un Intervalo");
                }
                else
                {
                    bresul = true;
                }

                if (!bresul)
                {
                    ASPxPageControl1.Visible    = true;
                    trPauta.Visible             = true;
                    trFind.Visible              = false;
                }
            }
            catch (Exception ex)
            {
                ASPxPageControl1.Visible        = true;
                trPauta.Visible                 = true;
                trFind.Visible                  = false;

                MsgErrorLinas(ex);
                mjeErr                          = ex.Message;
            }

            return bresul;
        }

        private bool ValidarFrecuencia(ref string mjeErr)
        {
            bool bresul = false;

          try
            {
              if (ucIdentifFrecuencia.SelectedValue == null)
                  mjeErr = ("Debe seleccionar una Frecuencia.");
              else
                  bresul = true;

              if (!bresul)
                {
                    ASPxPageControl1.Visible = true;
                    trPauta.Visible          = true;
                    trFind.Visible           = false;
                }
          
            }
            catch (Exception ex)
            {
                ASPxPageControl1.Visible    = true;
                trPauta.Visible             = true;
                trFind.Visible              = false;

                MsgErrorLinas(ex);
                mjeErr                      = ex.Message;

            } return bresul;
        }
        
        private void VaciarCamposPauta()
        { 
            spPautaID.Value                     = "";
            ucIdentifFrecuencia.SelectedValue   = null;
            teHoraInicio.Value                  = "";
            teHoraFin.Value                     = "";
            ucIdentifIntervalo.SelectedValue    = null;
            spVersionCosto.Value                = "";
            txUsuCosto.Value                    = "";
            txUsuCierre.Value                   = "";
            deFecCosto.Value                    = "";
            deFecCierre.Value                   = "";
            spCantSalidas.Value                 = "";
            
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

        private void ReCargarControles(OrdenadoCabDTO ordenado)
        {
            EspacioContDTO espacio  = GetEspacioContenido();
            lblErrorLineas.Text     = string.Empty;
            Labelx.Text             = string.Empty;
            
            //Controles de la pauta.
            spPautaID.Number                    = Convert.ToInt32(ordenado.PautaId);
            ucIdentifFrecuencia.SelectedValue   = ordenado.IdentifFrecuencia;
            ucIdentifIntervalo.SelectedValue    = ordenado.IdentifIntervalo;

            if (FormsHelper.ConvertToDateTime(ordenado.HoraInicio) != Convert.ToDateTime("00:00:00"))
                teHoraInicio.DateTime = FormsHelper.ConvertToDateTime(ordenado.HoraInicio);

            teHoraInicioInsertar.DateTime = teHoraInicio.DateTime;

            if (FormsHelper.ConvertToDateTime(ordenado.HoraFin) != Convert.ToDateTime("00:00:00"))
                teHoraFin.DateTime = FormsHelper.ConvertToDateTime(ordenado.HoraFin);

            teHoraFinInsertar.DateTime = teHoraFin.DateTime;

            //Controles de solo lectura.
            spVersionCosto.Value    = ordenado.VersionCosto;
            txUsuCosto.Text         = ordenado.UsuCosto;
            deFecCosto.Date         = ordenado.FecCosto;
            deFecCosto.Value        = ordenado.FecCosto;
            
            if(ordenado.FecCierre > Convert.ToDateTime("01/01/1900"))
            {
                txUsuCierre.Text    = ordenado.UsuCierre;
                deFecCierre.Value   =  ordenado.FecCierre;
            }
            else
            {
                txUsuCierre.Text    = "";
                deFecCierre.Value   = "";
                ordenado.UsuCierre  = "";
                ordenado.FecCierre  = null;
            }

            spCantSalidas.Value = ordenado.CantSalidas;

            litCambiarPauta.Text = string.Format("Espacio: {0} | Período: {1}", ucIdentifEspacio.SelectedText, deAnoMes.Date.ToString("yyyy-MM"));

            //Actualizo controles.
            ucFrecuenciaChanged();

            ASPxPageControl1.Visible = true;
            trPauta.Visible = true;
            trFind.Visible = false;

            //Solo se pueden modificar ordenados NO cerrados.
            btnSave.Enabled = (ordenado.FecCierre == null);
            
            //Inicializo Fechas...
            
            deHoraDesdeOrigenReemplazar.DateTime  = FormsHelper.ConvertToDateTime(ordenado.HoraInicio);
            deHoraHastaOrigenReemplazar.DateTime = FormsHelper.ConvertToDateTime(ordenado.HoraFin);

            DateTime date   = new DateTime();
            date            = Convert.ToDateTime(deAnoMes.Date);

            deFechaDesdeDestinoCopiar.Date = new DateTime(date.Year, date.Month, 1);
            deFechaDesdeOrigenCopiar.Date  = new DateTime(date.Year, date.Month, 1);
            deFechaHastaOrigenCopiar.Date  = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
            deFechaHastaDestinoCopiar.Date = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
            deFechaDesdeReemplazar.Date    = new DateTime(date.Year, date.Month, 1);
            deFechaHastaReemplazar.Date    = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
        }

        protected void btnGenerarLineas_Click(object sender, EventArgs e)
        {
            try
            {
                
                string msg          = string.Empty;
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
		        IntervaloDTO intervalo                      = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));
                FrecuenciaDTO frecuencia                    = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                List<FrecuenciaDetDTO> frecuenciaDetalles   = CRUDHelper.ReadAll( string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAviso.SelectedValue)); 

                //Genero las nuevas líneas.
                Lineas = GenerarLineas(FormsHelper.ConvertToTimeSpan(horaInicio), FormsHelper.ConvertToTimeSpan(horaFin), intervalo, frecuencia, aviso, frecuenciaDetalles);

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

        private List<OrdenadoDetDTO> GenerarLineas(TimeSpan horaInicio, TimeSpan horaFin, IntervaloDTO intervalo, FrecuenciaDTO frecuenciaCab, AvisosDTO aviso, List<FrecuenciaDetDTO> frecuenciaDetalles)
        {
            List<OrdenadoDetDTO> lineas         = Lineas;
            List<OrdenadoDetDTO> preExistentes  = new List<OrdenadoDetDTO>();

            if (intervalo == null)
            {
                intervalo             = new IntervaloDTO();
                intervalo.CantMinutos = Convert.ToDecimal(0);
            }

            if(Validaciones(horaInicio,horaFin,intervalo,frecuenciaCab,aviso,frecuenciaDetalles) == false)
            {
                RefreshAbmGrid(gv);
                return lineas;
            }
            
            TimeSpan incremento = TimeSpan.FromMinutes(Convert.ToDouble(intervalo.CantMinutos));

            TimeSpan horaTemp;
            OrdenadoDetDTO linea;
            List<DateTime> periodo;

            try
            {
                //Obtengo la lista de los días partiedo de fechas o nombres de dia.
                if (frecuenciaCab.SemMes == "SEMANA")
                    periodo = Ordenados.GetDatesByDayNames(deAnoMes.Date.Year, deAnoMes.Date.Month, GetDiasSeleccionados());

                else
                    periodo = Ordenados.GetDatesByDayNumbers( deAnoMes.Date.Year, deAnoMes.Date.Month, GetDiasSeleccionados());

                foreach (DateTime fecha in periodo)
                {
                    horaTemp = horaInicio;

                    if (aviso.Duracion == 0)
                        incremento = new TimeSpan(0);
                    else if ( aviso.Duracion==null)
                        incremento = new TimeSpan(0);


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

                        linea           = new OrdenadoDetDTO();
                        linea.RecId     = lineas.Count;
                        linea.Fecha     = fechaTmp;
                        linea.Hora      = horaTemp;
                        linea.Dia       = fecha.Day;
                        linea.DiaSemana = fecha.ToString("dddd", new CultureInfo("es-ES")).ToUpper().Trim();

                        if (aviso != null)
                        {
                            linea.IdentifAviso = aviso.IdentifAviso;
                            linea.Duracion = aviso.Duracion;
                        }
                        else
                        {
                            linea.IdentifAviso = string.Empty;
                            linea.Duracion = null;
                        }

                        linea.Salida = spSalidasInsertar.Number;

                        if (!lineas.Exists((x) => (x.Fecha == fechaTmp && x.Hora == horaTemp && x.Salida == linea.Salida && x.IdentifAviso == linea.IdentifAviso)))
                            lineas.Add(linea);
                        else
                            preExistentes.Add(linea);

                        horaTemp = horaTemp.Add(incremento);

                         if ( aviso.Duracion == null || aviso.Duracion == 0)
                            horaTemp = horaFin;
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

        private void RefreshAbmGridUnsorted(ASPxGridView gv)
        {
            gv.DataSource = Lineas;
            gv.DataBind();
        }

        protected void gv_CustomColumnDisplayText(object sender, ASPxGridViewColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName != "Hora") return;

            if(e.Value != null)
                e.DisplayText = e.Value.ToString().Substring(0, 5);
        }
        
        private void RefreshAbmGrid(ASPxGridView gv)
        {
            var lineas = Lineas;
            lineas = lineas.OrderBy(p => p.Dia).ThenBy(r => r.Hora).ThenBy(s => s.Salida).ToList();

            gv.DataSource = lineas;
            gv.DataBind();
        }

        protected void btnBack_Click(object sender, ImageClickEventArgs e)
        {
            Back();
            VaciarCamposPauta();
            VaciarCamposDetalle();
            gvHome.Selection.UnselectAll();
            ASPxGridViewExporter1.GridViewID = "gvHome";
        }

        private void Back()
        {
            ucIdentifEspacio.ComboBox.SelectedIndex = -1;
            txMedio.Text = string.Empty;
            deAnoMes.Date = DateTime.Now;

            trPauta.Visible = false;
            trFind.Visible = true;
            
            Lineas = null;
            OrdenadoCab = null;
            Costos = null;            
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {

        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            gv.Selection.UnselectAll();

            try
            {
                decimal duracionTot = 0;

                OrdenadoCabDTO ordenado = OrdenadoCab;

                ordenado.IdentifFrecuencia  = Convert.ToString(ucIdentifFrecuencia.SelectedValue);
                ordenado.HoraInicio         = FormsHelper.ConvertToTimeSpan(teHoraInicio.DateTime);
                ordenado.HoraFin            = FormsHelper.ConvertToTimeSpan(teHoraFin.DateTime);
                ordenado.IdentifIntervalo   = Convert.ToString(ucIdentifIntervalo.SelectedValue);
               
                //Sumatoria de duracion en registros de la tabla OrdenadoDet.
                {
                    Lineas.ForEach(x => { duracionTot += (x.Duracion != null) ? x.Duracion.Value : 0; });

                    ordenado.DuracionTot = duracionTot;
                }

                //Cantidad de registros tabla OrdenadoDet cuyo campo IdentifAviso <> “”
                ordenado.CantSalidas = Lineas.FindAll(x => (x.IdentifAviso != null && x.IdentifAviso != string.Empty)).Count;

                if (ordenado.CantSalidas!=0)
                {
                    if (ordenado.RecId == 0)
                    {
                        //Es nuevo...
                        ordenado.PautaId = Business.Ordenados.GetNextPautaId().ToString();

                        if (ordenado.IdentifIntervalo == "") ordenado.IdentifIntervalo = null;

                        Ordenados.Create(ordenado, Lineas);
                    }
                    else
                    {
                        //Es modificacion...
                        if (Lineas.Count > 0)
                        {
                            if (Lineas[0].Salida > 0)
                            {
                                Labelx.Text = string.Empty;

                                lblErrorLineas.Text = Labelx.Text;

                                if (teHoraInicio.Text == Lineas[0].Hora.ToString().Substring(0,5) && teHoraFin.Text == Lineas[0].Hora.ToString().Substring(0, 5))
                                {
                                    Labelx.Text = string.Empty;

                                    lblErrorLineas.Text = Labelx.Text;
                                }
                                else
                                {
                                    if(Convert.ToDateTime(teHoraInicio.Text) > (Convert.ToDateTime(teHoraFin.Text)))
                                    {
                                           Labelx.Text = "Hora de inicio < Hora de Fin. \r\n";
                                    }

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
                        else
                        {
                            //No hay lineas preexistentes
                        }

                        if (ordenado.IdentifIntervalo == "")
                        {
                            ordenado.IdentifIntervalo = null;
                        }

                        Ordenados.Update(ordenado, Lineas);
                    }

                    OrdenadoCab = ordenado;

                    //Re-Calculo todos los Costos.
                    Business.Ordenados.CalcularCosto(OrdenadoCab, Costos, Lineas, ((Accendo)this.Master).Usuario.UserName);

                    ReCargarControles(OrdenadoCab);

                }
                else if (ordenado.CantSalidas == 0)
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

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            Back();
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
                case "btnInsert":   InsertItems(); break;
                case "btnEdit":        EditItem(); break;
                case "btnDelete":   DeleteItems(); break;
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
                case "btnCopy":       CopyItems(); break;
                case "btnReplace": ReplaceItems(); break;
                case "btnSKU":         QuerySKU(); break;
                default:                           break;
            }
        }

        private void QuerySKU()
        {
            trQuerySKU.Visible = true;

            RefreshSKUGrid(gvSKU);
        }

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
                        var linea = Lineas.Find(x => x.RecId == (int)gv.GetSelectedFieldValues("RecId")[0]);

                        //Para tener el ID en el momento en que vaya a guardar.
                        pnlEditLine.Attributes.Add("RecId", linea.RecId.ToString());

                        ucIdentifAvisoEdit.SelectedValue = linea.IdentifAviso;

                        AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoEdit.SelectedValue));

                        if (aviso != null)
                        {
                            spDuracionEdit.Value = aviso.Duracion;

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
            //actualizar hora desde/hasta para lineas nuevas en solapa copiar.
            ASPxPageControl2.ActiveTabPage = ASPxPageControl2.TabPages[1];

            trAccion.Visible = true;

            trEditLine.Visible = false;
        }
        
        private void InsertItems()
        {
            //actualizar hora desde/hasta para lineas nuevas.

            teHoraInicioInsertar.DateTime   = teHoraInicio.DateTime;
            teHoraFinInsertar.DateTime      = teHoraFin.DateTime;
            ASPxPageControl2.ActiveTabPage  = ASPxPageControl2.TabPages[0];
            spSalidasInsertar.Enabled       =  (teHoraInicio.Text == teHoraFin.Text && teHoraInicio.Text == "00:00");
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
                    List<DTO.OrdenadoDetDTO> aux = new List<DTO.OrdenadoDetDTO>();

                    //Creo una nueva coleccion con todos los productos menos los seleccionados, y la guardo en el Viewstate.
                    foreach (var linea in Lineas)

                        if (!FormsHelper.IsSelectedRecId(linea.RecId, gv))
                            aux.Add(linea);

                    Lineas = aux;
                    
                    gv.DataSource = Lineas;

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

        private void MsgErrorLinas(Exception ex)
        {
            lblErrorLineas.Text = ex.Message.ToLower().Contains("Violation of UNIQUE KEY constraint".ToLower()) == true ? "No puede ingresar un nuevo registro con su clave duplicada." :
                                  ex.Message.ToLower().Contains("FK".ToLower()) == true ? "No puede modificar/eliminar la clave de este registro, ya que se encuentra relacionado a otra entidad." :
                                  ex.Message;
        }

        protected void btnCopiarPeriodos_Click(object sender, EventArgs e)
        {
            try
            {

                OrdenadoCabDTO ordenado = OrdenadoCab;

                DateTime fechaTopeMes = new DateTime( Convert.ToInt32(ordenado.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(ordenado.AnoMes.ToString().Substring(4, 2)), 1).AddMonths(1).AddDays(-1);

                //Validaciones de Origen

                var CantDiasOrigen = (deFechaHastaOrigenCopiar.Date - deFechaDesdeOrigenCopiar.Date);
                var CantDiasDestino = (deFechaHastaDestinoCopiar.Date - deFechaDesdeDestinoCopiar.Date);

                if (CantDiasDestino < CantDiasOrigen)
                    throw new Exception("El período de ORIGEN es mayor que el de DESTINO");

                int Mes = Convert.ToInt32(ordenado.AnoMes.ToString().Substring(4, 2));

                if (
                     deFechaHastaDestinoCopiar.Date.Month != Mes ||
                     deFechaDesdeOrigenCopiar.Date.Month  != Mes ||
                     deFechaHastaOrigenCopiar.Date.Month  != Mes ||
                     deFechaHastaDestinoCopiar.Date.Month != Mes
                    
                    )
                {
                    throw new Exception("El período seleccionado no coincide con el ORDENADO");
                }
                
                //FECHA DESDE ORIGEN NO NULA
                if (deFechaDesdeOrigenCopiar.Value == null)
                    throw new Exception("Debe ingresar una fecha DESDE ORIGEN.");

                //FECHA DESDE ORIGEN DENTRO DEL PERIODO DEL ORDENADO
                if (deFechaDesdeOrigenCopiar.Date.Year.ToString() + deFechaDesdeOrigenCopiar.Date.Month.ToString().PadLeft(2,'0')  != ordenado.AnoMes.ToString())
                    throw new Exception("El mes/año en DESDE ORIGEN no corresponde al ordenado seleccionado.");

                //FECHA HASTA ORIGEN NO NULA
                if (deFechaHastaOrigenCopiar.Value == null)
                    throw new Exception("Debe ingresar una fecha HASTA ORIGEN.");

                //FECHA HASTA ORIGEN DENTRO DEL PERIODO DEL ORDENADO
                if (deFechaHastaOrigenCopiar.Date.Year.ToString() + deFechaHastaOrigenCopiar.Date.Month.ToString().PadLeft(2, '0') != ordenado.AnoMes.ToString())
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

                //FECHA HASTA DESTINO NO NULA
                if (deFechaHastaDestinoCopiar.Value == null)
                    throw new Exception("Debe ingresar una fecha HASTA DESTINO.");

                //FECHA DESDE DESTINO DENTRO DEL PERIODO DEL ORDENADO
                if (deFechaDesdeDestinoCopiar.Date.Year.ToString() + deFechaDesdeDestinoCopiar.Date.Month.ToString().PadLeft(2, '0') != ordenado.AnoMes.ToString())
                    throw new Exception("El mes/año en DESDE DESTINO no corresponde al ordenado seleccionado.");

                Lineas = PeriodosCopiar(deFechaDesdeOrigenCopiar.Date, deFechaHastaOrigenCopiar.Date, deFechaDesdeDestinoCopiar.Date, deFechaHastaDestinoCopiar.Date, Lineas);

                RefreshAbmGrid(gv);
                
                trAccion.Visible = false;
                
                lblErrorLineas.Text = "No olvide GRABAR antes de continuar.";
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }

        /// <summary>
        /// Copio el período seleccionado como origen en el período seleccionado como destino
        /// </summary>
        /// <param name="fOrigenDesde"></param>
        /// <param name="fOrigenHasta"></param>
        /// <param name="fDestinoDesde"></param>
        /// <param name="fDestinoHasta"></param>
        /// <param name="lineas"></param>
        /// <returns></returns>
        private List<OrdenadoDetDTO> PeriodosCopiar(DateTime fOrigenDesde, DateTime fOrigenHasta, DateTime fDestinoDesde, DateTime fDestinoHasta, List<OrdenadoDetDTO> lineas)
        {

            List<OrdenadoDetDTO> LineasOrigen           = lineas;
            List<OrdenadoDetDTO> LineasSeleccionadas    = new List<OrdenadoDetDTO>();
            List<OrdenadoDetDTO> LineasDestino          = lineas;

            FrecuenciaDTO frecuencia = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));

            List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

            lineas = Lineas.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

            EspacioContDTO espacio = GetEspacioContenido();

            string[,] DiasSemana = { { "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO" }, { "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY" } };

            var lineasACopiar = lineas.FindAll((x) => (x.Fecha.DayOfYear >= fOrigenDesde.DayOfYear && x.Fecha.DayOfYear <= fOrigenHasta.DayOfYear));

            DateTime origenM;
            // ARMADO DE LINEAS SELECCIONADAS //
            for (int i = 0; i <= LineasOrigen.Count - 1; i++)
            {
                origenM = new DateTime(LineasOrigen[i].Fecha.Year, LineasOrigen[i].Fecha.Month, LineasOrigen[i].Fecha.Day);

                if (origenM >= fOrigenDesde && origenM <= fOrigenHasta)
                    LineasSeleccionadas.Add(lineas[i]);
            }

            //ORDENA LAS LINEAS SELECCIONADAS
            LineasSeleccionadas = LineasSeleccionadas.OrderBy(o => o.Dia).ThenBy(p => p.Hora).ThenBy(q => q.Salida).ToList();

            //CALCULO CANTIDAD DE DIAS EN EL MES
            int DiasEnMes = System.DateTime.DaysInMonth(LineasOrigen[0].Fecha.Year, LineasOrigen[0].Fecha.Month);

            //CALCULO CUAL ES EL DIA DE LA SEMANA DEL 1 DEL MES
            int DiaSemana = Convert.ToInt32(Convert.ToDateTime(LineasOrigen[0].Fecha.Year.ToString() + "-" + LineasOrigen[0].Fecha.Month.ToString("00") + "-" + "01").DayOfWeek);

            // RECORRO UNO A UNO LOS DIAS DEL MES
            for (int i = 1; i <= DiasEnMes; i++)
            {
                // PREGUNTO SI EL DIA DEL MES ESTA DENTRO DEL RANGO DE DIAS SELECCIONADOS COMO DESTINO
                if (i >= fDestinoDesde.Day && i <= fDestinoHasta.Day)
                {
                    DateTime muleto = new DateTime(LineasOrigen[0].Fecha.Year, LineasOrigen[0].Fecha.Month, i);
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
                        DAO.OrdenadoDetDAO odd = new DAO.OrdenadoDetDAO();
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
                                OrdenadoDetDTO newLine = new OrdenadoDetDTO();

                                LastId++;

                                newLine.DatareaId   = 0;
                                newLine.RecId       = LastId;
                                newLine.Costo       = LineasSeleccionadas[j].Costo;
                                newLine.CostoOp     = LineasSeleccionadas[j].CostoOp;
                                newLine.CostoOpUni  = LineasSeleccionadas[j].CostoOpUni;
                                newLine.CostoUni    = LineasSeleccionadas[j].CostoUni;
                                newLine.Dia         = i;
                                newLine.Fecha = new DateTime(LineasSeleccionadas[j].Fecha.Year, LineasSeleccionadas[j].Fecha.Month, i);

                                string dia = newLine.Fecha.DayOfWeek.ToString().ToUpper();

                                for (int k = 0; k <= DiasSemana.Length - 1; k++)
                                {
                                    if (DiasSemana[1, k].ToUpper() == dia)
                                    {
                                        newLine.DiaSemana = DiasSemana[0, k].ToUpper();

                                        break;
                                    }
                                }

                                newLine.Duracion        = LineasSeleccionadas[j].Duracion;
                                newLine.Hora            = LineasSeleccionadas[j].Hora;
                                newLine.IdentifAviso    = LineasSeleccionadas[j].IdentifAviso;
                                newLine.PautaId         = LineasSeleccionadas[j].PautaId;
                                newLine.Salida          = LineasSeleccionadas[j].Salida;

                                LineasDestino.Add(newLine);
                            }
                        }
                    }
                }

                DiaSemana = DiaSemana == 7 ? 1 : DiaSemana++;
            }
            lineas = LineasDestino.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

            return lineas;
        }

        private List<OrdenadoDetDTO> CopiarPeriodos( DateTime fOrigenDesde, DateTime fOrigenHasta, DateTime fDestinoDesde,  List<OrdenadoDetDTO> lineas)
        {
            bool preExistentesFlag = false;

            EspacioContDTO espacio = GetEspacioContenido();
       
            //Armo lista de elemtnos que SI reemplazo.
            //Busco en la coleccion, todas las lineas en el periodo, y con el aviso seleccionado.

            lineas = lineas.OrderBy(o => o.Fecha).ThenBy( q => q.Dia).ThenBy(p => p.Hora).ToList();

            var lineasACopiar =lineas.FindAll((x) =>(x.Fecha >= fOrigenDesde && x.Fecha <= fOrigenHasta)).OrderBy( n => n.Fecha).ThenBy(t => t.Dia).ThenBy(y => y.Hora).ToList();

            lineasACopiar = lineasACopiar.OrderBy(b => b.Dia).ToList();

            //Si encontre líneas a copiar...
            if (lineasACopiar.Count > 0)
            {
                OrdenadoDetDTO nuevaLinea;
                DateTime fechaTmp;
                
                List<OrdenadoDetDTO> lineasTmp      = new List<OrdenadoDetDTO>();
                List<OrdenadoDetDTO> preExistentes  = new List<OrdenadoDetDTO>();
                
                DateTime diasEnElFuturo = fOrigenDesde;

                int cantDias;
                
                fechaTmp = fDestinoDesde;

                cantDias = 0;

                //Por cada linea que encontre, genero una nueva e igual, x dias en el futuro.

                int newDia = 0;

                foreach (var linea in lineasACopiar)
                {
                    FrecuenciaDTO frecuencia                    = CRUDHelper.Read( string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                    List<FrecuenciaDetDTO> frecuenciaDetalles   = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                    nuevaLinea = new OrdenadoDetDTO();

                    nuevaLinea.RecId = NextTempRecId();

                    nuevaLinea.DatareaId = linea.DatareaId;

                    //cargo lineas nuevas

                    if (fOrigenDesde != linea.Fecha)
                    {
                        if(fOrigenDesde.Day < Convert.ToInt32(linea.Dia) && newDia != Convert.ToInt32(linea.Dia))
                        {
                            cantDias = Convert.ToInt32(linea.Dia) - fOrigenDesde.Day;
                            newDia = Convert.ToInt32(linea.Dia);
                            fechaTmp = fechaTmp.AddDays(cantDias);
                        }

                        if (fOrigenDesde.Day < Convert.ToInt32(linea.Dia) && newDia != Convert.ToInt32(linea.Dia))
                        {
                            cantDias = fOrigenDesde.Day - Convert.ToInt32(linea.Dia);
                            newDia = Convert.ToInt32(linea.Dia);
                            fechaTmp = fechaTmp.AddDays(cantDias);
                        }

                        diasEnElFuturo = linea.Fecha;

                    }

                    nuevaLinea.Fecha    = fDestinoDesde.AddDays(cantDias);
                    nuevaLinea.Dia      = nuevaLinea.Fecha.Day;

                    string[] sDias = { "DOMINGO","LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"};

                    for (int x = 0; x <= 6; x++)
                    {
                        if ((int)nuevaLinea.Fecha.DayOfWeek == x)
                            nuevaLinea.DiaSemana = sDias[x];
                    }

                    //evito que se carguen datos fuera de los dias pautados
                    bool retval = false;

                    for (int k = 0; k <= frecuenciaDetalles.Count - 1; k++)
                    {
                        if (nuevaLinea.DiaSemana.Trim() == frecuenciaDetalles[k].DiaSemana.Trim())
                            retval = true;
                            break;
                    }

                    if (retval == false)
                        throw new Exception("No se puede grabar en dias de semana distintos a los pautados.");

                    nuevaLinea.Hora         = linea.Hora        ;
                    nuevaLinea.Costo        = linea.Costo       ;
                    nuevaLinea.CostoOp      = linea.CostoOp     ;
                    nuevaLinea.CostoOpUni   = linea.CostoOpUni  ;
                    nuevaLinea.CostoSalida  = linea.CostoSalida ;
                    nuevaLinea.CostoUni     = linea.CostoUni    ;
                    nuevaLinea.Duracion     = linea.Duracion    ;
                    nuevaLinea.IdentifAviso = linea.IdentifAviso;
                    nuevaLinea.PautaId      = linea.PautaId     ;
                    nuevaLinea.Salida       = linea.Salida      ;

                    foreach (OrdenadoDetDTO l in lineas)
                    {
                        if (l.Fecha == fechaTmp.Date)
                            if (l.Hora == linea.Hora)
                                if (l.Salida == nuevaLinea.Salida)
                                    preExistentes.Add(linea);
                    }

                    lineasTmp.Add(nuevaLinea);//Agrego la nueva linea.
                }

                //Junto las dos listas (temporal y la que ya tenia).
                lineasTmp.AddRange(lineas);

                //Ordeno por fecha.
                lineasTmp = lineasTmp.OrderBy(o => o.Dia).ThenBy(p => p.Hora).ToList();

                if (preExistentesFlag == true)
                    lblErrorLineas.Text = "No se pudieron grabar todas las lineas. No olvide GRABAR antes de continuar.";

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
            decimal salidas=0;

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

                    Lineas = ReemplazarAvisosPorPeriodo(deFechaDesdeReemplazar.Date.AddHours(deHoraDesdeOrigenReemplazar.DateTime.Hour).AddMinutes(deHoraDesdeOrigenReemplazar.DateTime.Minute),
                                                        deFechaHastaReemplazar.Date.AddHours(deHoraHastaOrigenReemplazar.DateTime.Hour).AddMinutes(deHoraHastaOrigenReemplazar.DateTime.Minute),
                                                        avisoOrigen, avisoDestino, Lineas, salidas);
                    RefreshAbmGrid(gv);
                    trAccion.Visible = false;
                }
                else if (opEditSeleccionados.Checked)
                {
                    Lineas = ReemplazarAvisosSeleccionados(avisoOrigen, avisoDestino, Lineas,salidas);
                }
                else if (opEditTodas.Checked)
                {
                   Lineas = ReemplazarAvisosTodos(avisoOrigen, avisoDestino, Lineas,salidas);
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

        private List<OrdenadoDetDTO> ReemplazarAvisosTodos(string avisoOrigen, string avisoDestino, List<OrdenadoDetDTO> lineas, decimal salidas)
        {
            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, DateTime.Now, DateTime.Now, "Todos");
        }

        private List<OrdenadoDetDTO> ReemplazarAvisosSeleccionados(string avisoOrigen, string avisoDestino, List<OrdenadoDetDTO> lineas,decimal salidas)
        {
           return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, DateTime.Now,DateTime.Now, "Seleccionado");
        }

        private List<OrdenadoDetDTO> ReemplazarAvisosPorPeriodo(DateTime fDesde, DateTime fHasta, string avisoOrigen, string avisoDestino, List<OrdenadoDetDTO> lineas, decimal salidas)
        {
            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, fDesde, fHasta, "Ingresado");
        }

        private List<OrdenadoDetDTO> ReemplazarAvisos(string avisoOrigen, string avisoDestino, List<OrdenadoDetDTO> lineas, decimal salidas, DateTime fDesde, DateTime fHasta, string Tipo = "Todo")
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
                                        TimeSpan tsPautaInicio  = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicio.Text));
                                        TimeSpan tsPautaFin     = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraFin.Text));
                                        TimeSpan tsAvisoViejoi  = linea.Hora;
                                        TimeSpan tsAvisoViejod  = new TimeSpan(0, 0, Convert.ToInt32(linea.Duracion));
                                        TimeSpan tsAvisoViejof  = tsAvisoViejoi.Add(tsAvisoViejod);
                                        TimeSpan tsAvisoNuevoi  = linea.Hora;
                                        TimeSpan tsAvisoNuevod  = new TimeSpan(0, 0, Convert.ToInt32(spAvisoReempDuracion.Value));
                                        TimeSpan tsAvisoNuevof  = tsAvisoNuevoi.Add(tsAvisoNuevod);

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
                                    TimeSpan tsPautaInicio  = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicio.Text));
                                    TimeSpan tsPautaFin     = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraFin.Text));
                                    TimeSpan tsAvisoViejoi  = linea.Hora;
                                    TimeSpan tsAvisoViejod  = new TimeSpan(0, 0, Convert.ToInt32(linea.Duracion));
                                    TimeSpan tsAvisoViejof  = tsAvisoViejoi.Add(tsAvisoViejod);
                                    TimeSpan tsAvisoNuevoi  = linea.Hora;
                                    TimeSpan tsAvisoNuevod  = new TimeSpan(0, 0, Convert.ToInt32(spAvisoReempDuracion.Value));
                                    TimeSpan tsAvisoNuevof  = tsAvisoNuevoi.Add(tsAvisoNuevod);

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

                                    TimeSpan tsPautaInicio  = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicio.Text));
                                    TimeSpan tsPautaFin     = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraFin.Text));
                                    TimeSpan tsAvisoViejoi  = linea.Hora;
                                    TimeSpan tsAvisoViejod  = new TimeSpan(0, 0, Convert.ToInt32(linea.Duracion));
                                    TimeSpan tsAvisoViejof  = tsAvisoViejoi.Add(tsAvisoViejod);
                                    TimeSpan tsAvisoNuevoi  = linea.Hora;
                                    TimeSpan tsAvisoNuevod  = new TimeSpan(0, 0, Convert.ToInt32(spAvisoReempDuracion.Value));
                                    TimeSpan tsAvisoNuevof  = tsAvisoNuevoi.Add(tsAvisoNuevod);

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

        protected void btnUpdateEdit_Click(object sender, EventArgs e)
        {
            try
            {
                //Guarde el RecId en un atributo del panel al momento de cargarlo.
                int RecId = Convert.ToInt32(pnlEditLine.Attributes["RecId"]);

                var lineas = Lineas;

                lineas = lineas.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                decimal? duracion = 0;

                AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoEdit.SelectedValue));

                if (aviso != null)
                    duracion = aviso.Duracion;

                ///// RECALCULO DURACION DEL AVISO ///////////

                IntervaloDTO intervalo = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));

                var tiempo = intervalo.CantMinutos * 60;

                if (duracion > tiempo)
                    throw new Exception("Solapamiento de horarios. Verifique.");

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

                    for (int i = 0; i <= lineas.Count - 1; i++)
                    {
                        if (lineas[i].RecId == ProxRecId)
                        {
                            if (lineas[i].Salida == 0)
                            {
                                TimeSpan tpHoraX = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicioModificar.Text));

                                DateTime tpHoraY = new DateTime(lineas[i].Fecha.Year, lineas[i].Fecha.Month, lineas[i].Fecha.Day, tpHoraX.Hours, tpHoraX.Minutes, 0);

                                decimal? tpX = duracion;

                                tpHoraY = tpHoraY.Add(new TimeSpan(0, 0, Convert.ToInt32(tpX)));

                                DateTime dt = new DateTime(lineas[i].Fecha.Year, lineas[i].Fecha.Month, lineas[i].Fecha.Day, lineas[i].Hora.Hours, lineas[i].Hora.Minutes, 0);

                                if (tpHoraY >= dt && tpX > 0)
                                {
                                    throw new Exception("Solapamiento de horarios. Verifique.");
                                }
                            }
                        }
                    }
                }

                ///// FIN DE FUNCION DE RECALCULAR DURACION //

                
                if (Convert.ToDateTime(teHoraInicioModificar.Text) < Convert.ToDateTime(teHoraInicio.Text) || Convert.ToDateTime(teHoraInicioModificar.Text) > Convert.ToDateTime(teHoraFin.Text))
                    throw new Exception("La hora no esta dentro del horario pautado.");

                decimal? tp = 0;

                TimeSpan tpHora;

                decimal tpSalida = 0;

                lineas = lineas.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

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
                                            HoraLimite = y.Hora.ToString();
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

                                    x.Duracion = duracion;

                                    DateTime dt = (DateTime)teHoraInicioModificar.DateTime;

                                    TimeSpan ts = new TimeSpan(0, dt.Hour, dt.Minute, 0);

                                    x.Hora = ts;

                                    x.Salida = Convert.ToDecimal(spAvisoModifiSalidas.Value);
                                }
                            }
                        }
                        else
                        {
                            //Cuando encuentre el item que estaba editando, lo actualizo.
                            x.IdentifAviso = Convert.ToString(ucIdentifAvisoEdit.SelectedValue);

                            x.Duracion = duracion;

                            DateTime dt = (DateTime)teHoraInicioModificar.DateTime;

                            TimeSpan ts = new TimeSpan(0, dt.Hour, dt.Minute, 0);

                            x.Hora = ts;

                            x.Salida = Convert.ToDecimal(spAvisoModifiSalidas.Value);


                        }
                    }
                }
                );

                Lineas = lineas;

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

        protected void gv_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            ASPxGridView gv = (ASPxGridView)sender;

            e.Cancel = true;

            var aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", (string)e.NewValues["IdentifAviso"]));   

            var lineas = Lineas;

            lineas.FindAll(x => x.RecId == (int)e.Keys[0]).ForEach(
                (linea) => 
                { 
                    linea.IdentifAviso = aviso.IdentifAviso;
                    linea.Duracion     = aviso != null ? aviso.Duracion : null;
                    linea.Salida       = (decimal)e.NewValues["Salida"];

                });

            Lineas = lineas;

            gv.CancelEdit();

            RefreshAbmGrid(gv);
        }

        protected void gv_StartRowEditing(object sender, DevExpress.Web.Data.ASPxStartRowEditingEventArgs e)
        {
            trEditLine.Visible = false;
        }

        protected void btnRefreshSKU_Click(object sender, EventArgs e)
        {
            RefreshSKUGrid(gvSKU);
        }

        protected void gvHome_CustomCallback(object sender, DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs e)
        {
        }

        protected void btnCancelSKU_Click(object sender, EventArgs e)
        {
            trQuerySKU.Visible = false;
        }

        protected void mnuPrincipal_ItemClick(object source, MenuItemEventArgs e)
        {

           mnuPrincipal_ItemClick1(source, e);
        }
        
        protected void mnuPrincipal_ItemClick1(object source, MenuItemEventArgs e)
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

                        Pauta = "OP ORDENADO " + gvHome.GetSelectedFieldValues("año")[0].ToString() + gvHome.GetSelectedFieldValues("mes")[0].ToString() + " " + gvHome.GetSelectedFieldValues("IdentifEspacio")[0].ToString() + " ";

                        string filename = Pauta + System.DateTime.Now.Hour.ToString().PadLeft(2,'0') + System.DateTime.Now.Minute.ToString().PadLeft(2,'0') + System.DateTime.Now.Second.ToString().PadLeft(2,'0') + ".xlsx";

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
                                            lblErrorHome.Text = "";

                                            string[] dirs = Directory.GetFiles(@TextBox1.Text);

                                            dExiste = false;

                                            filename = @TextBox1.Text + filename;

                                            if (dirs.Length > 0)
                                            {
                                                for (int i = 0; i <= dirs.Length - 1; i++)
                                                {
                                                    if (dirs[i].ToString() == filename )
                                                    {
                                                        dExiste = true;

                                                        break;
                                                    }
                                                }
                                            }

                                            if (dExiste == false)
                                            { 
                                                OrdenadoCabDTO Cabecera         = Ordenados.Read(gvHome.GetSelectedFieldValues("PautaId")[0].ToString());
                                                List<OrdenadoDetDTO> Detalle    = Ordenados.ReadAllLineas(Cabecera);
                                                List<OrdenadoSKUDTO> SKUS       = Ordenados.ReadAllSKUs(Cabecera);
                                                EspacioContDTO Espacio          = CRUDHelper.Read( string.Format("IdentifEspacio = '{0}'",Cabecera.IdentifEspacio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

                                                csOP_Helper Helper = new csOP_Helper("ORDENADO", "", gvHome.GetSelectedFieldValues("PautaId")[0].ToString(),Cabecera,Detalle,SKUS,Espacio,filename);

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

                                        lblErrorHome.Text = "Ruta no válida.";
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

                    Business.Ordenados.CalcularCosto(OrdenadoCab, Costos, Lineas, ((Accendo)this.Master).Usuario.UserName);

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

        protected void ASPxPageControl2_ActiveTabChanged(object source, DevExpress.Web.ASPxTabControl.TabControlEventArgs e)
        {
            RecargarDiaHora();

            string msg = string.Empty;

            if (!ValidarFranjaHoraria(ref msg))
                lblErrorLineas.Text = msg;

            switch (e.Tab.Text)
            {
                case "Copiar Períodos": { CopyItems(); break; }
                case "Reemplazar Avisos":{ ReplaceItems(); break; }
                case "Insertar Líneas": { InsertItems(); break; }
            }
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
                    ucIdentifEspacio.SelectedValue  = null;
                    txMedio.Text                    = "";
                    deAnoMes.Value                  = new DateTime();
                }
                CargarOrdenado();

                gv.SortBy(gv.Columns["Dia"]   , DevExpress.Data.ColumnSortOrder.Ascending);
                gv.SortBy(gv.Columns["Fecha"] , DevExpress.Data.ColumnSortOrder.Ascending);
                gv.SortBy(gv.Columns["Hora"]  , -1);
                gv.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);

                DAO.OrdenadoCabDAO ocd = new DAO.OrdenadoCabDAO();

                DateTime fs = ocd.FechaServer;
            }
            else
            {
                lblErrorHome.Text = "Debe Seleccionar un registro para proceder";
            }
        }

        protected void btn_CancelDelete(object sender, EventArgs e)
        {
            tblDelete.Visible = false;
            lblErrorHome.Visible = false;
            gvHome.Selection.UnselectAll();
        }

        protected void btn_ShowDelete(object sender, EventArgs e)
        {
            if (gvHome.Selection.Count > 0)
            {
                tblDelete.Visible = true;
                lblErrorLineas.Text = string.Empty;
                lblErrorHome.Text = string.Empty;
                lblValidaAñoMes.Text= string.Empty;
            }
            else
            {
                lblErrorHome.Text = "Debe seleccionar un ordenado de la lista";
            }
        }
        
        protected void btn_EliminarLineaGvHome(object sender, EventArgs e)
        {
            if (gvHome.Selection.Count > 0)
            {
                int recId = Convert.ToInt32(gvHome.GetSelectedFieldValues(new string[] { "RecId" })[0]);

                string UsuCierre = Convert.ToString(gvHome.GetSelectedFieldValues(new string[] { "UsuCierre" })[0]);

                if (recId > 0)
                {
                    if (UsuCierre == "")
                    {
                        try
                        {

                            Ordenados.Delete(Ordenados.Read(recId));

                            gvHome.DataSource = Ordenados.VistaOrdenados();

                            gvHome.DataBind();

                            tblDelete.Visible = false;

                            lblErrorHome.Text = "Ordenado borrado correctamente.";
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
            }
            else
            {
                lblErrorHome.Text = "Debe Seleccionar un registro para proceder.";
            }
        }

        protected void txMedio_ValueChanged(object sender, EventArgs e)
        {
            OrdenadoCabDTO ord = Ordenados.Read(Convert.ToInt32(ddlNroPauta.SelectedValue));

            if (ord != null) //funcion para completar Medio
            {
                ucIdentifEspacio.SelectedValue = ord.IdentifEspacio;
                
                DateTime d = new DateTime(Convert.ToInt32(ord.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(ord.AnoMes.ToString().Substring(4, 2)), 01);

                deAnoMes.Value = d;

                ucEspacioChanged();
            }
        }

        protected void btnBack_Click(object sender, EventArgs e)
        {
            btnBack_Click(sender, null);
        }

        protected void teHoraFinInsertar_ValueChanged(object sender, EventArgs e)
        {
            spSalidasInsertar.Enabled = (teHoraInicio.Text != teHoraFinInsertar.Text && teHoraFinInsertar.Text == "00:00");
        }

        protected void teHoraFinInsertar_DateChanged(object sender, EventArgs e)
        {
        }

        protected bool Validaciones(TimeSpan horaInicio, TimeSpan horaFin, IntervaloDTO intervalo, FrecuenciaDTO frecuenciaCab, AvisosDTO aviso, List<FrecuenciaDetDTO> frecuenciaDetalles)
        {
            bool retval = false;

            decimal CantMinutosAviso = 0;

            List<OrdenadoDetDTO> lineas = Lineas;

            List<OrdenadoDetDTO> preExistentes = new List<OrdenadoDetDTO>();

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
            if (Convert.ToInt32(spDuracionInsertar.Value) == 0 && spSalidasInsertar.Enabled == false && lineas.Count == 0)
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
                if (ts2i.CompareTo(ts1i) == 0 && spDuracionEdit.Text != "0")
                {
                    lblErrorLineas.Text = "El aviso se superpone con otro";

                    retval = false;

                    break;
                }
                
                //Verificacion contra el registro siguiente
                //Quiere decir que hay un registro previo despues del actual
                if (j < lineas.Count)
                {
                    /* FECHA        HORAI   HORAF   DURACION
                     * 19/06/2013   12:00   12:30   00:15
                     * 19/06/2013   12:30   13:00   00:15
                     * 19/06/2013   13:00   13:30   00:15
                     * 19/06/2013   13:30   14:00   00:15
                     */
                    //Datos del registro siguiente
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
                else
                { 
                    //estoy en el ultimo registro, no hay nada mas desde aqui en la lista...
                }

                if (ts1i.Days == ts2i.Days)
                {
                    TimeSpan RangoInicio            = ts1i;
                    TimeSpan RangoFinal             = ts1f;
                    TimeSpan HoraInicioAvisoNuevo   = ts2i;
                    TimeSpan HoraFinAvisoNuevo      = ts2f;

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
    }
}