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
using PautasPublicidad.Web.Classes;

namespace PautasPublicidad.Web.Forms
{
    public partial class EstimadoAnulacionReemplazo : System.Web.UI.Page
    {
        #region "ViewState"

        int ProxRecId = 0;

        public List<EstimadoDetDTO> Lineas
        {
            get
            {
                var e = Estimado.Lineas;
                return e;
            }
            set
            {
                var e    = Estimado;
                e.Lineas = value;
                Estimado = e;
            }
        }

        public Business.Estimado Estimado
        {
            get
            {
                if (Session["EstimadoAnulacionReemplazo.Estimado"] != null && Session["EstimadoAnulacionReemplazo.Estimado"] is Business.Estimado)
                    return Session["EstimadoAnulacionReemplazo.Estimado"] as Business.Estimado;
                else
                    return null;
            }
            set
            {                
                Session.Add("EstimadoAnulacionReemplazo.Estimado", value);
            }
        }

        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            GridViewDataComboBoxColumn gvc    = gv.Columns["IdentifAviso"] as GridViewDataComboBoxColumn;
            gvc.Name                          = "IdentifAviso";
            gvc.Caption                       = "Aviso";
            gvc.FieldName                     = "IdentifAviso";
            gvc.PropertiesComboBox.TextField  = "Name"; //mapInfo.EntityTextField;
            gvc.PropertiesComboBox.ValueField = "IdentifAviso"; // mapInfo.EntityValueField;
            gvc.PropertiesComboBox.DataSource = Business.Avisos.ReadAll(""); //mapInfo.DAOHandler.ReadAll("");

            ucIdentifFrecuencia.ComboBox.AutoPostBack = true;
            ucIdentifFrecuencia.ComboBox.SelectedIndexChanged += new EventHandler(IdentifFrecuencia_SelectedIndexChanged);

            ucIdentifAviso.ComboBox.AutoPostBack = true;
            ucIdentifAviso.ComboBox.SelectedIndexChanged += new EventHandler(IdentifAviso_SelectedIndexChanged);

            ucIdentifAvisoEdit.ComboBox.AutoPostBack = true;
            ucIdentifAvisoEdit.ComboBox.SelectedIndexChanged += new EventHandler(IdentifAvisoEdit_SelectedIndexChanged);

            ucIdentifAvisoDestinoReemplazar.ComboBox.AutoPostBack = true;
            ucIdentifAvisoDestinoReemplazar.ComboBox.SelectedIndexChanged += new EventHandler(AvisoDestinoReemplazar_SelectedIndexChanged);

            ucIdentifFrecuencia.Inicializar(BusinessMapper.eEntities.Frecuencia);
            ucIdentifIntervalo.Inicializar(BusinessMapper.eEntities.Intervalo);

            if (!Page.IsPostBack && !Page.IsCallback && Request.QueryString["Estimado.RecId"] != null)
            {
                trAccion.Visible          = false;
                trEditLine.Visible        = false;
                trQuerySKU.Visible        = false;
                opEditPeriodo.Checked     = true;
                spSalidasInsertar.Enabled = true;

                Estimado = Estimados.Get(Convert.ToInt32(Request.QueryString["Estimado.RecId"]));
                ReCargarControles(Estimado);

                if (!Page.IsPostBack && !Page.IsCallback)
                {
                    FormsHelper.InicializarPropsGrilla(gv);
                    gv.SettingsEditing.Mode                      = GridViewEditingMode.Inline;
                    gv.SettingsBehavior.AllowSelectByRowClick    = true;
                    gv.SettingsBehavior.AllowSelectSingleRowOnly = false;
                }
            }

            ucIdentifAviso.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio));
            ucIdentifAviso.WhereFilter = string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio);

            ucIdentifAvisoOrigenReemplazar.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio));
            ucIdentifAvisoOrigenReemplazar.WhereFilter = string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio);

            ucIdentifAvisoDestinoReemplazar.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio));
            ucIdentifAvisoDestinoReemplazar.WhereFilter = string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio);

            ucIdentifAvisoEdit.Inicializar(BusinessMapper.eEntities.Avisos, string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio));
            ucIdentifAvisoEdit.WhereFilter = string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio);

            lblErrorLineas.Text = string.Empty;

            gv.KeyFieldName = "RecId";
            RefreshAbmGrid(gv);

            if (trQuerySKU.Visible)
            {
                RefreshSKUGrid(gvSKU);
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
                spAvisoReempDuracion.Value = aviso.Duracion;
        }

        private void RefreshSKUGrid(ASPxGridView gvSKU)
        {
            decimal total = 0;
            var dt = Estimados.BuildAllSKU(Lineas);

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

            e.DisplayText = e.Value.ToString().Substring(0, 5);

        }

        private void RefreshAbmGrid(ASPxGridView gvABM)
        {
            List<MiEstimadoDetDTO> miLista = new List<MiEstimadoDetDTO>();

             AgregarCodigoAviso(miLista);

            //reordena la lista
            gvABM.DataSource = miLista;

            gvABM.SortBy(gv.Columns["Fecha"] , DevExpress.Data.ColumnSortOrder.Ascending);
            gvABM.SortBy(gv.Columns["Dia"]   , DevExpress.Data.ColumnSortOrder.Ascending);
            gvABM.SortBy(gv.Columns["Hora"]  , DevExpress.Data.ColumnSortOrder.Ascending);
            gvABM.SortBy(gv.Columns["Salida"], DevExpress.Data.ColumnSortOrder.Ascending);

            gvABM.DataBind();
        }

        private void AgregarCodigoAviso(List<MiEstimadoDetDTO> miLista)
        {
            for (int i = 0; i <= Estimado.Lineas.Count - 1; i++)
            {
                var newList = new MiEstimadoDetDTO();

                newList.CodigoAviso  = Estimado.Lineas[i].IdentifAviso  ;
                newList.Costo        = Estimado.Lineas[i].Costo         ;
                newList.CostoOp      = Estimado.Lineas[i].CostoOp       ;
                newList.CostoOpUni   = Estimado.Lineas[i].CostoOpUni    ;
                newList.CostoUni     = Estimado.Lineas[i].CostoUni      ;
                newList.DatareaId    = Estimado.Lineas[i].DatareaId     ;
                newList.Dia          = Estimado.Lineas[i].Dia           ;
                newList.DiaSemana    = Estimado.Lineas[i].DiaSemana     ;
                newList.Duracion     = Estimado.Lineas[i].Duracion      ;
                newList.Fecha        = Estimado.Lineas[i].Fecha         ;
                newList.Hora         = Estimado.Lineas[i].Hora          ;
                newList.IdentifAviso = Estimado.Lineas[i].IdentifAviso  ;
                newList.PautaId      = Estimado.Lineas[i].PautaId       ;
                newList.RecId        = Estimado.Lineas[i].RecId         ;
                newList.Salida       = Estimado.Lineas[i].Salida        ;

                miLista.Add(newList);

            }
        }

        private void ReCargarControles(Business.Estimado estimado)
        {
            lblErrorLineas.Text = string.Empty;

            //Controles de la pauta.
            spPautaID.Number                        = Convert.ToInt32(estimado.Cabecera.PautaId);
            ucIdentifFrecuencia.SelectedValue       = estimado.Cabecera.IdentifFrecuencia;
            ucIdentifIntervalo.SelectedValue        = estimado.Cabecera.IdentifIntervalo;
            teHoraInicio.DateTime                   = FormsHelper.ConvertToDateTime(estimado.Cabecera.HoraInicio);
            teHoraInicioInsertar.DateTime           = teHoraInicio.DateTime;
            teHoraFin.DateTime                      = FormsHelper.ConvertToDateTime(estimado.Cabecera.HoraFin);
            teHoraFinInsertar.DateTime              = teHoraFin.DateTime;

            //Controles de solo lectura.
            spVersionCosto.Value                    = estimado.Cabecera.VersionCosto;
            txUsuCosto.Text                         = estimado.Cabecera.UsuCosto;
            deFecCosto.Date                         = estimado.Cabecera.FecCosto;
            txUsuCierre.Text                        = estimado.Cabecera.UsuCierre;
            deFecCierre.Value                       = estimado.Cabecera.FecCierre;
            spCantSalidas.Value                     = estimado.Cabecera.CantSalidas;

            //Actualizo controles.
            ucFrecuenciaChanged();

            btnSave.Enabled                         = (estimado.Cabecera.FecCierre == null);

            //Inicializo Fechas...
            deFechaDesdeOrigenCopiar.Date           = estimado.Cabecera.VigDesde;
            deFechaDesdeDestinoCopiar.Date          = estimado.Cabecera.VigDesde;
            deFechaDesdeReemplazar.Date             = estimado.Cabecera.VigDesde;
            deFechaHastaReemplazar.Date             = estimado.Cabecera.VigHasta;
            deFechaHastaOrigenCopiar.Date           = estimado.Cabecera.VigHasta; 
            deHoraDesdeOrigenReemplazar.DateTime    = FormsHelper.ConvertToDateTime(estimado.Cabecera.HoraInicio);
            deHoraHastaOrigenReemplazar.DateTime    = FormsHelper.ConvertToDateTime(estimado.Cabecera.HoraFin);

        }

        private void ucEspacioChanged()
        {
            EspacioContDTO espacio = GetEspacioContenido();

            string txMedio;

            if (espacio != null)
            {
                txMedio = espacio.IdentifMedio;
                DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                ucIdentifFrecuencia.SelectedValue = espacio.IdentifFrecuencia;

                if (espacio.HoraInicio.HasValue && espacio.HoraFin.HasValue)
                {
                    teHoraInicio.DateTime = FormsHelper.ConvertToDateTime(espacio.HoraInicio.Value);
                    teHoraFin.DateTime    = FormsHelper.ConvertToDateTime(espacio.HoraFin.Value);
                }
                else
                {
                    teHoraInicio.Value = null;
                    teHoraFin.Value    = null;
                }

                ucIdentifIntervalo.SelectedValue = espacio.IdentifIntervalo;
            }
            else
            {
                txMedio = string.Empty;
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
                        ceDiasInsertar.Items.Add(frecuenciaDetalle.Dia.Value.ToString(), frecuenciaDetalle.Dia.Value.ToString());
                    else
                        ceDiasInsertar.Items.Add(frecuenciaDetalle.DiaSemana, frecuenciaDetalle.DiaSemana.ToUpper().Trim());

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

                if (aviso.Duracion>0)
                {
                    spSalidasInsertar.Value = 0;
                    spSalidasInsertar.EnableClientSideAPI = false;
                }
                else
                {
                    spSalidasInsertar.Enabled = true;
                    lblErrorLineas.Text = "No olvide ingresar Numero de Salidas";
                }
            }
        }

        private EspacioContDTO GetEspacioContenido()
        {
            if (Estimado.Cabecera.IdentifEspacio != null)
            {
                return CRUDHelper.Read(string.Format("IdentifEspacio = '{0}'", Estimado.Cabecera.IdentifEspacio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
            }
            else
                return null;
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
            GenerarLineas(teHoraInicioInsertar.DateTime, teHoraFinInsertar.DateTime);
            RefreshAbmGrid(gv);

            trAccion.Visible = false;
        }

        private void GenerarLineas(DateTime horaInicio, DateTime horaFin)
        {
            try
            {
                IntervaloDTO intervalo                    = CRUDHelper.Read(string.Format("IdentifIntervalo = '{0}'", ucIdentifIntervalo.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Intervalo));
                FrecuenciaDTO frecuencia                  = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll( string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAviso.SelectedValue));

                //Validaciones...
                if (intervalo == null)
                    throw new Exception("Debe seleccionar un 'Intervalo'");

                if (frecuencia == null)
                    throw new Exception("Debe seleccionar una 'Frecuencia'");

                if(ucIdentifAviso.SelectedValue==null)
                    throw new Exception("Debe seleccionar una 'Aviso'");

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

        private List<EstimadoDetDTO> GenerarLineas(TimeSpan horaInicio, TimeSpan horaFin, IntervaloDTO intervalo, FrecuenciaDTO frecuenciaCab, AvisosDTO aviso, List<FrecuenciaDetDTO> frecuenciaDetalles)
        {
            List<EstimadoDetDTO> lineas        = Lineas;
            List<EstimadoDetDTO> preExistentes = new List<EstimadoDetDTO>();


            if (Validaciones(horaInicio, horaFin, intervalo, frecuenciaCab, aviso, frecuenciaDetalles) == false)
            {
                RefreshAbmGrid(gv);
                return lineas;
            }
            
            TimeSpan incremento = TimeSpan.FromMinutes(Convert.ToDouble(intervalo.CantMinutos));
            TimeSpan horaTemp;
            EstimadoDetDTO linea;
            List<DateTime> periodo;

            int year = Convert.ToInt32(Estimado.Cabecera.AnoMes.ToString().Substring(0,4));
            int month = Convert.ToInt32(Estimado.Cabecera.AnoMes.ToString().Substring(4,2));

            try
            {
                //Obtengo la lista de los días partiedo de fechas o nombres de dia.
                if (frecuenciaCab.SemMes == "SEMANA")
                    periodo = Ordenados.GetDatesByDayNames(year, month, GetDiasSeleccionados());
                else
                    periodo = Ordenados.GetDatesByDayNumbers(year, month, GetDiasSeleccionados());

                foreach (DateTime fecha in periodo)
                {
                    horaTemp = horaInicio;

                    //Mientras no supere la hora hasta...
                    while (horaTemp.CompareTo(horaFin) < 0)
                    {
                        DateTime fechaTmp = new DateTime(fecha.Year, fecha.Month, fecha.Day, horaTemp.Hours, horaTemp.Minutes, horaTemp.Seconds);

                        linea               = new EstimadoDetDTO();

                        linea.RecId         = lineas.Count;
                        linea.Fecha         = fechaTmp; 
                        linea.Hora          = horaTemp;
                        linea.Dia           = fecha.Day;
                        linea.DiaSemana     = fecha.ToString("dddd", new CultureInfo("es-ES")).ToUpper().Trim();
                        linea.IdentifAviso  = aviso != null ? aviso.IdentifAviso : string.Empty;
                        linea.Duracion      = aviso != null ? aviso.Duracion : null;
                        linea.Salida        = spSalidasInsertar.Value != null ? spSalidasInsertar.Number : 0;

                        lineas.Add(linea);

                        if (!lineas.Exists(
                            (x) => (x.Fecha == fechaTmp && x.Hora == horaTemp)))
                        {
                            lineas.Add(linea);
                        }
                        else
                        {
                            preExistentes.Add(linea);
                        }

                        horaTemp = horaTemp.Add(incremento);
                    }
                }
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }

            if (preExistentes.Count > 0)
                lblErrorLineas.Text = "No se pudieron grabar todas las lineas";
        
            return lineas;
        }

        //HTTD
        protected void ASPxPageControl1_ActiveTabChanged(object source, DevExpress.Web.ASPxTabControl.TabControlEventArgs e)
        {

        }

        //HTTD
        protected void btnDelete_Click(object sender, EventArgs e)
        {

        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                decimal duracionTot = 0;
                EstimadoCabDTO estimadoCab = Estimado.Cabecera; //OrdenadoCab;

                estimadoCab.IdentifFrecuencia = Convert.ToString(ucIdentifFrecuencia.SelectedValue);
                estimadoCab.HoraInicio        = FormsHelper.ConvertToTimeSpan(teHoraInicio.DateTime);
                estimadoCab.HoraFin           = FormsHelper.ConvertToTimeSpan(teHoraFin.DateTime);
                estimadoCab.IdentifIntervalo  = Convert.ToString(ucIdentifIntervalo.SelectedValue);
                estimadoCab.Version           = estimadoCab.Version + 1;
                estimadoCab.FecUltModif       = DateTime.Now;

                //Sumatoria de duracion en registros de la tabla EstimadoDet.
                if (estimadoCab.DuracionTot != 0)
                {
                    Lineas.ForEach(x => { duracionTot += x.Duracion.Value; });
                }

                estimadoCab.DuracionTot = duracionTot;

                //Cantidad de registros tabla EstimadoDet cuyo campo IdentifAviso <> “”
                estimadoCab.CantSalidas = Lineas.FindAll(x => (x.IdentifAviso != null && x.IdentifAviso != string.Empty)).Count;

                Estimado.Cabecera = estimadoCab;

                //Confirma el cierre.
                Estimado = Estimados.Confirmar(estimadoCab, Lineas, Estimado.SKUs);

                ReCargarControles(Estimado);

                lblErrorLineas.Text = "Se Grabo correctamente";

                CostosDTO costos = Business.Ordenados.FindCosto(estimadoCab.IdentifEspacio, Convert.ToInt32(estimadoCab.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(estimadoCab.AnoMes.ToString().Substring(4, 2)));

                //Re-Calculo todos los Costos.
                Business.Estimados.CalcularCosto(estimadoCab, costos, Lineas,((Accendo)this.Master).Usuario.UserName);

            }
            catch (Exception ex)
            {
                lblErrorLineas.Text = ex.Message;

            }
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            trAccion.Visible = false;
            gv.Selection.UnselectAll();
        }

        protected void mnuDetalle_ItemClick(object source, MenuItemEventArgs e)
        {
            switch (e.Item.Name)
            {
                case "btnInsert":
                    InsertItems();
                    break;

                case "btnEdit":
                    EditItem();
                    break;

                case "btnDelete":
                    DeleteItems();
                    break;

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

                case "btnCopy":
                    CopyItems();
                    break;

                case "btnReplace":
                    ReplaceItems();
                    break;

                default:
                    break;
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
                        var linea = Lineas.Find(x => x.RecId == (int)gv.GetSelectedFieldValues("RecId")[0]);

                        //Para tener el ID en el momento en que vaya a guardar.
                        pnlEditLine.Attributes.Add("RecId", linea.RecId.ToString());


                        // teHoraEdit.DateTime = FormsHelper.ConvertToDateTime(linea.Hora);
                        ucIdentifAvisoEdit.SelectedValue = linea.IdentifAviso;
                        AvisosDTO aviso                  = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", ucIdentifAvisoEdit.SelectedValue));

                        if (aviso != null)
                        {
                            spDuracionEdit.Value = aviso.Duracion;
                            ASPxTimeEdit miTimeEdit = new ASPxTimeEdit();

                            miTimeEdit.DateTime            = Convert.ToDateTime(gv.GetSelectedFieldValues("Hora")[0].ToString());
                            teHoraInicioModificar.DateTime = miTimeEdit.DateTime;
                            spAvisoModifiSalidas.Value     = gv.GetSelectedFieldValues("Salida")[0].ToString();
                            teHoraInicioModificar.Enabled  = spAvisoModifiSalidas.Text == "0";
                            spAvisoModifiSalidas.Enabled   = teHoraInicioModificar.Enabled == false;
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
            trAccion.Visible               = true;
            trEditLine.Visible             = false;
        }

        private void CopyItems()
        {
            ASPxPageControl2.ActiveTabPage = ASPxPageControl2.TabPages[1];
            trAccion.Visible               = true;
            trEditLine.Visible             = false;
        }

        private void InsertItems()
        {
            ASPxPageControl2.ActiveTabPage = ASPxPageControl2.TabPages[0];
            spSalidasInsertar.Enabled      = (teHoraInicio.Text == teHoraFin.Text && teHoraInicio.Text == "00:00");
            spSalidasInsertar.Value        = spSalidasInsertar.Enabled ? 1 : 0;
            teHoraInicioInsertar.Enabled   = (spSalidasInsertar.Enabled == false);
            teHoraFinInsertar.Enabled      = (spSalidasInsertar.Enabled == false);
            trAccion.Visible               = true;
            trEditLine.Visible             = false;
        }

        private void DeleteItems()
        {
            try
            {
                if (FormsHelper.GetSelectedId(gv) != null)
                {
                    List<DTO.EstimadoDetDTO> aux = new List<DTO.EstimadoDetDTO>();

                    foreach (var linea in Lineas)

                        if (!FormsHelper.IsSelectedRecId(linea.RecId, gv))
                            aux.Add(linea);

                    Lineas = aux;
                    gv.Selection.UnselectAll();
                    RefreshAbmGrid(gv);
                    gv.Selection.UnselectAll();
                    lblErrorLineas.Text = "Linea eliminada correctamente. No olvide Guardar antes de salir";
                }
                else
                {
                    lblErrorLineas.Text = "Debe seleccionar una linea para eliminar";
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
            if (ex.Message.ToLower().Contains("Violation of UNIQUE KEY constraint".ToLower()))
                lblErrorLineas.Text = "No puede ingresar un nuevo registro con su clave duplicada.";
            else if (ex.Message.ToLower().Contains("FK".ToLower()))
                lblErrorLineas.Text = "No puede modificar/eliminar la clave de este registro, ya que se encuentra relacionado a otra entidad.";
            else
                lblErrorLineas.Text = ex.Message;
        }

        #region "Solapas internas"
        protected void btnCopiarPeriodos_Click(object sender, EventArgs e)
        {
            try
            {
                //Validaciones...
                if (deFechaDesdeOrigenCopiar.Value == null)
                    throw new Exception("Debe seleccionar una 'Fecha Desde' (origen).");
                if (deFechaHastaOrigenCopiar.Value == null)
                    throw new Exception("Debe seleccionar una 'Fecha Hasta' (origen).");
                if (deFechaDesdeDestinoCopiar.Value == null)
                    throw new Exception("Debe seleccionar una 'Fecha Desde' (destino).");

                Lineas = PeriodosCopiar(deFechaDesdeOrigenCopiar.Date, deFechaHastaOrigenCopiar.Date, deFechaDesdeDestinoCopiar.Date, deFechaHastaDestinoCopiar.Date, Lineas);

                RefreshAbmGrid(gv);
                trAccion.Visible = false;
            }
            catch (Exception ex)
            {
                MsgErrorLinas(ex);
            }
        }

        private List<EstimadoDetDTO> CopiarPeriodos(DateTime fOrigenDesde, DateTime fOrigenHasta, DateTime fDestinoDesde, List<EstimadoDetDTO> lineas)
        {

            bool preExistentesFlag = false;
            EspacioContDTO espacio = GetEspacioContenido();

            //Armo lista de elementos que SI reemplazo.
            //Busco en la coleccion, todas las lineas en el periodo, y con el aviso seleccionado.
            var lineasACopiar = lineas.FindAll(
                (x) =>
                    (x.Fecha >= fOrigenDesde
                    && x.Fecha <= fOrigenHasta
                    ));


            //Si encontre líneas a copiar...
            if (lineasACopiar.Count > 0)
            {
                EstimadoDetDTO nuevaLinea;
                DateTime fechaTmp;
                List<EstimadoDetDTO> lineasTmp = new List<EstimadoDetDTO>();
                List<EstimadoDetDTO> preExistentes = new List<EstimadoDetDTO>();

                DateTime diasEnElFuturo = fOrigenDesde;
                int cantDias;
                fechaTmp = fDestinoDesde;

                cantDias = 0;
                //Por cada linea que encontre, genero una nueva e igual, x dias en el futuro.
                foreach (var linea in lineasACopiar)
                {
                    FrecuenciaDTO frecuencia = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));

                    List<FrecuenciaDetDTO> frecuenciaDetalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", frecuencia.IdentifFrecuencia),
                                                                BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                    nuevaLinea           = new EstimadoDetDTO();
                    nuevaLinea.RecId     = NextTempRecId();
                    nuevaLinea.DatareaId = linea.DatareaId;

                    //cargo lineas nuevas
                    if (fOrigenDesde != linea.Fecha)
                    {
                        cantDias++;
                        diasEnElFuturo = linea.Fecha;
                        fechaTmp       = fechaTmp.AddDays(cantDias);
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
                        throw new Exception("No se puede grabar en dias de semana distintos a los pautados.");

                    nuevaLinea.Hora         = linea.Hora        ;
                    nuevaLinea.Costo        = linea.Costo       ;
                    nuevaLinea.CostoOp      = linea.CostoOp     ;
                    nuevaLinea.CostoOpUni   = linea.CostoOpUni  ;
                    nuevaLinea.CostoUni     = linea.CostoUni    ;
                    nuevaLinea.Duracion     = linea.Duracion    ;
                    nuevaLinea.IdentifAviso = linea.IdentifAviso;
                    nuevaLinea.PautaId      = linea.PautaId     ;
                    nuevaLinea.Salida       = linea.Salida      ;

                    foreach (EstimadoDetDTO l in lineas)
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
                lineasTmp.Sort((x, y) => DateTime.Compare(x.Fecha, y.Fecha));

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


        private List<EstimadoDetDTO> PeriodosCopiar(DateTime fOrigenDesde, DateTime fOrigenHasta, DateTime fDestinoDesde, DateTime fDestinoHasta, List<EstimadoDetDTO> lineas)
        {
            List<EstimadoDetDTO> LineasOrigen        = lineas;
            List<EstimadoDetDTO> LineasSeleccionadas = new List<EstimadoDetDTO>();
            List<EstimadoDetDTO> LineasDestino       = lineas;

            FrecuenciaDTO frecuencia                  = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
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
                {
                    LineasSeleccionadas.Add(lineas[i]);
                }
            }

            //ORDENA LAS LINEAS SELECCIONADAS
            LineasSeleccionadas = LineasSeleccionadas.OrderBy(o => o.Dia).ThenBy(p => p.Hora).ThenBy(q => q.Salida).ToList();

            //CALCULO CANTIDAD DE DIAS EN EL MES
            int DiasEnMes = System.DateTime.DaysInMonth(LineasOrigen[0].Fecha.Year, LineasOrigen[0].Fecha.Month);

            //CALCULO CUAL ES EL DIA DE LA SEMANA DEL 1 DEL MES
            int DiaSemana = Convert.ToInt32(Convert.ToDateTime(
                LineasOrigen[0].Fecha.Year.ToString() + "-" +
                LineasOrigen[0].Fecha.Month.ToString("00") + "-" + "01").DayOfWeek);

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
                        DAO.EstimadoDetDAO odd = new DAO.EstimadoDetDAO();
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
                                EstimadoDetDTO newLine = new EstimadoDetDTO();
                                LastId++;

                                newLine.DatareaId  = 0                                ;
                                newLine.RecId      = LastId                           ;
                                newLine.Costo      = LineasSeleccionadas[j].Costo     ;
                                newLine.CostoOp    = LineasSeleccionadas[j].CostoOp   ;
                                newLine.CostoOpUni = LineasSeleccionadas[j].CostoOpUni;
                                newLine.CostoUni   = LineasSeleccionadas[j].CostoUni  ;
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

                                newLine.Duracion     = LineasSeleccionadas[j].Duracion      ;
                                newLine.Hora         = LineasSeleccionadas[j].Hora          ;
                                newLine.IdentifAviso = LineasSeleccionadas[j].IdentifAviso  ;
                                newLine.PautaId      = LineasSeleccionadas[j].PautaId       ;
                                newLine.Salida       = LineasSeleccionadas[j].Salida        ;

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

        protected void btnReemplazarAvisos_Click(object sender, EventArgs e)
        {
            string avisoOrigen   ;
            string avisoDestino  ;
            decimal salidas = 0  ;

            try
            {
                //Validaciones generales...
                if (ucIdentifAvisoOrigenReemplazar.SelectedValue != null)
                {
                    avisoOrigen = ucIdentifAvisoOrigenReemplazar.SelectedValue.ToString();
                }
                else
                {
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
                    Lineas = ReemplazarAvisosSeleccionados(avisoOrigen, avisoDestino, Lineas, salidas);
                }
                else if (opEditTodas.Checked)
                {
                    Lineas = ReemplazarAvisosTodos(avisoOrigen, avisoDestino, Lineas, salidas);
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


        private List<EstimadoDetDTO> ReemplazarAvisosPorPeriodo(DateTime fDesde, DateTime fHasta, string avisoOrigen, string avisoDestino, List<EstimadoDetDTO> lineas, decimal salidas)
        {
            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, fDesde, fHasta, "Ingresado");
        }

        private List<EstimadoDetDTO> ReemplazarAvisosTodos(string avisoOrigen, string avisoDestino, List<EstimadoDetDTO> lineas, decimal salidas)
        {
            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, DateTime.Now, DateTime.Now, "Todos");
        }

        private List<EstimadoDetDTO> ReemplazarAvisosSeleccionados(string avisoOrigen, string avisoDestino, List<EstimadoDetDTO> lineas, decimal salidas)
        {
            return ReemplazarAvisos(avisoOrigen, avisoDestino, lineas, salidas, DateTime.Now, DateTime.Now, "Seleccionado");
        }

        private List<EstimadoDetDTO> ReemplazarAvisos(string avisoOrigen, string avisoDestino, List<EstimadoDetDTO> lineas, decimal salidas, DateTime fDesde, DateTime fHasta, string Tipo = "Todo")
        {
            switch (Tipo)
            {
                case "Ingresado":
                    {
                        //Busco las línas con el periodo seleccionado, y cuyo aviso, sea igual al 'avisoOrigen',
                        //Para cada una de las líneas encontradas, reemplazo el aviso.

                        string FechaInicio = string.Empty;
                        string FechaFinal  = string.Empty;

                        FechaInicio       = fDesde.ToShortDateString();
                        FechaFinal        = fHasta.ToShortDateString();

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

                    Lineas = ProcesoHelper.ReemplazarAvisosPorPeriodo(fDesde, fHasta, avisoOrigen, avisoDestino, Lineas);
                }
                else if (opEditSeleccionados.Checked)
                {
                    Lineas = ProcesoHelper.ReemplazarAvisosSeleccionados(avisoOrigen, avisoDestino, gv.GetSelectedFieldValues("RecId"), Lineas);
                }
                else if (opEditTodas.Checked)
                {
                    Lineas = ProcesoHelper.ReemplazarAvisosTodos(avisoOrigen, avisoDestino, Lineas);
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

                    for (int i = 0; i <= lineas.Count - 1; i++)
                    {
                        if (lineas[i].RecId == ProxRecId)
                        {
                            if (lineas[i].Salida == 0)
                            {
                                TimeSpan tpHoraX = FormsHelper.ConvertToTimeSpan(Convert.ToDateTime(teHoraInicioModificar.Text));
                                DateTime tpHoraY = new DateTime(lineas[i].Fecha.Year, lineas[i].Fecha.Month, lineas[i].Fecha.Day, tpHoraX.Hours, tpHoraX.Minutes, 0);
                                decimal? tpX     = duracion;
                                tpHoraY          = tpHoraY.Add(new TimeSpan(0, 0, Convert.ToInt32(tpX)));
                                DateTime dt      = new DateTime(lineas[i].Fecha.Year, lineas[i].Fecha.Month, lineas[i].Fecha.Day, lineas[i].Hora.Hours, lineas[i].Hora.Minutes, 0);

                                if (tpHoraY >= dt && tpX > 0)
                                    throw new Exception("Solapamiento de horarios. Verifique.");
                            }
                        }
                    }
                }

                ///// FIN DE FUNCION DE RECALCULAR DURACION //


                if (Convert.ToDateTime(teHoraInicioModificar.Text) < Convert.ToDateTime(teHoraInicio.Text) ||
                    Convert.ToDateTime(teHoraInicioModificar.Text) > Convert.ToDateTime(teHoraFin.Text))
                {
                    throw new Exception("La hora no esta dentro del horario pautado.");
                }

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
                                    x.Duracion     = duracion;
                                    DateTime dt    = (DateTime)teHoraInicioModificar.DateTime;
                                    TimeSpan ts    = new TimeSpan(0, dt.Hour, dt.Minute, 0);
                                    x.Hora         = ts;
                                    x.Salida       = Convert.ToDecimal(spAvisoModifiSalidas.Value);
                                }
                            }
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
        }

        protected void gv_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            ASPxGridView gv = (ASPxGridView)sender;
            e.Cancel = true;

            //Do things...
            var aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", (string)e.NewValues["IdentifAviso"]));
            var lineas = Lineas;

            lineas.FindAll(x => x.RecId == (int)e.Keys[0]).ForEach(
                (linea) =>
                {
                    linea.IdentifAviso = aviso.IdentifAviso;
                    linea.Duracion = aviso != null ? aviso.Duracion : null;
                    linea.Salida = (decimal)e.NewValues["Salida"];
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

        #region "Orden de Publicidad"
        protected void EmitirOP()
        {
            EspacioContDTO myspace = GetEspacioContenido();

            string TipoOP = string.Empty;

            if (myspace != null)
            {
                TipoOP = "OP_" + myspace.FormatoOP;
            }
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

                    {
                        string Pauta = string.Empty;

                        Pauta = "OP ESTIMADO " + Estimado.Cabecera.AnoMes.ToString() + " V" + Estimado.Cabecera.Version + " " + Estimado.Cabecera.IdentifEspacio + " "; 

                        string filename = Pauta + System.DateTime.Now.Hour.ToString().PadLeft(2, '0') + System.DateTime.Now.Minute.ToString().PadLeft(2, '0') + System.DateTime.Now.Second.ToString().PadLeft(2,'0') + ".xlsx";

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
                                            string[] dirs       = Directory.GetFiles(@TextBox1.Text);
                                            dExiste             = false;
                                            filename            = @TextBox1.Text + filename;

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
                                                //ACA VA LA LLAMADA A LA FUNCION DEL HELPER
                                                // ES EL CORE, YA TENEMOS EL NOMBRE DE ARCHIVO Y LA RUTA

                                                EstimadoCabDTO Cabecera      = Estimados.Read(Estimado.Cabecera.RecId);
                                                List<EstimadoDetDTO> Detalle = Estimados.ReadAllLineas(Cabecera);
                                                List<EstimadoSKUDTO> SKUS    = Estimados.ReadAllSKUs(Cabecera);
                                                EspacioContDTO Espacio       = CRUDHelper.Read( string.Format("IdentifEspacio = '{0}'",Cabecera.IdentifEspacio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
                                                csOP_Helper Helper           = new csOP_Helper("ESTIMADO", "", Estimado.Cabecera.PautaId, Cabecera, Detalle, SKUS, Espacio,filename);

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

                    break;


                case "btnCost":

                    CostosDTO costos = Business.Ordenados.FindCosto(Estimado.Cabecera.IdentifEspacio, Convert.ToInt32(Estimado.Cabecera.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(Estimado.Cabecera.AnoMes.ToString().Substring(4, 2)));
                    Business.Estimados.CalcularCosto(Estimado.Cabecera, costos, Lineas, ((Accendo)this.Master).Usuario.UserName);
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

                default:
                    break;
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
                    e.Text       = string.Format("{0:HH:mm}", dt1);
                }

                if (dataColumn.FieldName.ToUpper() == "FECHA")
                {
                    DateTime dt = Convert.ToDateTime(e.Value.ToString());
                    e.Text      = string.Format("{0:dd/MM/yyyy}", dt);
                }
            }
        }

        #region "Botón Volver"
        protected void btnBack_Click(object sender, ImageClickEventArgs e)
        {
            Back();
        }

        private void Back()
        {
            Response.Redirect("EstimadoBusqueda.aspx");
        }
        #endregion

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            Back();
        }

        //HTTD
        protected void ASPxMenu1_ItemClick(object source, MenuItemEventArgs e)
        {

        }

        protected void btnCancelSKU_Click(object sender, EventArgs e)
        {
            trQuerySKU.Visible = false;
        }

        protected bool Validaciones(TimeSpan horaInicio, TimeSpan horaFin, IntervaloDTO intervalo, FrecuenciaDTO frecuenciaCab, AvisosDTO aviso, List<FrecuenciaDetDTO> frecuenciaDetalles)
        {
            bool retval              = false;
            decimal CantMinutosAviso = 0;

            List<EstimadoDetDTO> lineas        = Lineas;
            List<EstimadoDetDTO> preExistentes = new List<EstimadoDetDTO>();

            //TODO: VALIDACION 1: Validar que la hora de inicio de la inserción, no sea superior a la hora fin de la pauta.
            if (Convert.ToDateTime(teHoraInicioInsertar.Text) > Convert.ToDateTime(teHoraFin.Text) || Convert.ToDateTime(teHoraInicioInsertar.Text) < Convert.ToDateTime(teHoraInicio.Text))
            {
                lblErrorLineas.Text = "Las horas ingresadas no están dentro del rango de la cabecera.";
                retval              = false;
                return retval;
            }

            //TODO: VALIDACION 2: La duración del aviso excede el intervalo seleccionado para la Pauta. NO se generaron líneas.
            if (intervalo == null)
            {
                intervalo             = new IntervaloDTO();
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
                        retval              = false;

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
            else
            {
                for (int i = 0; i <= lineas.Count - 1; i++)
                    preExistentes.Add(lineas[i]);

                return ValidacionDiasContraPreexistente(horaInicio, horaFin, preExistentes, frecuenciaDetalles, aviso);
            }
        }

        protected void teHoraFinInsertar_ValueChanged(object sender, EventArgs e)
        {
            spSalidasInsertar.Enabled = (teHoraInicio.Text != teHoraFinInsertar.Text && teHoraFinInsertar.Text == "00:00");
        }

        protected void teHoraFinInsertar_DateChanged(object sender, EventArgs e)
        {
        }

        protected void btnVersions_Click(object sender, EventArgs e)
        {
        }

        protected bool ValidacionDiasContraPreexistente(TimeSpan horainicio, TimeSpan horafin, List<EstimadoDetDTO> Preexistentes, List<FrecuenciaDetDTO> frecuencia, AvisosDTO aviso)
        {
            bool retVal   = true;
            Preexistentes = Preexistentes.OrderBy(o => o.Dia).ThenBy(p => p.Hora).ThenBy(q => q.Salida).ToList();

            //Recorro las Preexistentes una vez por cada dia seleccionado
            for (int i = 0; i <= ceDiasInsertar.Items.Count - 1; i++)
            {
                if (ceDiasInsertar.Items[i].Selected == true)
                {
                    for (int h = 0; h <= Preexistentes.Count - 1; h++)
                    {
                        //Compara si existe el dia de semana seleccionado entre los preexistentes
                        if (Preexistentes[h].DiaSemana.ToString().Trim() == ceDiasInsertar.Items[i].ToString().Trim())
                        {
                            //TODO: VALIDACION 5: Comparo que la linea insertada no pise valores preexistentes
                            for (int j = 0; j <= Preexistentes.Count - 1; j++)
                            {
                                //Hora de inicio dentro del array de preexistentes
                                TimeSpan ts1i = new TimeSpan(0, Preexistentes[j].Hora.Hours, Preexistentes[j].Hora.Minutes, 0);

                                //Duracion del aviso dentro del array de preexistentes
                                TimeSpan ts1d = new TimeSpan(0, 0, Convert.ToInt32(Preexistentes[j].Duracion));

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

                                    retVal = false;

                                    break;
                                }

                                //Verificacion contra el registro siguiente
                                //Quiere decir que hay un registro previo despues del actual
                                if (j < Preexistentes.Count)
                                {
                                    //Datos del registro siguiente
                                    TimeSpan tsli                   = new TimeSpan(0, Preexistentes[j + 1].Hora.Hours, Preexistentes[j + 1].Hora.Minutes, 0);
                                    TimeSpan tsld                   = new TimeSpan(0, 0, Convert.ToInt32(Preexistentes[j + 1].Duracion));
                                    TimeSpan tslf                   = tsli.Add(tsld);
                                    TimeSpan RangoInicio            = ts1i;
                                    TimeSpan RangoFinal             = ts1f;
                                    TimeSpan HoraInicioAvisoNuevo   = ts2i;
                                    TimeSpan HoraFinAvisoNuevo      = ts2f;

                                    if (HoraFinAvisoNuevo.CompareTo(tsli) == 1)
                                    {
                                        lblErrorLineas.Text = "El aviso se superpone con otro";

                                        retVal = false;

                                        break;
                                    }
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

                                        retVal = false;

                                        break;
                                    }

                                    if (ts2i.CompareTo(ts1i) == 1)
                                    {
                                        //Comparo que la hora de inicio del aviso sea mayor que la hora de finalizacion del preexistente
                                        if (ts1f <= ts2i && aviso.Duracion != 0)
                                        {
                                            retVal = true;

                                            break;
                                        }
                                        else
                                        {
                                            lblErrorLineas.Text = "El aviso se superpone con otro";

                                            retVal = false;
                                        }
                                    }
                                }
                                else
                                {
                                    retVal = false;
                                }
                            }
                        }
                    }
                }
            }

            return retVal;
        }

        protected void btnVersions_Click1(object sender, EventArgs e)
        {

        }
    }
}