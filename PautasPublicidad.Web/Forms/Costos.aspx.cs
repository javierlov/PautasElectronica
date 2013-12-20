using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxGridView;
using PautasPublicidad.Business;
using DevExpress.Web.ASPxMenu;
using PautasPublicidad.DAO;
using PautasPublicidad.DTO;
using System.Data;
using System.Data.SqlClient;

namespace PautasPublicidad.Web.Forms
{
    public partial class Costos : System.Web.UI.Page
    {
        const string TODO      = "TODO"     ;
        const string DETALLADO = "DETALLADO";
        const string SEMANA    = "SEMANA"   ;
        const string MES       = "MES"      ;

        dynamic DAO;

        static CostosDTO miCosto;
        static string FechaDesde;
        static string FechaHasta;

        public List<DTO.CostosProveedorDTO> CostosProveedor
        {
            get
            {
                if (ViewState["CostosProveedor"] != null)
                    return (List<DTO.CostosProveedorDTO>)ViewState["CostosProveedor"];
                else
                    return new List<DTO.CostosProveedorDTO>();
            }
            set
            {
                ViewState.Add("CostosProveedor", value);
            }
        }

        public List<DTO.CostosFrecuenciaDTO> CostosFrecuencia
        {
            get
            {
                if (ViewState["CostosFrecuencia"] != null)
                    return (List<DTO.CostosFrecuenciaDTO>)ViewState["CostosFrecuencia"];
                else
                    return new List<DTO.CostosFrecuenciaDTO>();
            }
            set
            {
                ViewState.Add("CostosFrecuencia", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            //Invisibilizo los controles btnRefresh y btnAdd del ucIdentifEspacio0
            UcIdentifEspacio0.Controls[0].Controls[0].Controls[3].Visible = false;
            UcIdentifEspacio0.Controls[0].Controls[0].Controls[5].Visible = false;
            //

            DAO = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Costos);

            lblError.Text = string.Empty;
            lblErrorProveedor.Text = string.Empty;
            lblErrorFrecuencia.Text = string.Empty;

            if (!Page.IsPostBack && !Page.IsCallback)
            {
                trMsg.Visible = false;
                trAbm.Visible = false;
                
                //Inicializo los controles del ABM
                trTipoHorario.Visible = false;
                trFrecuencia.Visible = false;
                trDia.Visible = false;
                trDiaSemana.Visible = false;
                trHoraDesde.Visible = false;
                trHoraHasta.Visible = false;

                //Inicializo las 3 grillas.
                FormsHelper.BuildColumnsByEntity(BusinessMapper.eEntities.Costos, gv);
                FormsHelper.BuildColumnsByEntity(BusinessMapper.eEntities.Costos, gvVersiones);
                
                gvABMFrecuencia.Columns.Add(new GridViewDataColumn("Dia", "Día"));
                gvABMFrecuencia.Columns.Add(new GridViewDataComboBoxColumn() { Caption = "Día Semana", FieldName = "DiaSemana" });
                gvABMFrecuencia.Columns.Add(new GridViewDataTimeEditColumn() { Caption = "Hora Desde", FieldName = "HoraDesde" });
                gvABMFrecuencia.Columns.Add(new GridViewDataTimeEditColumn() { Caption = "Hora Hasta", FieldName = "HoraHasta" });
                gvABMFrecuencia.Columns.Add(new GridViewDataColumn("Costo", "Costo"));

                FormsHelper.FillDias((gvABMFrecuencia.Columns["DiaSemana"] as GridViewDataComboBoxColumn).PropertiesComboBox.Items);
                
                gvABMProveedor.Columns.Add(FormsHelper.BuildComboColumn("Proveedor", "IdentifProv", BusinessMapper.eEntities.Proveedor));
                gvABMProveedor.Columns.Add(FormsHelper.BuildComboColumn("Categoría de Costo", "Categoria", "Directo", "DIRECTO", "Indirecto", "INDIRECTO"));
                gvABMProveedor.Columns.Add(new GridViewDataCheckColumn() { Caption = "Incluido en Orden Publicidad", FieldName = "IncluidoOP" });
                gvABMProveedor.Columns.Add(new GridViewDataCheckColumn() { Caption = "Estimado", FieldName = "Estimado" });
                gvABMProveedor.Columns.Add(FormsHelper.BuildComboColumn("Tipo de Costo", "TipoCosto", "Fijo Mensual", "FIJO_MENSUAL", "Segundo Fijo", "SEGUNDO_FIJO", "Unidad Pautada", "UNIDAD_PAUTADA"));
                gvABMProveedor.Columns.Add(FormsHelper.BuildComboColumn("Moneda", "IdentifMon", BusinessMapper.eEntities.Monedas));
                gvABMProveedor.Columns.Add(new GridViewDataColumn("GrossingUp", "Grossing Up"));
                gvABMProveedor.Columns.Add(new GridViewDataColumn("Costo", "Costo"));
                
                FormsHelper.InicializarPropsGrilla(gv);
                FormsHelper.InicializarPropsGrilla(gvABMFrecuencia);
                FormsHelper.InicializarPropsGrilla(gvABMProveedor);
                FormsHelper.InicializarPropsGrilla(gvVersiones);

                gvABMFrecuencia.Settings.ShowGroupPanel = false;
                gvABMProveedor.Settings.ShowGroupPanel = false;
                gvVersiones.Settings.ShowGroupPanel = false;
            }

            ucIdentifEspacio.Inicializar(BusinessMapper.eEntities.EspacioCont);

            UcIdentifEspacio0.Inicializar(BusinessMapper.eEntities.EspacioCont);
            UcIdentifEspacio0.ComboBox.AutoPostBack = true;
            UcIdentifEspacio0.ComboBox.SelectedIndexChanged += new EventHandler(IdentifEspacio0_SelectedIndexChanged);

            ucIdentifFrecuencia.Inicializar(BusinessMapper.eEntities.Frecuencia);
            ucIdentifMon.Inicializar(BusinessMapper.eEntities.Monedas);
            ucIdentifProv.Inicializar(BusinessMapper.eEntities.Proveedor);

            ucIdentifEspacio.ComboBox.AutoPostBack = true;
            ucIdentifEspacio.ComboBox.SelectedIndexChanged += new EventHandler(IdentifEspacio_SelectedIndexChanged);
            ucIdentifFrecuencia.ComboBox.AutoPostBack = true;
            ucIdentifFrecuencia.ComboBox.SelectedIndexChanged += new EventHandler(IdentifFrecuencia_SelectedIndexChanged);
            rbFrecuencia.AutoPostBack = true;
            rbFrecuencia.SelectedIndexChanged += new EventHandler(rbFrecuencia_SelectedIndexChanged);
            rbHorario.AutoPostBack = true;
            rbHorario.SelectedIndexChanged += new EventHandler(rbHorario_SelectedIndexChanged);

            gvABMFrecuencia.KeyFieldName = "RecId";
            gvABMProveedor.KeyFieldName  = "RecId";
            gvVersiones.KeyFieldName     = "RecId";

            ASPxMenu1.ItemClick += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);

            RefreshGrid(gv);
            RefreshAbmGrid(gvABMProveedor);
            RefreshAbmGrid(gvABMFrecuencia);
            RefreshGrid(gvVersiones);
        }

        #region Eventos de Controles (que muestran u ocultan segun valores seleccionados)

        void IdentifEspacio_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucEspacioChanged();
        }

        void IdentifFrecuencia_SelectedIndexChanged(object sender, EventArgs e)
        {
            ucFrecuenciaChanged();
        }

        void rbFrecuencia_SelectedIndexChanged(object sender, EventArgs e)
        {
            rbFrecuenciaChanged();
        }

        void rbHorario_SelectedIndexChanged(object sender, EventArgs e)
        {
            rbHorarioChanged();
        }

        private void ucEspacio0Changed()
        { 
        }

        private void ucEspacioChanged()
        {
            trTipoHorario.Visible = false;
            rbHorario.SelectedItem = rbHorario.Items.FindByValue(TODO);
            
            if (ucIdentifEspacio.SelectedValue != null)
            {
                DTO.EspacioContDTO espacio = CRUDHelper.Read(string.Format("IdentifEspacio = '{0}'", ucIdentifEspacio.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

                trTipoHorario.Visible = (espacio.HoraFin.HasValue && espacio.HoraFin.Value.TotalMinutes > 0);

                if (!(espacio.HoraFin.HasValue && espacio.HoraFin.Value.TotalMinutes > 0))
                    rbHorario.SelectedItem = rbHorario.Items.FindByValue(TODO);
            }

            rbHorarioChanged();
        }

        private void rbHorarioChanged()
        {
            trHoraDesde.Visible = rbHorario.SelectedItem != null ? rbHorario.SelectedItem.Value.ToString() == DETALLADO : false;
            trHoraHasta.Visible = rbHorario.SelectedItem != null ? rbHorario.SelectedItem.Value.ToString() == DETALLADO : false;

            if (!trHoraDesde.Visible) teHoraDesde.Value = null;

            if (!trHoraHasta.Visible) teHoraHasta.Value = null;
        }

        private void rbFrecuenciaChanged()
        {
            trFrecuencia.Visible = (rbFrecuencia.SelectedItem.Value.ToString() == DETALLADO);

            if (!(rbFrecuencia.SelectedItem.Value.ToString() == DETALLADO))
                ucIdentifFrecuencia.SelectedValue = null;

            ucFrecuenciaChanged();
        }

        private void ucFrecuenciaChanged()
        {
            if (ucIdentifFrecuencia.SelectedValue != null)
            {
                DTO.FrecuenciaDTO frecuencia = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", ucIdentifFrecuencia.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));

                trDia.Visible       = (frecuencia.SemMes.Trim().ToUpper() == MES);
                trDiaSemana.Visible = (frecuencia.SemMes.Trim().ToUpper() == SEMANA);
            }
            else
            {
                trDia.Visible       = false;
                trDiaSemana.Visible = false;
            }

            if (!trDia.Visible)         //Si no se muestra, va en NULL.
                spDia.Value = null;

            if (!trDiaSemana.Visible)   //Si no se muestra, va en NULL.
                clDiaSemana.SelectedItem = null;
        }

        #endregion

        private void RefreshGrid(ASPxGridView gridView)
        {
            if (gridView.ID == "gv")
                gridView.DataSource = Business.Costos.ReadAll("");
            else
                if (pnlVersiones.Attributes["RecId"] != null)
                    gvVersiones.DataSource = Business.Costos.ReadAllVersiones(Business.Costos.Read(Convert.ToInt32(pnlVersiones.Attributes["RecId"])));


            gridView.DataBind();
        }

        private void RefreshAbmGrid(ASPxGridView gvABM)
        {
            if (gvABM.ID == "gvABMFrecuencia")
                gvABM.DataSource = CostosFrecuencia;
            else
                gvABM.DataSource = CostosProveedor;
            
            gvABM.DataBind();
        }

        protected void detailGrid_DataSelect(object sender, EventArgs e)
        {
            ASPxGridView gvDetail = (ASPxGridView)sender;

            if (gvDetail.ID.ToUpper() == "detailGridFrecuencia".ToUpper() || gvDetail.ID.ToUpper() == "detailGridFrecuenciaVer".ToUpper())
            {
                gvDetail.Columns.Clear();

                gvDetail.Columns.Add(new GridViewDataColumn("Dia", "Día"));
                gvDetail.Columns.Add(new GridViewDataComboBoxColumn() { Caption = "Día Semana", FieldName = "DiaSemana" });
                gvDetail.Columns.Add(new GridViewDataTimeEditColumn() { Caption = "Hora Desde", FieldName = "HoraDesde" });
                gvDetail.Columns.Add(new GridViewDataTimeEditColumn() { Caption = "Hora Hasta", FieldName = "HoraHasta" });
                gvDetail.Columns.Add(new GridViewDataColumn("Costo", "Costo"));

                FormsHelper.FillDias((gvDetail.Columns["DiaSemana"] as GridViewDataComboBoxColumn).PropertiesComboBox.Items);

                if (gvDetail.ID.ToUpper() == "detailGridFrecuencia".ToUpper())
                {
                    gvDetail.DataSource = Business.Costos.ReadAllFrecuencia((string)gvDetail.GetMasterRowFieldValues("IdentifEspacio"),
                                                                            (DateTime)gvDetail.GetMasterRowFieldValues("VigDesde"),
                                                                            (DateTime)gvDetail.GetMasterRowFieldValues("VigHasta"));
                }
                else
                {
                    gvDetail.DataSource = Business.Costos.ReadAllFrecuenciaVersiones((string)gvDetail.GetMasterRowFieldValues("IdentifEspacio"),
                                                                                     (DateTime)gvDetail.GetMasterRowFieldValues("VigDesde"),
                                                                                     (DateTime)gvDetail.GetMasterRowFieldValues("VigHasta"),
                                                                                     (decimal)gvDetail.GetMasterRowFieldValues("Version"));
                }
            }
            
            if (gvDetail.ID.ToUpper() == "detailGridProveedor".ToUpper() || gvDetail.ID.ToUpper() == "detailGridProveedorVer".ToUpper())
            {
                gvDetail.Columns.Clear();

                gvDetail.Columns.Add(FormsHelper.BuildComboColumn("Proveedor", "IdentifProv", BusinessMapper.eEntities.Proveedor));
                gvDetail.Columns.Add(FormsHelper.BuildComboColumn("Categoría de Costo", "Categoria", "Directo", "DIRECTO", "Indirecto", "INDIRECTO"));
                gvDetail.Columns.Add(new GridViewDataCheckColumn() { Caption = "Incluido en Orden Publicidad", FieldName = "IncluidoOP" });
                gvDetail.Columns.Add(new GridViewDataCheckColumn() { Caption = "Estimado", FieldName = "Estimado" });
                gvDetail.Columns.Add(FormsHelper.BuildComboColumn("Tipo de Costo", "TipoCosto", "Fijo Mensual", "FIJO_MENSUAL", "Segundo Fijo", "SEGUNDO_FIJO", "Unidad Pautada", "UNIDAD_PAUTADA"));
                gvDetail.Columns.Add(FormsHelper.BuildComboColumn("Moneda", "IdentifMon", BusinessMapper.eEntities.Monedas));
                gvDetail.Columns.Add(new GridViewDataColumn("GrossingUp", "Grossing Up"));
                gvDetail.Columns.Add(new GridViewDataColumn("Costo", "Costo"));
                
                if (gvDetail.ID.ToUpper() == "detailGridProveedor".ToUpper())
                    gvDetail.DataSource = Business.Costos.ReadAllProveedor((string)gvDetail.GetMasterRowFieldValues("IdentifEspacio"), (DateTime)gvDetail.GetMasterRowFieldValues("VigDesde"), (DateTime)gvDetail.GetMasterRowFieldValues("VigHasta"));

                else
                    gvDetail.DataSource = Business.Costos.ReadAllProveedorVersiones((string)gvDetail.GetMasterRowFieldValues("IdentifEspacio"), (DateTime)gvDetail.GetMasterRowFieldValues("VigDesde"), (DateTime)gvDetail.GetMasterRowFieldValues("VigHasta"), (decimal)gvDetail.GetMasterRowFieldValues("Version"));
            }
        }

        #region ABM

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            pnlControls.Visible = false;
        }

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (ASPxPageControl1.ActiveTabPage == ASPxPageControl1.TabPages[3])
                { 
                    if(ValidCostCopy(null,out err))
                    {
                        var entidad = miCosto;

                        ///// Vuelve a grabar la fecha - por las dudas que no haya aceptado la propuesta /////
                        miCosto.VigDesde = Convert.ToDateTime(deVigDesde0.Text);
                        miCosto.VigHasta = Convert.ToDateTime(deVigHasta0.Text);

                        Business.Costos.Create(entidad,CostosFrecuencia,CostosProveedor);

                        RefreshGrid(gv);
                    }
                    else
                    {
                    throw new Exception(err);
                    }
                }
                else
                {
                    if (Valid(null, out err))
                    {
                        var entity = new DTO.CostosDTO();
                        FormsHelper.FillEntity(tblCosto, entity);

                        Business.Costos.Create(entity, CostosFrecuencia, CostosProveedor);

                        pnlControls.Visible = false;
                        RefreshGrid(gv);
                    }
                    else
                    {
                        throw new Exception(err);
                    }
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        private bool Valid(int? recId, out string err)
        {
            err = string.Empty;

            if (ucIdentifEspacio.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Espacio de Contenido'.";
                return false;
            }
            
            if (deVigDesde.Value == null || deVigHasta.Value == null)
            {
                err = "Debe seleccionar 'Vigencia Desde' y 'Vigencia Hasta'.";
                return false;
            }

            if (rbFrecuencia.SelectedIndex == -1)
            {
                err = "Debe seleccionar un 'Tipo de Frecuencia'.";
                return false;
            }

            if (trFrecuencia.Visible && ucIdentifFrecuencia.SelectedValue == null)
            {
                err = "Debe seleccionar una 'Frecuencia'.";
                return false;

            }
            if (trTipoHorario.Visible && rbHorario.SelectedIndex == -1)
            {
                err = "Debe seleccionar un 'Tipo de Horario'.";
                return false;
            }

            if (CostosProveedor == null || CostosProveedor.Count == 0)
            {
                err = "Debe ingresar algún 'Costo por Proveedor'.";
                return false;
            }

            if (recId == null)
            {
                var costosEnPeriodo = Business.Costos.ReadAll(string.Format("VigHasta >= '{0}' AND VigDesde <= '{1}' AND IdentifEspacio = '{2}'",
                    deVigDesde.Date.ToString("yyyyMMdd"), deVigHasta.Date.ToString("yyyyMMdd"), ucIdentifEspacio.SelectedValue.ToString()));

                if (costosEnPeriodo.Count > 0)
                {
                    err = "Ya existe un costo con vigencia en el período seleccionado.";
                    return false;
                }
            }
            else
            {
                var costosEnPeriodo = Business.Costos.ReadAll(string.Format("VigHasta >= '{0}' AND VigDesde <= '{1}' AND IdentifEspacio = '{2}' AND RecId <> {3}",
                    deVigDesde.Date.ToString("yyyyMMdd"), deVigHasta.Date.ToString("yyyyMMdd"), ucIdentifEspacio.SelectedValue.ToString(), recId.Value));

                if (costosEnPeriodo.Count > 0)
                {
                    err = "Ya existe un costo con vigencia en el período seleccionado.";
                    return false;
                }
            }

            return true;
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            string err = string.Empty;
            
            try
            {
                int recId = Convert.ToInt32(pnlControls.Attributes["RecId"]);

                if (Valid(recId, out err))
                {
                    var entity = Business.Costos.Read(recId);
                    FormsHelper.FillEntity(tblCosto, entity);

                    Business.Costos.Update(entity, CostosFrecuencia, CostosProveedor);

                    pnlControls.Visible = false;

                    RefreshGrid(gv);
                }
                else
                {
                    throw new Exception(err);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (pnlControls.Attributes["RecId"] != null)
                {
                    Business.Costos.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]));

                    pnlControls.Visible = false;

                    RefreshGrid(gv);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        #endregion

        protected void ASPxMenu1_ItemClick(object source, MenuItemEventArgs e)
        {
            try
            {

                ASPxPageControl1.TabPages[3].ClientEnabled = e.Item.Name == "btnAdd";

                switch (e.Item.Name)
                {
                    case "btnAdd":

                        if (ASPxPageControl1.ActiveTabPage.Index == 3)
                        {
                            //for costosCopy only
                            UcIdentifEspacio0.ComboBox.Text = "";
                            deVigDesde0.Text                = "";
                            deVigHasta0.Text                = "";
                        }


                        FormsHelper.ClearControls(tblCosto, new DTO.CostosDTO());
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);

                        CostosFrecuencia = new List<DTO.CostosFrecuenciaDTO>();
                        CostosProveedor  = new List<DTO.CostosProveedorDTO>();

                        RefreshAbmGrid(gvABMFrecuencia);
                        RefreshAbmGrid(gvABMProveedor);

                        pnlControls.Visible = true;
                        pnlControls.HeaderText = "Agregar Registro";
                        break;

                    case "btnEdit":

                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ClearControls(tblCosto, new DTO.CostosDTO());
                            var entity = Business.Costos.Read(FormsHelper.GetSelectedId(gv).Value);
                            FormsHelper.FillControls(entity, tblCosto);
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);

                            CostosFrecuencia = Business.Costos.ReadAllFrecuencia(entity);
                            CostosProveedor = Business.Costos.ReadAllProveedor(entity);

                            ucEspacioChanged();
                            rbFrecuenciaChanged();

                            RefreshAbmGrid(gvABMFrecuencia);
                            RefreshAbmGrid(gvABMProveedor);

                            pnlControls.Attributes.Add("RecId", entity.RecId.ToString());
                            pnlControls.Visible    = true;
                            pnlControls.HeaderText = "Modificar Registro";
                        }
                        else
                        {
                            pnlControls.Visible = false;
                        }
                        break;

                    case "btnDelete":

                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Delete);
                            pnlControls.Attributes.Add("RecId", FormsHelper.GetSelectedId(gv).ToString());
                            pnlControls.Visible    = true;
                            pnlControls.HeaderText = "Eliminar Registros";
                        }
                        else
                        {
                            pnlControls.Visible = false;
                        }
                        break;

                    case "btnCommit":

                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            pnlCommit.Attributes.Add("RecId", FormsHelper.GetSelectedId(gv).Value.ToString());
                            pnlControls.Visible = false;
                            pnlCommit.Visible   = true;
                        }
                        else
                        {
                            pnlCommit.Visible = false;
                        }
                        break;

                    case "btnQuery":

                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            pnlVersiones.Attributes.Add("RecId", FormsHelper.GetSelectedId(gv).Value.ToString());
                            RefreshGrid(gvVersiones);
                            pnlVersiones.Visible = true;
                        }
                        else
                        {
                            pnlVersiones.Visible = false;
                        }                        
                        break;

                    case "btnExport":
                    case "btnExportXls":

                        if (ASPxGridViewExporter1 != null)
                            ASPxGridViewExporter1.WriteXlsToResponse();
                        break;

                    case "btnExportPdf":
                        if (ASPxGridViewExporter1 != null)
                            ASPxGridViewExporter1.WritePdfToResponse();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        #region CostoProveedor

        protected void btnAddProveedor_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
            
                if (ValidProveedor(out err))
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.CostosProveedorDTO> aux = CostosProveedor;

                    var costosProveedor = new DTO.CostosProveedorDTO();

                    FormsHelper.FillEntity(tblProveedor, costosProveedor);
                    costosProveedor.Costo = spCostoProveedor.Number;

                    costosProveedor.RecId = aux.Count;
                    aux.Add(costosProveedor);

                    CostosProveedor = aux;
                    RefreshAbmGrid(gvABMProveedor);

                    //Limpio controles.
                    ucIdentifProv.SelectedValue = null;
                    rbCategoria.SelectedItem    = null;
                    cbIncluidoOP.Checked        = false;
                    cbEstimado.Checked          = false;
                    rbTipoCosto.SelectedItem    = null;
                    ucIdentifMon.SelectedValue  = null;
                    spGrossingUp.Value          = null;
                    spCostoProveedor.Value      = 0;
                }
                else
                {
                    throw new Exception(err);
                }
            }
            catch (Exception ex)
            {
                MsgErrorProveedor(ex);
            }
        }

        protected void btnDeleteProveedor_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormsHelper.GetSelectedId(gvABMProveedor) != null)
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.CostosProveedorDTO> aux = new List<DTO.CostosProveedorDTO>();

                    //Creo una nueva coleccion con todos los productos menos el eliminado, y la guardo en el Viewstate.
                    foreach (var proveedor in CostosProveedor)
                        if (Convert.ToInt32(FormsHelper.GetSelectedId(gvABMProveedor)) != proveedor.RecId)
                            aux.Add(proveedor);

                    CostosProveedor = aux;
                    RefreshAbmGrid(gvABMProveedor);
                }
            }
            catch (Exception ex)
            {
                MsgErrorProveedor(ex);
            }
        }

        private void MsgErrorProveedor(Exception ex)
        {
            if (ex.Message.ToLower().Contains("Violation of UNIQUE KEY constraint".ToLower()))
                lblErrorProveedor.Text = "No puede ingresar un nuevo registro con su clave duplicada.";
            else if (ex.Message.ToLower().Contains("FK".ToLower()))
                lblErrorProveedor.Text = "No puede modificar/eliminar la clave de este registro, ya que se encuentra relacionado a otra entidad.";
            else
                lblErrorProveedor.Text = ex.Message;
        }
        
        private bool ValidProveedor(out string err)
        {
            err = string.Empty;

            if (Convert.ToInt32(spGrossingUp.Value) <= 0)
                {
                    err = "Debe ingresar Grossing Up mayor a cero";
                return false;
                }

            if (ucIdentifProv.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Proveedor'.";
                return false;
            }

            if (rbCategoria.SelectedIndex == -1)
            {
                err = "Debe seleccionar una 'Categoría de Costo'.";
                return false;
            }

            if (rbTipoCosto.SelectedIndex == -1)
            {
                err = "Debe seleccionar un 'Tipo de Costo'.";
                return false;
            }

            if (ucIdentifMon.SelectedValue == null)
            {
                err = "Debe seleccionar una 'Moneda'.";
                return false;
            }

            if (spCostoProveedor.Number <= 0 && cbGeneraOC.Checked) //
            {
                err = "El 'Costo' debe ser mayor a cero.";
                return false;
            }

            foreach (var costoProveedor in CostosProveedor)
            {
                if (costoProveedor.IdentifProv == Convert.ToString(ucIdentifProv.SelectedValue))
                {
                    err = "No puede ingresar dos veces el mismo proveedor.";
                    return false;
                }
                
            }
            return true;
        }
        
        #endregion

        #region CostoFrecuencia

        protected void btnAddFrecuencia_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (ValidFrecuencia(out err))
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.CostosFrecuenciaDTO> aux = CostosFrecuencia;

                    if (trDiaSemana.Visible)
                    {
                        for (int i = 0; i < clDiaSemana.SelectedItems.Count; i++)
                        {

                            var costosFrecuencia = new DTO.CostosFrecuenciaDTO();

                            FormsHelper.FillEntity(tblFrecuencia, costosFrecuencia);

                            costosFrecuencia.DiaSemana = Convert.ToString(clDiaSemana.SelectedValues[i]);
                            costosFrecuencia.Costo     = spCostoFrecuencia.Number;
                            costosFrecuencia.RecId     = aux.Count;

                            aux.Add(costosFrecuencia);
                        }
                    }
                    else
                    {
                        var costosFrecuencia = new DTO.CostosFrecuenciaDTO();

                        FormsHelper.FillEntity(tblFrecuencia, costosFrecuencia);

                        costosFrecuencia.Costo = spCostoFrecuencia.Number;
                        costosFrecuencia.RecId = aux.Count;

                        aux.Add(costosFrecuencia);
                    }

                    CostosFrecuencia = aux;
                    RefreshAbmGrid(gvABMFrecuencia);

                    //Limpio controles.
                    spDia.Value = null;

                    for (int i = 0; i < clDiaSemana.Items.Count; i++)
                        clDiaSemana.Items[i].Selected = false;
                        clDiaSemana.SelectedItem      = null;                    
                        teHoraDesde.Value             = null;
                        teHoraHasta.Value             = null;
                        spCostoFrecuencia.Value       = null;
                }
                else
                {
                    throw new Exception(err);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorFrecuencia, ex);
            }
        }

        protected void btnDeleteFrecuencia_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormsHelper.GetSelectedId(gvABMFrecuencia) != null)
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.CostosFrecuenciaDTO> aux = new List<DTO.CostosFrecuenciaDTO>();

                    //Creo una nueva coleccion con todos los productos menos el eliminado, y la guardo en el Viewstate.
                    foreach (var frecuencia in CostosFrecuencia)
                        if (Convert.ToInt32(FormsHelper.GetSelectedId(gvABMFrecuencia)) != frecuencia.RecId)
                            aux.Add(frecuencia);

                    CostosFrecuencia = aux;
                    RefreshAbmGrid(gvABMFrecuencia);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorFrecuencia, ex);
            }
        }
      
        private bool ValidFrecuencia(out string err)
        {
            err = string.Empty;

            if (trDia.Visible && spDia.Number < 1)
            {
                err = "Debe ingresar un 'Día' válido del més.";
                return false;
            }

            if (trDiaSemana.Visible && clDiaSemana.SelectedItems.Count == 0)            
            {
                err = "Debe seleccionar un 'Día de la Semana'.";
                return false;
            }

            if (trHoraDesde.Visible && teHoraDesde.Value == null)
            {
                err = "Debe seleccionar una 'Hora Desde'.";
                return false;
            }

            if (trHoraHasta.Visible && teHoraHasta.Value == null)
            {
                err = "Debe seleccionar una 'Hora Hasta'.";
                return false;
            }

            if (trHoraHasta.Visible && trHoraDesde.Visible && teHoraHasta.DateTime < teHoraDesde.DateTime)
            {
                err = "La 'Hora Hasta' debe ser mayor a la 'Hora Desde'.";
                return false;
            }

            if (spCostoFrecuencia.Number <= 0)
            {
                err = "El 'Costo' debe ser mayor a cero.";
                return false;
            }
            return true;
        }

        #endregion

        protected void btnCancel2_Click(object sender, EventArgs e)
        {
            pnlCommit.Visible = false;
        }

        protected void btnCommit_Click(object sender, EventArgs e)
        {
            try
            {
                Business.Costos.Commit(Convert.ToInt32(pnlCommit.Attributes["RecId"]), ((Accendo)this.Master).Usuario);

                pnlCommit.Visible = false;
                RefreshGrid(gv);
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblMsgCommit, ex);
            }
        }

        protected void btnCancel3_Click(object sender, EventArgs e)
        {
            pnlVersiones.Attributes.Remove("RecId");
            pnlVersiones.Visible = false;
        }

        #region CopiarCostos

        private bool ValidCostCopy(int? recId, out string err)
        {
            err = string.Empty;

            if (UcIdentifEspacio0.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Espacio de Contenido'.";
                return false;
            }

            if (Convert.ToDateTime(deVigDesde0.Text) < Convert.ToDateTime(FechaDesde))
            {
                err = "La Fecha Desde debe ser mayor que Fecha Hasta del último costo.";
                return false;
            }

            if (Convert.ToDateTime(deVigHasta0.Text) < Convert.ToDateTime(deVigDesde0.Text))
            {
                err = "La Fecha Desde debe ser mayor que la Fecha Hasta.";
                return false;
            }

            return true;
        }

        void IdentifEspacio0_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sWhere = "IdentifEspacio = '" + this.UcIdentifEspacio0.SelectedValue + "' ORDER BY VIGHASTA DESC";
            string oldFDesde = string.Empty;
            string oldFHasta = string.Empty;

            var oCosto      = DTOHelper.InstanciarObjetoPorNombreDeTabla("Costos");
            var oFrecuencia = DTOHelper.InstanciarObjetoPorNombreDeTabla("CostosFrecuencia");
            var oProveedor  = DTOHelper.InstanciarObjetoPorNombreDeTabla("CostosProveedor");

            DAOBase<DTO.CostosDTO> M = new CostosDAO();
            CostosDTO costo = M.ReadUnique(sWhere);

            FechaDesde = costo.VigHasta.AddDays(1).ToShortDateString();
            oldFDesde = costo.VigDesde.ToShortDateString();

            if (Convert.ToDateTime(FechaDesde).Month < 12)
                FechaHasta = new DateTime(Convert.ToDateTime(FechaDesde).Year, Convert.ToDateTime(FechaDesde).Month + 1, 1).AddDays(-1).ToShortDateString();
            else
                FechaHasta = new DateTime(Convert.ToDateTime(FechaDesde).Year,12,31).ToShortDateString();
           
            deVigDesde0.Text = FechaDesde;
            deVigHasta0.Text = FechaHasta;

            miCosto = costo;

            sWhere = "RECID > 0 ORDER BY RECID DESC";
            costo = M.ReadUnique(sWhere);

            miCosto.RecId         = costo.RecId + 1;
            miCosto.Confirmado    = "";
            miCosto.FecConfirmado = null;
            miCosto.VigDesde      = Convert.ToDateTime(FechaDesde);
            miCosto.VigHasta      = Convert.ToDateTime(FechaHasta);
            miCosto.Version       = null;

            sWhere = "IdentifEspacio = '" + this.UcIdentifEspacio0.SelectedValue + "' AND  YEAR(VIGDESDE)  = " + Convert.ToDateTime(oldFDesde).Year +
                                                                                    " AND MONTH(VIGDESDE)  = " + Convert.ToDateTime(oldFDesde).Month +
                                                                                    " AND   DAY(VIGDESDE)  = " + Convert.ToDateTime(oldFDesde).Day;
            DAOBase<DTO.CostosProveedorDTO> O = new CostosProveedorDAO();

            CostosProveedor = O.ReadAll(sWhere);

            DAOBase<DTO.CostosFrecuenciaDTO> N = new CostosFrecuenciaDAO();

            CostosFrecuencia = N.ReadAll(sWhere);

        }

        #endregion

        protected void deVigDesde0_DateChanged(object sender, EventArgs e) 
        {
        }

        protected void deVigHasta0_DateChanged(object sender, EventArgs e) 
        {
        }

    }
}