using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxGridView;
using PautasPublicidad.Business;
using DevExpress.Web.ASPxMenu;
using DevExpress.Web.ASPxEditors;

namespace PautasPublicidad.Web.Forms
{
    public partial class Avisos : System.Web.UI.Page
    {
        dynamic DAO;

        public List<DTO.AvisosIdAtenDTO> Atencion
        {
            get
            {
                if (ViewState["Atencion"] != null)
                    return (List<DTO.AvisosIdAtenDTO>)ViewState["Atencion"];
                else
                    return new List<DTO.AvisosIdAtenDTO>();
            }
            set
            {
                ViewState.Add("Atencion", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            DAO = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Avisos);
            
            lblError.Text = string.Empty;
            lblErrorProducto.Text = string.Empty;
            
            if (!Page.IsPostBack && !Page.IsCallback)
            {
                trMsg.Visible = false;
                trAbm.Visible = false;

                gvABM.SettingsBehavior.AllowSelectByRowClick = true;
                gvABM.SettingsBehavior.AllowSelectSingleRowOnly = true;
                FormsHelper.BuildColumnsByEntity(BusinessMapper.eEntities.Avisos, gv);
                FormsHelper.InicializarPropsGrilla(gv);

                cbIdentifIdentAte.CallbackPageSize = 10;
                cbIdentifIdentAte.EnableCallbackMode = true;
                cbIdentifIdentAte.IncrementalFilteringMode = DevExpress.Web.ASPxEditors.IncrementalFilteringMode.Contains;
                cbIdentifIdentAte.ValueField = "IdentifIdentAte";
                cbIdentifIdentAte.Columns.Add("IdentifIdentAte", "Codigo");
                cbIdentifIdentAte.Columns.Add("Name", "Nombre");
                cbIdentifIdentAte.Columns.Add("Asignado", "Asignado");
            }

            ucIdentifEspacio.Inicializar(BusinessMapper.eEntities.EspacioCont);
            ucIdentifFormAviso.Inicializar(BusinessMapper.eEntities.FormAviso);
            ucIdentifPieza.Inicializar(BusinessMapper.eEntities.PiezasArte);

            ucIdentifPieza.ComboBox.AutoPostBack = true;
            ucIdentifPieza.ComboBox.SelectedIndexChanged += new EventHandler(Pieza_SelectedIndexChanged);

            gvABM.KeyFieldName = "RecId";
            ASPxMenu1.ItemClick += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);

            btnAdd.Click += new EventHandler(btnAdd_Click);
            btnCancel.Click += new EventHandler(btnCancel_Click);
            btnDelete.Click += new EventHandler(btnDelete_Click);
            btnSave.Click += new EventHandler(btnSave_Click);
            btnAddAtencion.Click += new EventHandler(btnAddAtencion_Click);
            btnDeleteAtencion.Click += new EventHandler(btnDeleteAtencion_Click);

            RefreshGrid(gv);
            RefreshAbmGrid(gvABM);
        }

        void Pieza_SelectedIndexChanged(object sender, EventArgs e)
        {
            DTO.PiezasArteDTO pieza = CRUDHelper.Read(
                string.Format("IdentifPieza = '{0}'", ucIdentifPieza.SelectedValue),
                BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

            if ((pieza != null && pieza.Duracion.HasValue && pieza.Duracion > 0))
            {
                trDuracion.Visible = true;
                spDuracion.Value = pieza.Duracion;
            }
            else
            {
                trDuracion.Visible = false;
                spDuracion.Value = null;
            }
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = CRUDHelper.ReadAll("", DAO);
            gv.DataBind();
        }

        private void RefreshAbmGrid(ASPxGridView gvABM)
        {
            gvABM.DataSource = Atencion;
            gvABM.DataBind();
        }

        protected void detailGrid_DataSelect(object sender, EventArgs e)
        {
            ASPxGridView gvDetail = (ASPxGridView)sender;
            gvDetail.DataSource = Business.Avisos.ReadAllAtencion((string)gvDetail.GetMasterRowFieldValues("IdentifAviso"));
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
                if (Valid(out err))
                {
                    var entity = new DTO.AvisosDTO();
                    FormsHelper.FillEntity(tblControls, entity);

                    if (!trDuracion.Visible)
                        (entity as DTO.AvisosDTO).Duracion = null;

                    Business.Avisos.Create(entity, Atencion);

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

        protected void btnSave_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (Valid(out err))
                {
                    var entity = Business.Avisos.Read(Convert.ToInt32(pnlControls.Attributes["RecId"]));
                    FormsHelper.FillEntity(tblControls, entity);

                    if (!trDuracion.Visible)
                        (entity as DTO.AvisosDTO).Duracion = null;

                    Business.Avisos.Update(entity, Atencion);

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
                    Business.Avisos.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]));

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
                switch (e.Item.Name)
                {
                    case "btnAdd":
                        FormsHelper.ClearControls(tblControls, new DTO.AvisosDTO());
                        cbIdentifIdentAte.SelectedIndex = -1;
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);

                        Atencion = new List<DTO.AvisosIdAtenDTO>();
                        RefreshAbmGrid(gvABM);

                        pnlControls.Visible = true;
                        pnlControls.HeaderText = "Agregar Registro";
                        trDias.Visible = true;
                        break;

                    case "btnEdit":
                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ClearControls(tblControls, new DTO.AvisosDTO());
                            cbIdentifIdentAte.SelectedIndex = -1;
                            var entity = CRUDHelper.Read(FormsHelper.GetSelectedId(gv).Value, DAO);
                            FormsHelper.FillControls(entity, tblControls);
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);

                            Atencion = Business.Avisos.ReadAllAtencion((string)entity.IdentifAviso);

                            gvABM.Attributes.Add("IdentifAviso", entity.IdentifAviso);
                            trDias.Visible = true;
                            RefreshAbmGrid(gvABM);

                            pnlControls.Attributes.Add("RecId", entity.RecId.ToString());
                            pnlControls.Visible = true;
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
                            pnlControls.Visible = true;
                            pnlControls.HeaderText = "Eliminar Registros";
                        }
                        else
                        {
                            pnlControls.Visible = false;
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

        protected void btnAddAtencion_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (ValidAtencion(out err))
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.AvisosIdAtenDTO> aux   = Atencion;
                    var arteProducto                = new DTO.AvisosIdAtenDTO();
                    arteProducto.RecId              = aux.Count;
                    arteProducto.IdentifIdentAte    = Convert.ToString(cbIdentifIdentAte.SelectedItem.Value); //.SelectedItem);

                    aux.Add(arteProducto);

                    cbIdentifIdentAte.SelectedIndex = -1;
                    Atencion                        = aux;

                    RefreshAbmGrid(gvABM);
                }
                else
                {
                    throw new Exception(err);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorProducto, ex);
            }
        }

        protected void btnDeleteAtencion_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormsHelper.GetSelectedId(gvABM) != null)
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.AvisosIdAtenDTO> aux = new List<DTO.AvisosIdAtenDTO>();

                    //Creo una nueva coleccion con todos los productos menos el eliminado, y la guardo en el Viewstate.
                    foreach (var producto in Atencion)
                        if (Convert.ToInt32(FormsHelper.GetSelectedId(gvABM)) != producto.RecId)
                            aux.Add(producto);

                    Atencion = aux;
                    RefreshAbmGrid(gvABM);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorProducto, ex);
            }
        }

        private bool Valid(out string err)
        {
            err = string.Empty;

            if (string.IsNullOrEmpty(txIdentifAviso.Text.Trim()))
            {
                err = "Debe ingresar 'Aviso'.";
                return false;
            }

            if (ucIdentifEspacio.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Espacio de Contenido'.";
                return false;
            }

            if (ucIdentifFormAviso.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Formato del Aviso'.";
                return false;
            }

            if (ucIdentifPieza.SelectedValue == null)
            {
                err = "Debe seleccionar una 'Pieza de Arte'.";
                return false;
            }

            if (deVigDesde.Value == null || deVigHasta.Value == null)
            {
                err = "Debe seleccionar 'Vigencia Desde' y 'Vigencia Hasta'.";
                return false;
            }

            if (deVigDesde.Value != null && deVigHasta.Value != null
                && deVigHasta.Date.CompareTo(deVigDesde.Date) < 1)
            {
                err = "La 'Vigencia Hasta' debe ser mayor a la 'Vigencia Desde'.";
                return false;
            }

            if (Atencion == null || Atencion.Count == 0)
            {
                err = "Debe seleccionar algún identificador de atención.";
                return false;
            }
            return true;
        }

        private bool ValidAtencion(out string err)
        {
            err = string.Empty;
            if (cbIdentifIdentAte.SelectedItem == null /*ucIdentifIdenAten.SelectedValue == null*/)
            {
                err = "Debe seleccionar un 'Identificador de Atención'.";
                return false;
            }

            foreach (var att in Atencion)
            {
                if (att.IdentifIdentAte == Convert.ToString(cbIdentifIdentAte.SelectedItem.Value))
                {
                    err = "No puede ingresar dos veces el mismo 'Identificador de Atención'.";
                    return false;
                }
                
            }
            return true;
        }

        protected void cbIdentifIdentAte_ItemsRequestedByFilterCondition(object source, DevExpress.Web.ASPxEditors.ListEditItemsRequestedByFilterConditionEventArgs e)
        {
            ASPxComboBox comboBox = (ASPxComboBox)source;
            var daoIdentAtencion  = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion);
            comboBox.DataSource   = daoIdentAtencion.FindPaginado(e.Filter, (e.BeginIndex + 1), (e.EndIndex + 1));
            comboBox.DataBind();
        }

        protected void cbIdentifIdentAte_ItemRequestedByValue(object source, DevExpress.Web.ASPxEditors.ListEditItemRequestedByValueEventArgs e)
        {
            if (e.Value == null || Convert.ToString(e.Value) == string.Empty)
                return;

            ASPxComboBox comboBox = (ASPxComboBox)source;
            var daoIdentAtencion  = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion);
            comboBox.DataSource   = daoIdentAtencion.FindValue(Convert.ToString(e.Value));
            comboBox.DataBind();
        }

        protected void btnAdd_Click1(object sender, EventArgs e)
        {

        }

        protected void btnSave_Click1(object sender, EventArgs e)
        {

        }

        protected void ASPxMenu1_ItemClick1(object source, MenuItemEventArgs e)
        {

        }
    }
}