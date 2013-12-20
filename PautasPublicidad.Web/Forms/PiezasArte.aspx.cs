using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxGridView;
using PautasPublicidad.Business;
using DevExpress.Web.ASPxMenu;

namespace PautasPublicidad.Web.Forms
{
    public partial class PiezasArte : System.Web.UI.Page
    {
        dynamic DAO;

        public List<DTO.PiezasArteSKUDTO> Productos
        {
            get
            {
                if (ViewState["Productos"] != null)
                    return (List<DTO.PiezasArteSKUDTO>)ViewState["Productos"];
                else
                    return new List<DTO.PiezasArteSKUDTO>();
            }
            set
            {
                ViewState.Add("Productos", value);
            }
        }

        private const string PRIMARIO = "PRIMARIO";

        protected void Page_Load(object sender, EventArgs e)
        {
            DAO = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte);

            if (!Page.IsPostBack && !Page.IsCallback)
            {
                trMsg.Visible = false;
                trAbm.Visible = false;

                gvABM.SettingsBehavior.AllowSelectByRowClick = true;
                gvABM.SettingsBehavior.AllowSelectSingleRowOnly = true;
                FormsHelper.BuildColumnsByEntity(BusinessMapper.eEntities.PiezasArte, gv);
                FormsHelper.InicializarPropsGrilla(gv);
            }

            ucIdentifAnun.Inicializar(BusinessMapper.eEntities.AnunInternos);
            ucIdentifTipoPieza.Inicializar(BusinessMapper.eEntities.TipoPieza);
            ucIdentifSKU.Inicializar(BusinessMapper.eEntities.SKU);

            ucIdentifTipoPieza.ComboBox.AutoPostBack = true;
            ucIdentifTipoPieza.ComboBox.SelectedIndexChanged += new EventHandler(TipoPieza_SelectedIndexChanged);

            gvABM.KeyFieldName = "RecId";
            ASPxMenu1.ItemClick += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);

            rbTipoProd.AutoPostBack = true;
            rbTipoProd.SelectedIndexChanged += new EventHandler(rbTipoProd_SelectedIndexChanged);

            RefreshGrid(gv);
            RefreshAbmGrid(gvABM);
            lblError.Text = string.Empty;
            lblErrorProducto.Text = string.Empty;
        }

        void rbTipoProd_SelectedIndexChanged(object sender, EventArgs e)
        {
            trCoeficiente.Visible = (rbTipoProd.SelectedItem.Value.ToString() == PRIMARIO);
        }

        void TipoPieza_SelectedIndexChanged(object sender, EventArgs e)
        {
            DTO.TipoPiezaDTO tipoPieza = CRUDHelper.Read(string.Format("IdentifTipoPieza = '{0}'", ucIdentifTipoPieza.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.TipoPieza));

            trDuracion.Visible = (tipoPieza!=null && tipoPieza.Duracion);
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = CRUDHelper.ReadAll("", DAO);
            gv.DataBind();
        }

        private void RefreshAbmGrid(ASPxGridView gvABM)
        {
            gvABM.DataSource = Productos;
            gvABM.DataBind();
        }

        protected void detailGrid_DataSelect(object sender, EventArgs e)
        {
            ASPxGridView gvDetail = (ASPxGridView)sender;
            gvDetail.DataSource = Business.PiezasArte.ReadAllProductos((string)gvDetail.GetMasterRowFieldValues("IdentifPieza")); 
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
                    var entity = new DTO.PiezasArteDTO();
                    FormsHelper.FillEntity(tblControls, entity);

                    if (!trDuracion.Visible)
                        (entity as DTO.PiezasArteDTO).Duracion = null;

                    Business.PiezasArte.Create(entity, Productos);

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
                    var entity = CRUDHelper.Read(Convert.ToInt32(pnlControls.Attributes["RecId"]), DAO);
                    FormsHelper.FillEntity(tblControls, entity);

                    if (!trDuracion.Visible)
                        (entity as DTO.PiezasArteDTO).Duracion = null;

                    Business.PiezasArte.Update(entity, Productos);

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
                    Business.PiezasArte.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]));

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
                        FormsHelper.ClearControls(tblControls, new DTO.PiezasArteDTO());
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);

                        Productos = new List<DTO.PiezasArteSKUDTO>();
                        RefreshAbmGrid(gvABM);

                        pnlControls.Visible = true;
                        pnlControls.HeaderText = "Agregar Registro";
                        trDias.Visible = true;

                        
                        break;

                    case "btnEdit":
                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ClearControls(tblControls, new DTO.PiezasArteDTO());
                            var entity = CRUDHelper.Read(FormsHelper.GetSelectedId(gv).Value, DAO);
                            FormsHelper.FillControls(entity, tblControls);
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);

                            Productos = Business.PiezasArte.ReadAllProductos((string)entity.IdentifPieza);

                            gvABM.Attributes.Add("IdentifPieza", entity.IdentifPieza);
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

        protected void btnAddProducto_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (ValidProducto(out err))
                {
                    List<DTO.PiezasArteSKUDTO> aux = Productos;

                    var arteProducto = new DTO.PiezasArteSKUDTO();

                    arteProducto.RecId = aux.Count;
                    arteProducto.IdentifSKU = Convert.ToString(ucIdentifSKU.SelectedValue);
                    arteProducto.TipoProd = rbTipoProd.SelectedItem.Value.ToString();
                    if (arteProducto.TipoProd != PRIMARIO)
                        arteProducto.Coeficiente = null;
                    else
                        arteProducto.Coeficiente = Convert.ToDecimal(spCoeficiente.Value);

                    aux.Add(arteProducto);

                    ucIdentifSKU.ComboBox.SelectedIndex = -1;
                    spCoeficiente.Value = null;
                    Productos = aux;
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

        protected void btnDeleteProducto_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormsHelper.GetSelectedId(gvABM) != null)
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.PiezasArteSKUDTO> aux = new List<DTO.PiezasArteSKUDTO>();
                    
                    //Creo una nueva coleccion con todos los productos menos el eliminado, y la guardo en el Viewstate.
                    foreach (var producto in Productos)
                        if (Convert.ToInt32(FormsHelper.GetSelectedId(gvABM)) != producto.RecId)
                            aux.Add(producto);
                    
                    Productos = aux;
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
            
            if (string.IsNullOrEmpty(txIdentifPieza.Text.Trim()))
            {
                err = "Debe ingresar 'Pieza de Arte'.";
                return false;
            }

            if (ucIdentifAnun.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Anunciante Interno'.";
                return false;
            }

            if (ucIdentifTipoPieza.SelectedValue == null)
            {
                err = "Debe seleccionar un 'Tipo de Pieza'.";
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

            if (Productos == null || Productos.Count == 0)
            {
                err = "Debe seleccionar algún producto.";
                return false;
            }

            decimal t = 0;

            foreach (var p in Productos)
                if (p.Coeficiente.HasValue)
                    t += p.Coeficiente.Value;

            if (Math.Round(t, 1) != (decimal)1)
            {
                err = "La sumatoria de coeficientes (productos 'primarios') debe dar '1'.";
                return false;
            }
            return true;
        }
                
        private bool ValidProducto(out string err)
        {
            err = string.Empty;

            if (ucIdentifSKU.SelectedValue == null)
            {
                err = "Debe seleccionar un producto.";
                return false;
            }

            if (!((DTO.SKUDTO)CRUDHelper.Read(string.Format("IdentifSKU = '{0}'", ucIdentifSKU.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.SKU))).Activo)
            {
                err = "El producto no se encuentra Activo.";
                return false;
            }

            if (rbTipoProd.SelectedItem == null)
            {
                err = "Debe seleccionar si el tipo de producto es Primario o Secundario.";
                return false;
            }
            if (spCoeficiente.Value != null && (spCoeficiente.Number > 1 || spCoeficiente.Number < 0))
            {
                err = "El 'Coeficiente' debe estar entre los valores 0 y 1.";
                return false;
            }

            return true;
        }
    }
}