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
    public partial class Frecuencia : System.Web.UI.Page
    {
        public List<DTO.FrecuenciaDetDTO> Detalles
        {
            get
            {
                if (ViewState["Detalles"] != null)
                    return (List<DTO.FrecuenciaDetDTO>)ViewState["Detalles"];
                else
                    return new List<DTO.FrecuenciaDetDTO>();
            }
            set
            {
                ViewState.Add("Detalles", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
            {
                trMsg.Visible = false;
                trAbm.Visible = false;

                gvABM.SettingsBehavior.AllowSelectByRowClick = true;
                gvABM.SettingsBehavior.AllowSelectSingleRowOnly = true;

                FormsHelper.InicializarPropsGrilla(gv);
                GridViewDataComboBoxColumn c = (GridViewDataComboBoxColumn)gvABM.Columns["DiaSemana"];
                FormsHelper.FillDias(c.PropertiesComboBox.Items);
            }

            gvABM.KeyFieldName = "RecId";
            ASPxMenu1.ItemClick += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);
            rbSemMes.AutoPostBack = true;
            rbSemMes.SelectedIndexChanged += new EventHandler(rbSemMes_SelectedIndexChanged);


            RefreshGrid(gv);
            RefreshAbmGrid(gvABM);

            lblError.Text = string.Empty;
            lblErrorDia.Text = string.Empty;
        }

        protected void rbSemMes_SelectedIndexChanged(object sender, EventArgs e)
        {
            rbSemMesChanged();
        }

        private void rbSemMesChanged()
        {
            int retVal = 0;

            retVal = rbSemMes.SelectedItem.Text != null && rbSemMes.SelectedItem.Text != "" ? 1 : 0;

            if (retVal == 1)
            {
                pnlDias.Visible      = true;
                cbDias.SelectedIndex = -1;
                cbDias.Items.Clear();

                if (Convert.ToString(rbSemMes.SelectedItem.Value).ToLower() == "semana")
                {
                    //Cargo dias de la semana.
                    FormsHelper.FillDias(cbDias.Items);

                    gvABM.Columns["DiaSemana"].Visible = true;
                    gvABM.Columns["Dia"].Visible       = false;
                }
                else
                {
                    //Cargo nros de dias del mes.
                    for (int i = 1; i < 32; i++)
                        cbDias.Items.Add(i.ToString(), i.ToString());

                    gvABM.Columns["DiaSemana"].Visible = false;
                    gvABM.Columns["Dia"].Visible       = true;
                }
            }
            else
            {
                pnlDias.Visible = true;
            }
        }

        void gvABM_RowDeleting(object sender, DevExpress.Web.Data.ASPxDataDeletingEventArgs e)
        {
            e.Cancel = true;

            CRUDHelper.Delete(Convert.ToInt32(e.Keys[0]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

            RefreshAbmGrid(gvABM);
        }

        void gvABM_RowInserting(object sender, DevExpress.Web.Data.ASPxDataInsertingEventArgs e)
        {

        }

        protected void detailGrid_DataSelect(object sender, EventArgs e)
        {
            ASPxGridView gvDetail = (ASPxGridView)sender;

            gvDetail.DataSource = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", gvDetail.GetMasterRowFieldValues("IdentifFrecuencia")), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

            gvDetail.Columns["IdentifFrecuencia"].Visible = false;

            string semMes = Convert.ToString(gvDetail.GetMasterRowFieldValues("SemMes"));

            gvDetail.Columns["DiaSemana"].Visible = (semMes.Trim().ToLower() == "semana");
            gvDetail.Columns["Dia"].Visible = (semMes.Trim().ToLower() != "semana");
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = CRUDHelper.ReadAll("", BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
            gv.DataBind();
        }

        private void RefreshAbmGrid(ASPxGridView gvABM)
        {
            gvABM.DataSource = Detalles;
            gvABM.DataBind();
        }

        protected void ASPxGridView1_BeforePerformDataSelect(object sender, EventArgs e)
        {
        }

        protected void ASPxGridView1_CustomUnboundColumnData(object sender, DevExpress.Web.ASPxGridView.ASPxGridViewColumnDataEventArgs e)
        {

        }

        protected void ASPxMenu1_ItemClick(object source, MenuItemEventArgs e)
        {
            try
            {
                switch (e.Item.Name)
                {
                    case "btnAdd":

                        FormsHelper.ClearControls(tblControls, new DTO.FrecuenciaDTO());
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);

                        pnlDias.Visible         = false;
                        Detalles                = new List<DTO.FrecuenciaDetDTO>();

                        RefreshAbmGrid(gvABM);

                        pnlControls.Visible     = true;
                        pnlControls.HeaderText  = "Agregar Registro";
                        trDias.Visible          = true;
                        rbSemMes.Enabled        = true;

                        break;

                    case "btnEdit":

                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ClearControls(tblControls, new DTO.FrecuenciaDTO());
                            var entity = CRUDHelper.Read(FormsHelper.GetSelectedId(gv).Value, BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                            FormsHelper.FillControls(entity, tblControls);
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);

                            Detalles = CRUDHelper.ReadAll(string.Format("IdentifFrecuencia = '{0}'", entity.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FrecuenciaDet));

                            rbSemMes.Enabled = (Detalles.Count == 0);

                            rbSemMesChanged();

                            gvABM.Attributes.Add("IdentifFrecuencia", entity.IdentifFrecuencia);
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
                        rbSemMes.Enabled = (Detalles.Count == 0);
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
                            ASPxGridViewExporter1.GridViewID = "gv";
                            ASPxGridViewExporter1.WriteXlsToResponse();
                        break;

                    case "btnExportPdf":
                        if (ASPxGridViewExporter1 != null)
                            ASPxGridViewExporter1.GridViewID = "gv";
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

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            pnlControls.Visible = false;
            gv.Enabled = true;
        }

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (Valid(out err))
                {
                    var entity = new DTO.FrecuenciaDTO();
                    FormsHelper.FillEntity(tblControls, entity);

                    Business.Frecuencias.Create(entity, Detalles);

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

        private bool Valid(out string err)
        {
            err = string.Empty;

            if (txIdentifFrecuencia.Text.Trim() == string.Empty)
            {
                err = "Debe ingresar un nombre para la 'Frecuencia'.";
                return false;
            }

            if (rbSemMes.SelectedItem == null)
            {
                err = "Debe seleccionar 'Semana' o 'Més'.";
                return false;
            }

            if (Detalles == null || Detalles.Count == 0)
            {
                err = "Debe ingresar al menos un 'Día'.";
                return false;
            }
            return true;
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var entity = CRUDHelper.Read(Convert.ToInt32(pnlControls.Attributes["RecId"]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));

                FormsHelper.FillEntity(tblControls, entity);

                Business.Frecuencias.Update(entity, Detalles);

                pnlControls.Visible = false;
                RefreshGrid(gv);
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
                    Business.Frecuencias.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]));

                    pnlControls.Visible = false;
                    gv.Enabled = true;
                    RefreshGrid(gv);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        

        protected void btnAddDia_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (cbDias.Value == null) { return; }

                if (ValidDias(out err))
                {
                    DTO.FrecuenciaDetDTO entity = new DTO.FrecuenciaDetDTO();

                    entity.RecId = Detalles.Count;

                    if (Convert.ToString(rbSemMes.SelectedItem.Value).ToLower().Trim() == "semana")
                    {
                        entity.DiaSemana = Convert.ToString(cbDias.SelectedItem.Value);

                        rbSemMes.Enabled = false;

                        for (int i = 0; i <= Detalles.Count-1; i++)
                        {
                            if (Detalles[i].DiaSemana.ToUpper().Trim() == entity.DiaSemana.ToUpper().Trim())
                            {
                                entity.DiaSemana = null;

                                break;
                            }

                        }
                        entity.Dia = null;
                    }
                    else
                    {
                        entity.Dia = Convert.ToInt32(cbDias.SelectedItem.Value);

                        rbSemMes.Enabled = false;

                        for (int i = 0; i <= Detalles.Count - 1; i++)
                        {
                            if (Detalles[i].Dia == entity.Dia)
                            {
                                entity.Dia = null;

                                break;
                            }

                        }

                        entity.DiaSemana = null;

                    }

                    //Solo guardo el registro si alguno de los campos tiene valor!

                    if (entity.DiaSemana != null || entity.Dia != null)
                    {
                        List<DTO.FrecuenciaDetDTO> aux = Detalles;
                        aux.Add(entity);
                        Detalles = aux;

                        RefreshAbmGrid(gvABM);

                        cbDias.SelectedIndex = -1;
                    }
                }
                else
                {
                    throw new Exception(err);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorDia, ex);
            }
        }

        private bool ValidDias(out string err)
        {
            bool retval = true;
            err = "";
            return retval;
        }

        protected void btnDeleteDia_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormsHelper.GetSelectedId(gvABM) != null)
                {
                    //Si no hago esto con un aux, no funciona, porque 'Productos' se actualiza en el Viewstate.
                    List<DTO.FrecuenciaDetDTO> aux = new List<DTO.FrecuenciaDetDTO>();

                    //Creo una nueva coleccion con todos los productos menos el eliminado, y la guardo en el Viewstate.
                    foreach (var detalle in Detalles)
                        if (Convert.ToInt32(FormsHelper.GetSelectedId(gvABM)) != detalle.RecId)
                            aux.Add(detalle);

                    Detalles = aux;
                    RefreshAbmGrid(gvABM);

                    rbSemMes.Enabled = (Detalles.Count == 0);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorDia, ex);
            }
        }

        protected void ASPxMenu1_ItemClick1(object source, MenuItemEventArgs e)
        {

        }

        protected void rbSemMes_SelectedIndexChanged1(object sender, EventArgs e)
        {

        }
        
    }
}