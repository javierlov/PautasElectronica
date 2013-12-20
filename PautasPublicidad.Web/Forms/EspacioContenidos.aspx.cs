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
using System.Xml;

namespace PautasPublicidad.Web.Forms
{
    public partial class EspacioContenidos : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
            {
                lblError.Text        = string.Empty;
                trMsg.Visible        = false;
                trFrecuencia.Visible = false;
                trHoraFin.Visible    = false;
                trHoraInicio.Visible = false;
                trIntervalo.Visible  = false;

                FormsHelper.InicializarPropsGrilla(gv);
                FormsHelper.BuildColumnsByEntity(BusinessMapper.eEntities.EspacioCont, gv);
            }
            
            ucIdentifFrecuencia.Inicializar(BusinessMapper.eEntities.Frecuencia);
            ucIdentifIntervalo.Inicializar(BusinessMapper.eEntities.Intervalo);
            ucIdentifMedio.Inicializar(BusinessMapper.eEntities.MediosPub);
            ucIdentifTipoEsp.Inicializar(BusinessMapper.eEntities.TipoEspacio, true);
            ucIdentifTipoEsp.ComboBox.SelectedIndexChanged += new EventHandler(TipoEspacio_SelectedIndexChanged);

            ASPxMenu1.ItemClick += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);

            RefreshGrid(gv);

            lblError.Text = string.Empty;
        }

        void TipoEspacio_SelectedIndexChanged(object sender, EventArgs e)
        {
            TipoDeEspacioChanged();
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = CRUDHelper.ReadAll("", BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
            gv.DataBind();
        }

        protected void ASPxMenu1_ItemClick(object source, MenuItemEventArgs e)
        {
            try
            {
                switch (e.Item.Name)
                {
                    case "btnAdd":

                        FormsHelper.ClearControls(tblControls, new DTO.EspacioContDTO());
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);
                        pnlControls.Visible = true;
                        pnlControls.HeaderText = "Agregar Registro";
                        break;

                    case "btnEdit":

                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ClearControls(tblControls, new DTO.EspacioContDTO());
                            var entity = CRUDHelper.Read(FormsHelper.GetSelectedId(gv).Value, BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));
                            FormsHelper.FillControls(entity, tblControls);
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);
                            pnlControls.Attributes.Add("RecId", entity.RecId.ToString());
                            pnlControls.Visible = true;
                            pnlControls.HeaderText = "Modificar Registro";

                            TipoDeEspacioChanged();
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

        private void TipoDeEspacioChanged()
        {
            DTO.TipoEspacioDTO espacio = CRUDHelper.Read(string.Format("IdentifTipoEsp='{0}'", ucIdentifTipoEsp.SelectedValue), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.TipoEspacio));
            trFrecuencia.Visible       = espacio.Frecuencia;
            trHoraInicio.Visible       = espacio.Hora;
            trHoraFin.Visible          = espacio.Hora;
            trIntervalo.Visible        = espacio.Intervalo;

            UpdatePanel1.Update();
        }

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
                    var entity = new DTO.EspacioContDTO();
                    FormsHelper.FillEntity(tblControls, entity);

                    if (!trIntervalo.Visible)
                        entity.IdentifIntervalo = null;

                    if (!trFrecuencia.Visible)
                        entity.IdentifFrecuencia = null;

                    if (!trHoraFin.Visible || !trHoraInicio.Visible)
                    {
                        entity.HoraFin    = null;
                        entity.HoraInicio = null;
                    }

                    CRUDHelper.Create(entity, BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

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
                    var entity = CRUDHelper.Read(Convert.ToInt32(pnlControls.Attributes["RecId"]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

                    FormsHelper.FillEntity(tblControls, entity);

                    if (!trIntervalo.Visible)
                        entity.IdentifIntervalo = null;

                    if (!trFrecuencia.Visible)
                        entity.IdentifFrecuencia = null;

                    if (!trHoraFin.Visible || !trHoraInicio.Visible)
                    {
                        entity.HoraFin    = null;
                        entity.HoraInicio = null;
                    }

                    CRUDHelper.Update(entity, BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

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
                    CRUDHelper.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.EspacioCont));

                    pnlControls.Visible = false;
                    RefreshGrid(gv);
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

            if (teHoraFin.Value != null && teHoraInicio.Value != null && 
                !FormsHelper.HoraEsMayor(teHoraInicio.DateTime, teHoraFin.DateTime))
            {
                err = "La 'Hora de Fin' debe ser mayor a la 'Hora de Inicio'.";

                return false;
            }           
            return true;
        }
    }
}