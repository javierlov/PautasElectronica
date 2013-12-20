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
    public partial class Monedas : System.Web.UI.Page
    {
        //dynamic DAO;

        public List<DTO.TipoCambioDTO> TipoCambio
        {
            get
            {
                if (ViewState["TipoCambio"] != null)
                    return (List<DTO.TipoCambioDTO>)ViewState["TipoCambio"];
                else
                    return new List<DTO.TipoCambioDTO>();
            }
            set
            {
                ViewState.Add("TipoCambio", value);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            //DAO = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Monedas);

            lblError.Text = string.Empty;
            lblErrorProducto.Text = string.Empty;

            if (!Page.IsPostBack && !Page.IsCallback)
            {
                trMsg.Visible = false;
                trAbm.Visible = false;

                gvABM.SettingsBehavior.AllowSelectByRowClick    = true;
                gvABM.SettingsBehavior.AllowSelectSingleRowOnly = true;
                FormsHelper.BuildColumnsByEntity(BusinessMapper.eEntities.Monedas, gv);
                FormsHelper.InicializarPropsGrilla(gv);
            }

            gvABM.KeyFieldName = "RecId";

            RefreshGrid(gv);
            RefreshAbmGrid(gvABM);
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = Business.Monedas.ReadAll("");
            gv.DataBind();
        }

        private void RefreshAbmGrid(ASPxGridView gvABM)
        {
            gvABM.DataSource = TipoCambio;
            gvABM.DataBind();
        }

        protected void detailGrid_DataSelect(object sender, EventArgs e)
        {
            ASPxGridView gvDetail = (ASPxGridView)sender;
            gvDetail.DataSource   = Business.Monedas.ReadAllTiposCambio((string)gvDetail.GetMasterRowFieldValues("IdentifMon"));
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
                    var entity = new DTO.MonedasDTO();
                    FormsHelper.FillEntity(tblControls, entity);

                    Business.Monedas.Create(entity, TipoCambio);

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
            return true;
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            string err = string.Empty;

            try
            {
                if (Valid(out err))
                {
                    var entity = Business.Monedas.Read(Convert.ToInt32(pnlControls.Attributes["RecId"]));
                    FormsHelper.FillEntity(tblControls, entity);

                    Business.Monedas.Update(entity, TipoCambio);

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
                    Business.Monedas.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]));

                    pnlControls.Visible = false;
                    RefreshGrid(gv);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        protected void ASPxMenu1_ItemClick(object source, DevExpress.Web.ASPxMenu.MenuItemEventArgs e)
        {
            try
            {
                switch (e.Item.Name)
                {
                    case "btnAdd":
                        FormsHelper.ClearControls(tblControls, new DTO.MonedasDTO());
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);

                        TipoCambio = new List<DTO.TipoCambioDTO>();
                        RefreshAbmGrid(gvABM);

                        pnlControls.Visible = true;
                        pnlControls.HeaderText = "Agregar Registro";
                        trDias.Visible = true;
                        break;

                    case "btnEdit":
                        if (FormsHelper.GetSelectedId(gv) != null)
                        {
                            FormsHelper.ClearControls(tblControls, new DTO.MonedasDTO());

                            var entity = Business.Monedas.Read(FormsHelper.GetSelectedId(gv).Value);

                            FormsHelper.FillControls(entity, tblControls);
                            FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);

                            TipoCambio = Business.Monedas.ReadAllTiposCambio((string)entity.IdentifMon);

                            spDuracion0.Value = null;
                            deVigDesde0.Value = DateTime.Now;

                            gvABM.Attributes.Add("IdentifMon", entity.IdentifMon);

                            trDias.Visible = true;

                            RefreshAbmGrid(gvABM);

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
                if (ValidCambio(out err))
                {
                    List<DTO.TipoCambioDTO> aux = TipoCambio;

                    var tipoCambio = new DTO.TipoCambioDTO();

                    for(int i = 0;i<=TipoCambio.Count-1;i++)
                    {
                        if (TipoCambio[i].FechaInicio.ToShortDateString() == deVigDesde0.Date.ToShortDateString())
                        {
                            return;
                        }
                    }

                    tipoCambio.RecId       = aux.Count;
                    tipoCambio.Valor       = spDuracion0.Number;
                    tipoCambio.FechaInicio = deVigDesde0.Date;

                    for (int i = 0; i <= aux.Count - 1; i++)
                    { 
                        if(aux[i].FechaInicio == tipoCambio.FechaInicio)
                        {
                            err = "Fecha ya asignada.";

                            return;
                        }

                    }
                    
                    aux.Add(tipoCambio);

                    spDuracion0.Value = null;
                    deVigDesde0.Date  = DateTime.Today;

                    TipoCambio = aux;
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
                    List<DTO.TipoCambioDTO> aux = new List<DTO.TipoCambioDTO>();

                    //Creo una nueva coleccion con todos los productos menos el eliminado, y la guardo en el Viewstate.
                    foreach (var tipoCambio in TipoCambio)
                        if (Convert.ToInt32(FormsHelper.GetSelectedId(gvABM)) != tipoCambio.RecId)
                            aux.Add(tipoCambio);

                    TipoCambio = aux;
                    RefreshAbmGrid(gvABM);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblErrorProducto, ex);
            }
        }

        private bool ValidCambio(out string err)
        {
            err = string.Empty;
            return true;
        }
    }
}