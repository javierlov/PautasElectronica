using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxGridView;
using PautasPublicidad.Business;
using DevExpress.Web.ASPxMenu;
using PautasPublicidad.DTO;

namespace PautasPublicidad.Web.Forms
{
    public partial class IdentificadoresAtencion : System.Web.UI.Page
    {
        protected void Page_PreInit(object sender, EventArgs e) 
        {
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
            {
                trTelefono.Visible = false;
                lblError.Text      = string.Empty;
                trMsg.Visible      = false;

                FormsHelper.InicializarPropsGrilla(gv);
            }

            ASPxMenu1.ItemClick += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);
            
            RefreshGrid(gv);
            lblError.Text = string.Empty;
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = CRUDHelper.ReadAll("", BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion));
            gv.DataBind();
        }

        protected void ASPxMenu1_ItemClick(object source, MenuItemEventArgs e)
        {
            switch (e.Item.Name)
            {
                case "btnAdd":
                    FormsHelper.ClearControls(tblControls, new DTO.IdentAtencionDTO());
                    FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Add);    
                    pnlControls.Visible = true;
                    pnlControls.HeaderText = "Agregar Registro";

                    break;

                case "btnEdit":
                    if (FormsHelper.GetSelectedId(gv) != null)
                    {
                        FormsHelper.ClearControls(tblControls, new DTO.IdentAtencionDTO());
                        var entity = CRUDHelper.Read(FormsHelper.GetSelectedId(gv).Value, BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion));
                        FormsHelper.FillControls(entity, tblControls);
                        FormsHelper.ShowOrHideButtons(tblControls, FormsHelper.eAccionABM.Edit);
                        pnlControls.Attributes.Add("RecId", entity.RecId.ToString());
                        pnlControls.Visible = true;
                        pnlControls.HeaderText = "Modificar Registro";

                        TipIdentifChanged();
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

        protected void rbTipIdentif_SelectedIndexChanged(object sender, EventArgs e)
        {
            TipIdentifChanged();
        }

        private void TipIdentifChanged()
        {
            if (Convert.ToString(rbTipIdentif.SelectedItem.Value) == "TELEFONO")
            {
                trTelefono.Visible = true;
            }
            else
            {
                trTelefono.Visible      = false;
                spDNIS.Value            = "";
                txTelefono.Text         = string.Empty;
                rbTipoCDN.SelectedIndex = -1;
            }
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            pnlControls.Visible = false;
        }

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                var entity = new DTO.IdentAtencionDTO();
                FormsHelper.FillEntity(tblControls, entity);
                CRUDHelper.Create(entity,
                    BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion));

                pnlControls.Visible = false;
                RefreshGrid(gv);
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion);
                IdentAtencionDTO entity = (DTO.IdentAtencionDTO)CRUDHelper.Read(Convert.ToInt32(pnlControls.Attributes["RecId"]), dao);
                
                FormsHelper.FillEntity(tblControls, entity);

                if (!entity.Estado && dao.ExisteAvisoEnVigencia(entity.IdentifIdentAte))
                {
                    throw new Exception("No se puede deshabilitar, ya que se encuentra en uso por un aviso vigente.");
                }
                
                CRUDHelper.Update(entity,
                    BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion));

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
                    CRUDHelper.Delete(Convert.ToInt32(pnlControls.Attributes["RecId"]), 
                        BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion));

                    pnlControls.Visible = false;
                    RefreshGrid(gv);
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }
    }
}