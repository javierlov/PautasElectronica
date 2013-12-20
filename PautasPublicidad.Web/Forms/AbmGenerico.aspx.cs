using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Web.Controls;
using PautasPublicidad.Business;
using System.Xml;
using DevExpress.Web.ASPxGridView;
using DevExpress.Web.ASPxMenu;

namespace PautasPublicidad.Web
{
    public partial class AbmGenerico : System.Web.UI.Page
    {
        public BusinessMapper.eEntities Entity 
        { 
            get 
            { 
                return BusinessMapper.GetEntityByName((string)ViewState["EntityName"]); 
            } 
        }

        public string Usuario;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack && !Page.IsCallback)
            {
                DTO.UsuariosDTO ui = (DTO.UsuariosDTO)Session["App.User"]; 

                if (Request.QueryString["EntityName"] == "SetUp")
                {

                    if(ui.UserName != "admin") 
                    {
                        Response.Redirect("~/Forms/Ordenado.aspx");
                    }
                }
                if (Request.QueryString["EntityName"] != null)
                {
                    ViewState.Add("EntityName", Request.QueryString["EntityName"]);
                }

                FormsHelper.BuildColumnsByEntity(Entity, gv);
                FormsHelper.InicializarPropsGrilla(gv);

                ucABM1.Visible = false;
            }

            ASPxMenu1.ItemClick     += new DevExpress.Web.ASPxMenu.MenuItemEventHandler(ASPxMenu1_ItemClick);
            ucABM1.ActualizarGrilla += new ucABM.ABMEventHandler(ucABM1_ActualizarGrilla);

            ucABM1.Inicializar(Entity);
            
            RefreshGrid(gv);
        }

        void ucABM1_ActualizarGrilla(object sender, ABMEventArgs e)
        {
            RefreshGrid(gv);
        }

        private void RefreshGrid(ASPxGridView gv)
        {
            gv.DataSource = CRUDHelper.ReadAll("", BusinessMapper.GetDaoByEntity(Entity));
            gv.DataBind();
        }

        protected void ASPxMenu1_ItemClick(object source, MenuItemEventArgs e)
        {
            FormsHelper.ToolBarClick(ucABM1, e.Item.Name, gv, ASPxGridViewExporter1);
        }

        protected void ASPxMenu1_ItemClick1(object source, MenuItemEventArgs e)
        {

        }
    }
}