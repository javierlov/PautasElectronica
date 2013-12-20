using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.DTO;
using System.Xml;
using DevExpress.Web.ASPxNavBar;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web
{
    public partial class Accendo : System.Web.UI.MasterPage
    {
        public UsuariosDTO Usuario 
        {
            get 
            {
                if (Session["App.User"] == null)
                {
                    Response.Redirect("~/Login.aspx", true);
                    return null;
                }
                else
                {
                    return (UsuariosDTO)Session["App.User"];
                }
            }
        }

        public EmpresaDTO Empresa
        {
            get
            {
                if (Session["App.Empresa"] == null)
                {
                    Response.Redirect("~/Login.aspx", true);
                    return null;
                }
                else
                {
                    return (EmpresaDTO)Session["App.Empresa"];
                }
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            lblUsuario.Text = Usuario.UserName;
            lblEmpresa.Text = Empresa.Name;

            var g = ASPxNavBar1.Groups.FindByName("Parametrizacion");
            
            BuildMenuByAbmConfig(g);
        }

        private void BuildMenuByAbmConfig(NavBarGroup grupo)
        {

            try {

                //List<ABMControl> controls = new List<ABMControl>();

                if (BusinessMapper.AbmConfigXmlPath == null || BusinessMapper.AbmConfigXmlPath == string.Empty)
                    throw new Exception("Path del archivo AbmConfig.xml sin definir.");

                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(BusinessMapper.AbmConfigXmlPath);

                grupo.Items.Clear();
                foreach (XmlNode nodo in xDoc.SelectNodes("Entities/Entity"))
                {
                    if (nodo.Attributes["AbmUrl"] != null)
                    {
                        grupo.Items.Add(
                            new NavBarItem(
                                nodo.Attributes["Title"].Value,
                                nodo.Attributes["EntityName"].Value, "",
                                nodo.Attributes["AbmUrl"].Value));
                    }
                    else
                    {
                        grupo.Items.Add(
                            new NavBarItem(
                                nodo.Attributes["Title"].Value,
                                nodo.Attributes["EntityName"].Value, "",
                                string.Format("~/Forms/AbmGenerico.aspx?EntityName={0}",
                                    nodo.Attributes["EntityName"].Value)));
                    }
                }
            
            }
            catch (Exception ex)
            {
                string a = ex.Message;
            }
        }

        protected void lnkCerrarSesion_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Login.aspx", true);
        }

        protected void ASPxNavBar1_ItemClick(object source, NavBarItemEventArgs e)
        {

        }

    }
}