using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.DTO;
using System.Xml;

namespace PautasPublicidad.Web
{
    public partial class Login : System.Web.UI.Page
    {
        public UsuariosDTO Usuario;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                Session.Clear();
                LoadSettings();
                TestConnection();
                LoadEmpresas();
            }
        }

        private void LoadEmpresas()
        {
            ddlEmpresa1.Items.Clear();

            var empresas = Business.Conexion.GetEmpresas();

            foreach (var item in empresas)
            {
                ddlEmpresa1.Items.Add(new ListItem(item.Name, item.DatareaId.ToString()));
            }
        }

        private void TestConnection()
        {
            string error;

            Business.Conexion.TestConnection(out error);

            ASPxButton1.Enabled = (error == "");
            lblMsg.Text         = error;
        }

        private void LoadSettings()
        {
            XmlDocument xDoc    = new XmlDocument();
            string settingsPath = Server.MapPath("~/Settings.xml");

            if (!System.IO.File.Exists(settingsPath))
                throw new Exception("No existe el archivo 'Settings.xml' en el raiz del sitio web.");

            xDoc.Load(settingsPath);

            if (xDoc.SelectSingleNode("/Settings/ConnectionString") != null)
                Business.Conexion.SetConnectionString(xDoc.SelectSingleNode("/Settings/ConnectionString").InnerText);
            else
                Business.Conexion.SetConnectionString(@"Data Source=.;Initial Catalog=Publicidad;Integrated Security=True");

            if (xDoc.SelectSingleNode("/Settings/MapperInfoXmlPath") != null)
                Business.BusinessMapper.MapperInfoXmlPath = xDoc.SelectSingleNode("/Settings/MapperInfoXmlPath").InnerText;
            else
                Business.BusinessMapper.MapperInfoXmlPath = Server.MapPath("~/App_Data/MapperInfo.xml");

            if(xDoc.SelectSingleNode("/Settings/AbmConfigXmlPath") != null)
                Business.BusinessMapper.AbmConfigXmlPath = xDoc.SelectSingleNode("/Settings/AbmConfigXmlPath").InnerText;
            else
                Business.BusinessMapper.AbmConfigXmlPath = Server.MapPath("~/App_Data/AbmConfig.xml");
        }

        protected void btnLogin_Click1(object sender, EventArgs e)
        {

        }

        protected void ASPxButton1_Click(object sender, EventArgs e)
        {
            UsuariosDTO user;
            int DatareaId = Convert.ToInt32(ddlEmpresa1.SelectedValue);

            if (Business.Conexion.Login(txUserName.Text.Trim(), txPassword.Text.Trim(), DatareaId, out user))
            {
                Business.Conexion.SetDatareaId(DatareaId);

                Session.Add("App.Empresa", Business.Conexion.GetEmpresas().Find(x => x.DatareaId == DatareaId));
                Session.Add("App.User", user);

                Response.Redirect("~/Forms/Ordenado.aspx");
            }
            else
            {
                lblMsg.Text = "Nombre de Usuario o Contraseña Incorrecta.";
            }
        }
    }
}