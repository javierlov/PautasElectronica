using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;
using System.Xml;

namespace PautasPublicidad.Web
{
    public partial class PautasPublicidad : System.Web.UI.MasterPage
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            XmlDocument xDoc = new XmlDocument();
            string settingsPath = Server.MapPath("~/Settings.xml");

            if (!System.IO.File.Exists(settingsPath))
                throw new Exception("No existe el archivo 'Settings.xml' en el raiz del sitio web.");

            xDoc.Load(settingsPath);

            if (xDoc.SelectSingleNode("/Settings/ConnectionString") != null)
                Conexion.SetConnectionString(xDoc.SelectSingleNode("/Settings/ConnectionString").InnerText);
            else
                Conexion.SetConnectionString(@"Data Source=.;Initial Catalog=Publicidad;Integrated Security=True");
        }
       
    }
}