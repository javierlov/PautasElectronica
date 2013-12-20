using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Web.ASPxEditors;
using PautasPublicidad.DTO;
using PautasPublicidad.Business;

namespace PautasPublicidad.Web.Controls
{
    public partial class ucComboBox : System.Web.UI.UserControl
    {
        public delegate void ComboBoxEventHandler(object sender, EventArgs e);
        public event ComboBoxEventHandler SelectedIndexChanged;
        public delegate List<TablaBase> ReadAll(string filtro);

        #region Propiedades

        public ASPxComboBox ComboBox { get { return this.cmbEntidad; } }

        public string EntityName
        {
            get
            {
                return Convert.ToString(ViewState["EntityName"]);
            }
            set
            {
                ViewState.Add("EntityName", value);
            }
        }

        public MapInfo MapInfo
        {
            get
            {
                return BusinessMapper.GetMapInfo(EntityName);
            }
        }

        public object SelectedValue 
        { 
            get 
            {
                if (cmbEntidad.SelectedItem != null)
                    return cmbEntidad.SelectedItem.Value;
                else
                    return null;
            } 
            set 
            {
                cmbEntidad.SelectedItem = cmbEntidad.Items.FindByValue(value);
            } 
        }

        public string SelectedText 
        { 
            get 
            {
                if (cmbEntidad.SelectedItem != null)
                    return cmbEntidad.SelectedItem.Text;
                else
                    return null;
            } 
            set 
            { 
                cmbEntidad.SelectedItem.Text = value; 
            } 
        }

        public string WhereFilter
        { get; set; }

        #endregion

        protected virtual void OnSelectedIndexChanged(EventArgs e)
        {
            if (SelectedIndexChanged != null)
            {
                SelectedIndexChanged(this, e);
            }
        }

        [Obsolete("Utilizar: Inicializar(string entityName)")]
        public void Inicializar(ReadAll metodoDataSource, string textField, string valueField)
        {
            this.cmbEntidad.TextField  = textField;
            this.cmbEntidad.ValueField = valueField;
            this.cmbEntidad.DataSource = metodoDataSource("");

            this.cmbEntidad.DataBind();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (EntityName != null && EntityName != string.Empty)
            {
                btnAdd.OnClientClick = FormsHelper.GetWindowOpenScript(MapInfo.ABMUrl.Replace("~/Forms/",""), MapInfo.EntityName);
            }
        }

        protected void btnRefresh_Click(object sender, ImageClickEventArgs e)
        {
            Inicializar(MapInfo.DAOHandler.ReadAll(WhereFilter), MapInfo.EntityTextField, MapInfo.EntityValueField);
        }

        internal void Inicializar(string entityName)
        {
            EntityName = entityName;
            Inicializar(MapInfo.DAOHandler.ReadAll(""), MapInfo.EntityTextField, MapInfo.EntityValueField);
        }

        internal void Inicializar(BusinessMapper.eEntities entityName)
        {
            EntityName = entityName.ToString();
            Inicializar(MapInfo.DAOHandler.ReadAll(""), MapInfo.EntityTextField, MapInfo.EntityValueField);
        }

        internal void Inicializar(BusinessMapper.eEntities entityName, string where)
        {
            EntityName = entityName.ToString();
            Inicializar(MapInfo.DAOHandler.ReadAll(where), MapInfo.EntityTextField, MapInfo.EntityValueField);
        }

        internal void Inicializar(BusinessMapper.eEntities entityName, bool autoPostBack)
        {
            EntityName = entityName.ToString();
            Inicializar(MapInfo.DAOHandler.ReadAll(""), MapInfo.EntityTextField, MapInfo.EntityValueField);
            cmbEntidad.AutoPostBack = autoPostBack;
        }

        private void Inicializar(dynamic dataSource, string textField, string valueField)
        {
            if (this.cmbEntidad == null)
                this.cmbEntidad = new ASPxComboBox();

            this.cmbEntidad.Columns.Clear();
            this.cmbEntidad.Columns.Add(MapInfo.EntityValueField, "Código");
            this.cmbEntidad.Columns.Add(MapInfo.EntityTextField, "Elemento");

            this.cmbEntidad.TextField = textField;
            this.cmbEntidad.ValueField = valueField;
            this.cmbEntidad.DataSource = dataSource;
            this.cmbEntidad.DataBind();

            WhereFilter = string.Empty;
        }

        public Unit Width { get; set; }

        protected void cmbEntidad_SelectedIndexChanged(object sender, EventArgs e)
        {
            OnSelectedIndexChanged(e);
        }
    }
}