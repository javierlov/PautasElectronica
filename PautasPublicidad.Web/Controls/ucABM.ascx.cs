using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using DevExpress.Web.ASPxEditors;
using PautasPublicidad.DTO;
using PautasPublicidad.Business;
using System.Xml;

namespace PautasPublicidad.Web.Controls
{
    public partial class ucABM : System.Web.UI.UserControl
    {
        public int Width { set { pnlMain.Width = new Unit(value); } }
        
        public HtmlTable tablaABM { get { return tblABM; } }
        public HtmlTable tablaControles { get { return tblControls; } }
        public string HeaderText { set { pnlMain.HeaderText = value; } }

        public TablaBase ObjetoDTO { set; get; }
        Business.BusinessMapper.eEntities Entidad;
        private dynamic dao;

        public delegate dynamic ReadDelegate(int RecId);
        public delegate void DeleteDelegate(int RecId);
        public delegate void CreateDelegate(object entity);
        public delegate void UpdateDelegate(object entity);

        #region Eventos
        public delegate void ABMEventHandler(object sender, ABMEventArgs e);

        public event ABMEventHandler Cancelar;
        public event ABMEventHandler Agregar;
        public event ABMEventHandler Eliminar;
        public event ABMEventHandler Guardar;
        public event ABMEventHandler ActualizarGrilla;

        protected virtual void OnCancelar(ABMEventArgs e)
        {
            if (Cancelar != null)
            {
                Cancelar(this, e);
            }
        }

        protected virtual void OnAgregar(ABMEventArgs e)
        {
            if (Agregar != null) 
            {
                Agregar(this, e); 
            }
        }

        protected virtual void OnGuardar(ABMEventArgs e)
        {
            if (Guardar != null)
            {
                Guardar(this, e);
            }
        }

        protected virtual void OnEliminar(ABMEventArgs e)
        {
            if (Eliminar != null)
            {
                Eliminar(this, e);
            }
        }

        protected virtual void OnActualizarGrilla(ABMEventArgs e)
        {
            if (ActualizarGrilla != null)
            {
                ActualizarGrilla(this, e);
            }
        }
        #endregion


        public ReadDelegate ReadMethod { get; set;}
        private DeleteDelegate delete;
        private CreateDelegate create;
        private UpdateDelegate update;

        protected void Page_Load(object sender, EventArgs e)
        {
            lblError.Text = string.Empty;
        }

        public void LimpiarControles()
        {
            FormsHelper.ClearControls(this.tblABM, ObjetoDTO);
        }

        public void Inicializar(BusinessMapper.eEntities entidad)
        {
            ObjetoDTO       = DTOHelper.InstanciarObjetoPorNombreDeTabla(entidad.ToString());
            Entidad         = entidad;
            dao             = BusinessMapper.GetDaoByEntity(entidad); 
           
            this.create     = Create;
            this.update     = Update;
            this.ReadMethod = Read;
            this.delete     = Delete;

            List<ABMControl> controls = GetControlsByEntity(entidad);
            BuildControls(controls.ToArray());
        }

        private List<ABMControl> GetControlsByEntity(BusinessMapper.eEntities entidad)
        {
            List<ABMControl> controls = new List<ABMControl>();

            if (BusinessMapper.AbmConfigXmlPath == null || BusinessMapper.AbmConfigXmlPath == string.Empty)
                throw new Exception("Path del archivo AbmConfig.xml sin definir.");

            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(BusinessMapper.AbmConfigXmlPath);

            foreach (XmlNode nodeControl in xDoc.SelectSingleNode(
                "Entities/Entity[@EntityName='" + entidad.ToString() + "']").ChildNodes)
            {

                if (nodeControl.Attributes["ControlType"].Value == "ComboBox")
                {
                    controls.Add(new ABMControl()
                    {
                        ComboBox   = (ucComboBox)Page.LoadControl("~/Controls/ucComboBox.ascx"),
                        Title      = nodeControl.Attributes["Title"].Value,
                        FieldName  = nodeControl.Attributes["FieldName"].Value,
                        EntityName = nodeControl.Attributes["FieldName"].Value
                    });
                    controls[controls.Count - 1].ComboBox.Inicializar(nodeControl.Attributes["EntityName"].Value);
                }
                else if (nodeControl.Attributes["ControlType"].Value == "TextBox")
                {
                    controls.Add(new ABMControl()
                    {
                        TextBox   = new ASPxTextBox(),
                        Title     = nodeControl.Attributes["Title"].Value,
                        FieldName = nodeControl.Attributes["FieldName"].Value,
                    });
                }
                else if (nodeControl.Attributes["ControlType"].Value == "SpinEdit")
                {
                    controls.Add(new ABMControl()
                    {
                        Control   = new ASPxSpinEdit(),
                        Title     = nodeControl.Attributes["Title"].Value,
                        FieldName = nodeControl.Attributes["FieldName"].Value,
                    });
                }
                else if (nodeControl.Attributes["ControlType"].Value == "CheckBox")
                {
                    controls.Add(new ABMControl()
                    {
                        Control   = new ASPxCheckBox(),
                        Title     = nodeControl.Attributes["Title"].Value,
                        FieldName = nodeControl.Attributes["FieldName"].Value,
                    });
                }
                else if (nodeControl.Attributes["ControlType"].Value == "RadioButtonList")
                {
                    controls.Add(new ABMControl()
                    {
                        Control   = new ASPxRadioButtonList(),
                        Title     = nodeControl.Attributes["Title"].Value,
                        FieldName = nodeControl.Attributes["FieldName"].Value                        
                    });

                    controls[controls.Count - 1].Items = new List<Item>();
                    foreach (XmlNode nodeItem in nodeControl.SelectNodes("Items/Item"))
                    {
                        controls[controls.Count - 1].Items.Add(
                            new Item(nodeItem.Attributes["Value"].Value, nodeItem.Attributes["Name"].Value));
                    }
                }
            }
            return controls;
        }

        private void Create(object entity)
        {
            CRUDHelper.Create(entity, dao);
        }
        private dynamic Read(int id)
        {
            return dao.Read(id);
        }
        private void Update(object entity)
        {
            CRUDHelper.Update(entity, dao);
        }
        private void Delete(int id)
        {
            CRUDHelper.Delete(id, dao);
        }

        public void Inicializar(TablaBase objetoDTO, CreateDelegate create, ReadDelegate read, UpdateDelegate update, DeleteDelegate delete, params ABMControl[] controls)
        {
            ObjetoDTO = objetoDTO;

            this.create     = create;
            this.ReadMethod = read;
            this.update     = update;
            this.delete     = delete;

            BuildControls(controls);
        }

        private void BuildControls(ABMControl[] controls)
        {
            HtmlTableRow tr;
            HtmlTableCell td;

            tblControls.Rows.Clear();

            for (int i = 0; i < controls.Length; i++)
            {
                tr = new HtmlTableRow();

                td = new HtmlTableCell();
                td.Width = "20%";
                td.Controls.Add(new Literal() { Text = controls[i].Title });
                tr.Cells.Add(td);


                td = new HtmlTableCell();
                td.Align = "Left";
                if (controls[i].TextBox != null)
                {
                    controls[i].TextBox.Width = new Unit(100, UnitType.Percentage);
                    controls[i].TextBox.ID    = "tx" + controls[i].FieldName;
                    td.Controls.Add(controls[i].TextBox);
                }
                else if (controls[i].ComboBox != null)
                {
                    controls[i].ComboBox.Width = new Unit(100, UnitType.Percentage);
                    controls[i].ComboBox.ID    = "uc" + controls[i].FieldName;
                    td.Controls.Add(controls[i].ComboBox);
                }
                else if (controls[i].Control != null)
                {
                    if (controls[i].Control is ASPxSpinEdit)
                    {
                        ASPxSpinEdit spinEdit = (ASPxSpinEdit)controls[i].Control;
                        spinEdit.Width        = new Unit(100, UnitType.Percentage);
                        spinEdit.ID = "sp" + controls[i].FieldName;
                        td.Controls.Add(spinEdit);
                    }
                    if (controls[i].Control is ASPxCheckBox)
                    {
                        ASPxCheckBox checkBox = (ASPxCheckBox)controls[i].Control;
                        checkBox.Width        = new Unit(30, UnitType.Pixel);
                        checkBox.ID = "cb" + controls[i].FieldName;
                        td.Controls.Add(checkBox);
                    }
                    if (controls[i].Control is ASPxRadioButtonList)
                    {
                        ASPxRadioButtonList radioButtonList = (ASPxRadioButtonList)controls[i].Control;
                        radioButtonList.Width               = new Unit(100, UnitType.Percentage);
                        radioButtonList.ID = "rb" + controls[i].FieldName;
                        td.Controls.Add(radioButtonList);

                        if (controls[i].Items != null)
                        {
                            foreach (Item item in controls[i].Items)
                            {
                                radioButtonList.Items.Add(item.Name, item.Value);
                            }
                        }
                    }
                }
                tr.Cells.Add(td);

                tblControls.Rows.Add(tr);
            }
        }

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                FormsHelper.FillEntity(this.tblABM, ObjetoDTO);
                this.create(ObjetoDTO);

                this.Visible = false;
                OnAgregar(new ABMEventArgs());
                OnActualizarGrilla(new ABMEventArgs());
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
                //MsgError(ex);
            }
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            OnCancelar(new ABMEventArgs());

        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var entity = this.ReadMethod(Convert.ToInt32(this.Attributes["RecId"])); 
                FormsHelper.FillEntity(this.tblABM, entity);
                this.update(entity);

                this.Visible = false;
                OnGuardar(new ABMEventArgs());
                OnActualizarGrilla(new ABMEventArgs());
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
                if (this.Attributes["RecId"] != null)
                {
                    this.delete(Convert.ToInt32(this.Attributes["RecId"]));
                    this.Visible = false;
                    OnEliminar(new ABMEventArgs());
                    OnActualizarGrilla(new ABMEventArgs());
                }
            }
            catch (Exception ex)
            {
                FormsHelper.MsgError(lblError, ex);
            }
        }

        protected void btnCancel_Click1(object sender, EventArgs e)
        {
            this.Visible = false;
        }
    }

    public class ABMEventArgs : EventArgs
    {
    }

    public class ABMControl
    {
        public string Title        { get; set; }
        public string FieldName    { get; set; }
        public string EntityName   { get; set; }
        public ASPxTextBox TextBox { get; set; }
        public ucComboBox ComboBox { get; set; }
        public Control Control     { get; set; }
        public List<Item> Items    { get; set; }
    }

    public class Item
    {
        public string Value { get; set; }
        public string Name  { get; set; }

        public Item(string value, string name)
        {
            this.Value = value;
            this.Name  = name;
        }
    }

}