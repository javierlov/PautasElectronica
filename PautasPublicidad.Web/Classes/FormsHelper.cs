using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using PautasPublicidad.Business;
using System.Xml;
using DevExpress.Web.ASPxGridView;
using DevExpress.Web.ASPxGridView.Export;
using DevExpress.Web.ASPxMenu;
using PautasPublicidad.DTO;
using System.Globalization;
using System.Drawing;
using System.Web.UI.HtmlControls;
using System.Reflection;
using DevExpress.Web.ASPxEditors;
using PautasPublicidad.Web.Controls;

namespace PautasPublicidad.Web
{
    public static class FormsHelper
    {
        public enum eAccionABM
        {
            Add   ,
            Delete,
            Edit
        }

        private enum eGetOrSet
        {
            GetFromProperty,
            SetToProperty  ,
            ClearControls
        }

        public delegate void RefreshDelegate(DevExpress.Web.ASPxGridView.ASPxGridView gv);

        static public string GetWindowOpenScript(string url, string windowName)
        {
            return GetWindowOpenScript(url, windowName, 1024, 768);
        }

        static public string GetWindowOpenScript(string url, string windowName, int width, int height)
        {
            return "window.open('"
                + url
                + "' ,'" + windowName 
                + "','height=" + height + ",width=" + width + "');";
        }

        #region FindControl

        private static Control FindControlByName(string controlName, HtmlTable tblABM)
        {
            if (tblABM.FindControl(controlName) != null && tblABM.FindControl(controlName) is Control)
                return (Control)tblABM.FindControl(controlName);
            else
                return null;
        }

        private static ucComboBox FindComboBoxByName(string controlName, HtmlTable tblABM)
        {
            if (tblABM.FindControl("uc" + controlName) != null && tblABM.FindControl("uc" + controlName) is ucComboBox)
                return (ucComboBox)tblABM.FindControl("uc" + controlName);
            else
                return null;
        }

        private static ASPxTextBox FindTextBoxByName(string controlName, HtmlTable tblABM)
        {
            if (tblABM.FindControl("tx" + controlName) != null && tblABM.FindControl("tx" + controlName) is ASPxTextBox)
                return (ASPxTextBox)tblABM.FindControl("tx" + controlName);
            else
                return null;
        }

        private static HtmlTableCell FindTdByName(string controlName, HtmlTable tblABM)
        {
            if (tblABM.FindControl("td" + controlName) != null && tblABM.FindControl("td" + controlName) is HtmlTableCell)
                return (HtmlTableCell)tblABM.FindControl("td" + controlName);
            else
                return null;
        }

        private static HtmlTableRow FindTrByName(string controlName, HtmlTable tblABM)
        {
            if (tblABM.FindControl("tr" + controlName) != null && tblABM.FindControl("tr" + controlName) is HtmlTableRow)
                return (HtmlTableRow)tblABM.FindControl("tr" + controlName);
            else
                return null;
        }

        #endregion

        private static void SetTableCell(string controlName, HtmlTable tblABM, bool visible)
        {
            HtmlTableCell td = FindTdByName(controlName, tblABM); 
            if (td != null) 
                td.Visible = visible;
        }

        private static void SetTableRow(string controlName, HtmlTable tblABM, bool visible)
        {
            HtmlTableRow tr = FindTrByName(controlName, tblABM);
            if (tr != null)
                tr.Visible = visible;
        }

        internal static void ShowOrHideButtons(HtmlTable tblABM, eAccionABM accion)
        {
            switch (accion)
            {
                case eAccionABM.Add:
                    SetTableCell("Add", tblABM, true);
                    SetTableCell("Save", tblABM, false);
                    SetTableCell("Delete", tblABM, false);

                    SetTableRow("Msg", tblABM, false);
                    SetTableRow("Abm", tblABM, true);
                    break;

                case eAccionABM.Delete:
                    SetTableCell("Add", tblABM, false);
                    SetTableCell("Save", tblABM, false);
                    SetTableCell("Delete", tblABM, true);

                    SetTableRow("Msg", tblABM, true);
                    SetTableRow("Abm", tblABM, false);
                    break;

                case eAccionABM.Edit:
                    SetTableCell("Add", tblABM, false);
                    SetTableCell("Save", tblABM, true);
                    SetTableCell("Delete", tblABM, false);

                    SetTableRow("Msg", tblABM, false);
                    SetTableRow("Abm", tblABM, true);
                    break;
            }
        }

        internal static void FillControls(object entidad, HtmlTable tblABM)
        {
            GetOrSetProperties(tblABM, entidad, eGetOrSet.GetFromProperty);
        }

        internal static void FillEntity(HtmlTable tblABM, object entidad)
        {
            GetOrSetProperties(tblABM, entidad, eGetOrSet.SetToProperty);
        }

        internal static void ClearControls(HtmlTable tblABM, object entidad)
        {
            GetOrSetProperties(tblABM, entidad, eGetOrSet.ClearControls);
        }

        private static void GetOrSetProperties(HtmlTable tblABM, object entidad, eGetOrSet getOrSet)
        {
            ASPxTextBox tx;
            ucComboBox uc;
            PropertyInfo myPropInfo;
            Type entityType;
            PropertyInfo[] myPropertyInfo;

            //Recorro las properties de la entidad...
            //Por cada prop, busco en la tabla, un control tx[PROP], o uc[PROP] que coincida.
            //Si lo encuentro, lo cargo (Text o SelectedValue) con el valor de la propiedad.

            //Itero entre las propiedades de la entidad.
            entityType = entidad.GetType();
            myPropertyInfo = entityType.GetProperties((BindingFlags.Public | BindingFlags.Instance));
            for (int i = 0; i < myPropertyInfo.Length; i++)
            {
                tx = null;
                uc = null;
                myPropInfo = ((PropertyInfo)myPropertyInfo[i]);

                if (myPropInfo.CanRead)
                {
                    tx = FindTextBoxByName(myPropInfo.Name, tblABM);
                    if (tx != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            tx.Text = Convert.ToString(myPropInfo.GetValue(entidad, null));
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            myPropInfo.SetValue(entidad, tx.Text, null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            tx.Text = string.Empty;
                    }

                    uc = FindComboBoxByName(myPropInfo.Name, tblABM);
                    if (uc != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            uc.SelectedValue = Convert.ToString(myPropInfo.GetValue(entidad, null));
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            myPropInfo.SetValue(entidad, uc.SelectedValue, null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            uc.SelectedValue = string.Empty;
                    }

                    ASPxSpinEdit sp = (ASPxSpinEdit)FindControlByName("sp" + myPropInfo.Name, tblABM);
                    if (sp != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            if (myPropInfo.GetValue(entidad, null) != null)
                                sp.Value = Convert.ToDecimal(myPropInfo.GetValue(entidad, null));
                            else
                                sp.Value = null;
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            myPropInfo.SetValue(entidad, sp.Value, null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            sp.Value = null;
                    }

                    ASPxCheckBox cb = (ASPxCheckBox)FindControlByName("cb" + myPropInfo.Name, tblABM);
                    if (cb != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            cb.Checked = Convert.ToBoolean(myPropInfo.GetValue(entidad, null));
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            myPropInfo.SetValue(entidad, cb.Checked, null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            cb.Checked = false;
                    }

                    ASPxRadioButtonList rb = (ASPxRadioButtonList)FindControlByName("rb" + myPropInfo.Name, tblABM);
                    if (rb != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            rb.SelectedItem = rb.Items.FindByValue(Convert.ToString(myPropInfo.GetValue(entidad, null)).Trim());
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            if (rb.SelectedItem != null) 
                                myPropInfo.SetValue(entidad, rb.SelectedItem.Value, null);
                            else
                                myPropInfo.SetValue(entidad, null, null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            rb.SelectedIndex = -1;
                    }

                    ASPxTimeEdit te = (ASPxTimeEdit)FindControlByName("te" + myPropInfo.Name, tblABM);
                    if (te != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            if (myPropInfo.GetValue(entidad, null) != null)
                                te.DateTime = ConvertToDateTime((TimeSpan)myPropInfo.GetValue(entidad, null));
                            else
                                te.DateTime = DateTime.MinValue;
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            myPropInfo.SetValue(entidad, ConvertToTimeSpan(te.DateTime), null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            te.DateTime = DateTime.MinValue;
                    }

                    ASPxDateEdit de = (ASPxDateEdit)FindControlByName("de" + myPropInfo.Name, tblABM);
                    if (de != null)
                    {
                        if (getOrSet == eGetOrSet.GetFromProperty)
                            if (myPropInfo.GetValue(entidad, null) != null)
                                de.Date = Convert.ToDateTime(myPropInfo.GetValue(entidad, null));
                            else
                                de.Value = null;
                        else if (getOrSet == eGetOrSet.SetToProperty)
                            myPropInfo.SetValue(entidad, de.Date, null);
                        else if (getOrSet == eGetOrSet.ClearControls)
                            de.Value = null;
                    }
                }
            }
        }

        internal static bool HoraEsMayor(DateTime desde, DateTime hasta)
        {
            return ConvertToTimeSpan(hasta).CompareTo(ConvertToTimeSpan(desde)) == 1;
        }

        internal static bool FechaEsMayor(DateTime desde, DateTime hasta)
        {
            return hasta.CompareTo(desde) == 1;
        }

        internal static TimeSpan ConvertToTimeSpan(DateTime dateTime)
        {
            return new TimeSpan(dateTime.Hour, dateTime.Minute, dateTime.Second);
        }

        internal static DateTime ConvertToDateTime(TimeSpan timeSpan)
        {
            return new DateTime(1, 1, 1, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
        }

        internal static void InicializarPropsGrilla(ASPxGridView gv)
        {
            gv.KeyFieldName = "RecId";

            gv.SettingsBehavior.AllowSelectByRowClick = true;
            gv.SettingsBehavior.ColumnResizeMode = ColumnResizeMode.NextColumn;

            gv.Settings.ShowFilterRow = true;                           //Filtro simple.
            gv.Settings.ShowFilterRowMenu = true;
            gv.Settings.ShowFilterBar = GridViewStatusBarMode.Visible;  //Filtro super cheto.
            gv.Settings.ShowGroupPanel = true;

            gv.SettingsPager.AlwaysShowPager = true;
            gv.SettingsPager.PageSize = 20;

            foreach (var item in gv.Columns.OfType<GridViewDataComboBoxColumn>())
            {
                item.PropertiesComboBox.DropDownStyle = DropDownStyle.DropDownList;
                item.PropertiesComboBox.IncrementalFilteringMode = IncrementalFilteringMode.Contains;
                item.PropertiesComboBox.IncrementalFilteringDelay = 100;
            }
        }
        
        internal static bool IsSelectedRecId(int recId, ASPxGridView gv)
        {
            //Devuelvo si el RecId es uno de los seleccionados
            foreach (int auxRecId in gv.GetSelectedFieldValues("RecId"))
                if (auxRecId == recId)
                    return true;
            return false;
        }

        internal static int? GetSelectedId(ASPxGridView gv)
        {
            //Devuelvo el ultimo elegido o null
            int? i = null;
            foreach (int recId in gv.GetSelectedFieldValues("RecId"))
                i = recId;
            return i;
        }

        static internal void ToolBarClick(ucABM ucABM1, string itemName, ASPxGridView gv, ASPxGridViewExporter ASPxGridViewExporter1)
        {
            switch (itemName)
            {
                case "btnAdd":
                    ucABM1.LimpiarControles();
                    FormsHelper.ShowOrHideButtons(ucABM1.tablaABM, FormsHelper.eAccionABM.Add);

                    ucABM1.Visible = true;
                    ucABM1.HeaderText = "Agregar Registro";
                    break;

                case "btnEdit":
                    if (FormsHelper.GetSelectedId(gv) != null)
                    {
                        ucABM1.LimpiarControles();
                        var entity = ucABM1.ReadMethod(FormsHelper.GetSelectedId(gv).Value);
                        FormsHelper.FillControls(entity, ucABM1.tablaABM);
                        FormsHelper.ShowOrHideButtons(ucABM1.tablaABM, FormsHelper.eAccionABM.Edit);

                        ucABM1.Attributes.Add("RecId", entity.RecId.ToString());
                        ucABM1.Visible = true;
                        ucABM1.HeaderText = "Modificar Registro";
                    }
                    else
                    {
                        ucABM1.Visible = false;
                    }
                    break;

                case "btnDelete":
                    if (FormsHelper.GetSelectedId(gv) != null)
                    {
                        FormsHelper.ShowOrHideButtons(ucABM1.tablaABM, FormsHelper.eAccionABM.Delete);
                        ucABM1.Attributes.Add("RecId", FormsHelper.GetSelectedId(gv).ToString());
                        ucABM1.Visible = true;
                        ucABM1.HeaderText = "Eliminar Registros";
                    }
                    else
                    {
                        ucABM1.Visible = false;
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

        internal static GridViewDataComboBoxColumn BuildComboColumn(string caption, string fieldName, params string[] items)
        {
            GridViewDataComboBoxColumn gvc = new GridViewDataComboBoxColumn();

            gvc.Name = fieldName;
            gvc.Caption = caption;
            gvc.FieldName = fieldName;

            gvc.PropertiesComboBox.Items.Clear();
            for (int i = 0; i < items.Length; i+=2)
            {
                gvc.PropertiesComboBox.Items.Add(items[i], items[i+1]);
            }

            return gvc;
        }

        internal static GridViewDataComboBoxColumn BuildComboColumn(string caption, string fieldName, BusinessMapper.eEntities entity)
        {
            GridViewDataComboBoxColumn gvc = new GridViewDataComboBoxColumn();

            gvc.Name = fieldName;
            gvc.Caption = caption;
            gvc.FieldName = fieldName;

            var mapInfo = BusinessMapper.GetMapInfo(entity.ToString());

            gvc.PropertiesComboBox.TextField = mapInfo.EntityTextField;
            gvc.PropertiesComboBox.ValueField = mapInfo.EntityValueField;
            gvc.PropertiesComboBox.DataSource = mapInfo.DAOHandler.ReadAll("");

            return gvc;
        }

        internal static void BuildColumnsByEntity(BusinessMapper.eEntities entidad, ASPxGridView gv)
        {
            if (BusinessMapper.AbmConfigXmlPath == null || BusinessMapper.AbmConfigXmlPath == string.Empty)
                throw new Exception("Path del archivo AbmConfig.xml sin definir.");

            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(BusinessMapper.AbmConfigXmlPath);

            gv.Columns.Clear();

            if (xDoc.SelectSingleNode(
                "Entities/Entity[@EntityName='" + entidad.ToString() + "']") == null)
                throw new AbmConfigXMLException("No existe la configuración de mapeo para la entidad: " + entidad.ToString());


            foreach (XmlNode nodeControl in xDoc.SelectSingleNode(
                "Entities/Entity[@EntityName='" + entidad.ToString() + "']").ChildNodes)
            {

                if (nodeControl.Attributes["ControlType"].Value == "ComboBox")
                {
                    GridViewDataComboBoxColumn gvc = new GridViewDataComboBoxColumn();

                    gvc.Name = nodeControl.Attributes["FieldName"].Value;
                    gvc.Caption = nodeControl.Attributes["Title"].Value;
                    gvc.FieldName = nodeControl.Attributes["FieldName"].Value;

                    var mapInfo = BusinessMapper.GetMapInfo(nodeControl.Attributes["EntityName"].Value);

                    gvc.PropertiesComboBox.TextField = mapInfo.EntityTextField;
                    gvc.PropertiesComboBox.ValueField = mapInfo.EntityValueField;
                    gvc.PropertiesComboBox.DataSource = mapInfo.DAOHandler.ReadAll("");

                    gv.Columns.Add(gvc);
                }
                else if (nodeControl.Attributes["ControlType"].Value == "RadioButtonList")
                {
                    GridViewDataComboBoxColumn gvc = new GridViewDataComboBoxColumn();

                    gvc.Name = nodeControl.Attributes["FieldName"].Value;
                    gvc.Caption = nodeControl.Attributes["Title"].Value;
                    gvc.FieldName = nodeControl.Attributes["FieldName"].Value;

                    gvc.PropertiesComboBox.Items.Clear();
                    foreach (XmlNode item in nodeControl.ChildNodes[0].ChildNodes)
                    {
                        gvc.PropertiesComboBox.Items.Add(item.Attributes["Name"].Value, item.Attributes["Value"].Value);
                    }

                    gv.Columns.Add(gvc);
                }
                else if (nodeControl.Attributes["ControlType"].Value == "TextBox"
                    || nodeControl.Attributes["ControlType"].Value == "SpinEdit"
                    || nodeControl.Attributes["ControlType"].Value == "TimeEdit")
                {
                    GridViewDataTextColumn gvc = new GridViewDataTextColumn();

                    gvc.Name = nodeControl.Attributes["FieldName"].Value;
                    gvc.Caption = nodeControl.Attributes["Title"].Value;
                    gvc.FieldName = nodeControl.Attributes["FieldName"].Value;

                    gv.Columns.Add(gvc);
                }
                else if (nodeControl.Attributes["ControlType"].Value == "DateEdit")
                {
                    GridViewDataDateColumn gvc = new GridViewDataDateColumn();

                    gvc.Name = nodeControl.Attributes["FieldName"].Value;
                    gvc.Caption = nodeControl.Attributes["Title"].Value;
                    gvc.FieldName = nodeControl.Attributes["FieldName"].Value;

                    gv.Columns.Add(gvc);
                }
                else if (nodeControl.Attributes["ControlType"].Value == "CheckBox")
                {
                    GridViewDataCheckColumn gvc = new GridViewDataCheckColumn();

                    gvc.Name = nodeControl.Attributes["FieldName"].Value;
                    gvc.Caption = nodeControl.Attributes["Title"].Value;
                    gvc.FieldName = nodeControl.Attributes["FieldName"].Value;

                    gv.Columns.Add(gvc);
                }
            }
        }

        internal static void BuildColumns(ASPxGridView gv, params GridViewColumn[] columnas)
        {
            
        }

        internal static void FillDias(ListEditItemCollection items)
        {
            items.Clear();
            items.Add("Lunes", "LUNES");
            items.Add("Martes", "MARTES");
            items.Add("Miercoles", "MIERCOLES");
            items.Add("Jueves", "JUEVES");
            items.Add("Viernes", "VIERNES");
            items.Add("Sabado", "SABADO");
            items.Add("Domingo", "DOMINGO");
        }

        internal static void MsgError(Label lblError, Exception ex)
        {
            string entityName = "";
            if (ex.Message.ToLower().Contains("Violation of UNIQUE KEY constraint".ToLower()))
            {
                lblError.Text = "No puede ingresar un nuevo registro con su clave duplicada.";
            }
            else if (ex.Message.ToLower().Contains("FK".ToLower()))
            {
                try
                {
                    entityName = ex.Message.Substring(ex.Message.IndexOf("dbo."), ex.Message.IndexOf("\"", ex.Message.IndexOf("dbo.") + 1) - ex.Message.IndexOf("dbo.")).ToString();
                    entityName = string.Format(" ({0}).", entityName);
                }
                catch (Exception)
                {
                }
                lblError.Text = "No puede modificar la clave o eliminar este registro, ya que se encuentra relacionado a otra entidad " + entityName;
            }
            else
            {
                lblError.Text = ex.Message;
            }
        }
    }
}