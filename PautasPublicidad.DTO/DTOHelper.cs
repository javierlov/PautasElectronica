using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace PautasPublicidad.DTO
{
    public static class DTOHelper
    {
        private static Assembly miAssembly = Assembly.GetExecutingAssembly();

        public static TablaBase InstanciarObjetoPorNombreDeTabla(string NombreTabla)
        {
            //Paso del nombre de la ClaseDAO correspondiente a la entidad,
            //al nombre completo de la Clase a Instanciar.
            NombreTabla = "PautasPublicidad.DTO." + NombreTabla + "DTO";

            /* Casos especiales? */

            return (TablaBase)miAssembly.CreateInstance(NombreTabla);
        }

        public static string GetTableNameByType(Type t)
        {
            //Paso del nombre de la clase DAO o DTO, al nombre de la Tabla.
            string NombreTabla = t.Name.Replace("DAO", string.Empty).Replace("DTO", string.Empty);

            /* Casos especiales? */

            return NombreTabla;
        }

        public static void FillObjectByObject(object objOrigen, object objDestino)
        {
            object value;
            PropertyInfo[] estimadoCabPropertyInfo = objOrigen.GetType().GetProperties((BindingFlags.Public | BindingFlags.Instance));
            for (int i = 0; i < estimadoCabPropertyInfo.Length; i++)
            {
                PropertyInfo estimadoCabPropInfo = ((PropertyInfo)estimadoCabPropertyInfo[i]);
                PropertyInfo estimadoCabVersionPropInfo = objDestino.GetType().GetProperty(estimadoCabPropInfo.Name);
                
                try
                {
                    //Si existe la property y sus tipos coinciden, le seteo el valor.
                    if (estimadoCabVersionPropInfo != null && estimadoCabVersionPropInfo.PropertyType == estimadoCabPropInfo.PropertyType)
                    {
                        value = estimadoCabPropInfo.GetValue(objOrigen, null);
                        estimadoCabVersionPropInfo.SetValue(objDestino, value, null);
                    }
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
