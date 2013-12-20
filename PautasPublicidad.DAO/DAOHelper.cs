using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Reflection.Emit;

namespace PautasPublicidad.DAO
{
    internal static class DAOHelper
    {
        //Conexion 'Singletoneada'.
        private static SqlConnection conexionSingleton = null;
        private static SqlTransaction trans;

        public static int DatareaId { get; set; }
        public static string ConnStr { get; set; }

        public static int GetNextId(string Tabla, string NombreClave, System.Data.SqlClient.SqlTransaction tran)
        {
            try
            {
                SqlCommand Comando = new SqlCommand();
                object Aux;
                int ret;

                PrepararConexion(Comando, tran);

                //Armo la Query
                Comando.CommandType = CommandType.Text;
                Comando.CommandText = "SELECT MAX(" + NombreClave + ") " + "FROM [" + Tabla + "] ";

                //Obtengo el Nuevo Id
                Aux = EjecutarScalar(Comando);

                //Si el resultado es un numero, le sumo 1 y lo devuelvo.
                //Sino, devuelvo 1.
                if ((bool)int.TryParse(Aux.ToString(), out ret))
                    return (ret + 1);
                else
                    return 1;
            }
            catch (Exception ex)
            {
                throw new Exception("Error en: DataLibrary - GetNextId", ex);
            }
        }

        public static void LlenarDataTable(ref DataTable DataTableDestino, string SelectQuery)
        {

            SqlCommand Comando = new SqlCommand();
            SqlDataAdapter dAdapter = new SqlDataAdapter();

            PrepararConexion(Comando);

            //Comando.Connection = Conexion
            Comando.CommandType = CommandType.Text;

            //Levanto el primer elemento solo por tener la estructura.
            Comando.CommandText = SelectQuery;

            dAdapter.SelectCommand = Comando;
            dAdapter.Fill(DataTableDestino);

        }

        public static void LlenarDataSet(ref DataSet DataSetDestino, string SelectQuery, string NombreTabla)
        {

            SqlCommand Comando = new SqlCommand();
            SqlDataAdapter dAdapter = new SqlDataAdapter();

            PrepararConexion(Comando);

            //Comando.Connection = Conexion
            Comando.CommandType = CommandType.Text;

            //Levanto el primer elemento solo por tener la estructura.
            Comando.CommandText = SelectQuery;

            dAdapter.SelectCommand = Comando;
            dAdapter.Fill(DataSetDestino, NombreTabla);

        }

        public static SqlDataReader EjecutarStoreProcedure(string CommandString)
        {
            SqlCommand Comando = new SqlCommand();
            SqlDataReader aux  =null;

            PrepararConexion(Comando);

            try
            {
                //Adecua la sentencia Segun el Tipo de Dato:
                Comando.CommandText = CommandString;

                //Comando.Connection = Conexion
                Comando.CommandType = CommandType.StoredProcedure;

                if (Comando.Connection.State != ConnectionState.Open)
                    Comando.Connection.Open();

                return Comando.ExecuteReader();
            }
            catch
            {
                return aux;
            }
        }
        public static int EjecutarNonQuery(string CommandString)
        {
            SqlCommand Comando = new SqlCommand();

            PrepararConexion(Comando);

            try
            {
                object aux;

                //Adecua la sentencia Segun el Tipo de Dato:
                Comando.CommandText = CommandString;

                //Comando.Connection = Conexion
                Comando.CommandType = CommandType.Text;

                if (Comando.Connection.State != ConnectionState.Open)
                    Comando.Connection.Open();

                aux = EjecutarNonQuery(Comando);
                return Convert.ToInt32(aux);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public static int EjecutarNonQuery(string CommandString, System.Data.SqlClient.SqlTransaction tran)
        {
            SqlCommand Comando = new SqlCommand();

            PrepararConexion(Comando, tran);

            try
            {
                object aux;

                //Adecua la sentencia Segun el Tipo de Dato:
                Comando.CommandText = CommandString;

                //Comando.Connection = Conexion
                Comando.CommandType = CommandType.Text;

                if (Comando.Connection.State != ConnectionState.Open)
                    Comando.Connection.Open();

                aux = EjecutarNonQuery(Comando);
                return Convert.ToInt32(aux);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public static int EjecutarNonQuery(SqlCommand Comando)
        {
            try
            {
                return Comando.ExecuteNonQuery();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static object EjecutarScalar(SqlCommand Comando)
        {
            try
            {
                return Comando.ExecuteScalar();
            }
            catch (Exception)
            {
                throw;
            }
        }


        public static object EjecutarScalar(string CommandString)
        {
            SqlCommand Comando = new SqlCommand();

            PrepararConexion(Comando);

            try
            {
                //Adecua la sentencia Segun el Tipo de Dato:
                Comando.CommandText = CommandString;

                //Comando.Connection = Conexion
                Comando.CommandType = CommandType.Text;

                if (Comando.Connection.State != ConnectionState.Open)
                    Comando.Connection.Open();

                return EjecutarScalar(Comando);
            }
            catch (Exception)
            {
                throw;
            }
        }


        public static SqlDataReader EjecutarReader(SqlCommand Comando)
        {
            try
            {
                return Comando.ExecuteReader();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void CrearParametros(SqlCommand Comando, params object[] Params)
        {

            //Creo los parametros
            Comando.Parameters.Clear();

            for (int i = 0; i <= Params.GetUpperBound(0); i += 3) //Era 2 sin el tipo.
            {
                Comando.Parameters.Add("@" + Params[i].ToString(), ObtenerTipo((Type)Params[i + 2]));

                if (Params[i + 1] == null)
                    Comando.Parameters["@" + Params[i].ToString()].Value = System.DBNull.Value;
                else
                    Comando.Parameters["@" + Params[i].ToString()].Value = Params[i + 1];


                if (Params[i + 1] == null)
                {
                    Comando.Parameters["@" + Params[i].ToString()].Value = System.DBNull.Value;
                }
                else
                {
                    if (((Params[i + 1]) is DateTime) && ((DateTime)Params[i + 1]) == DateTime.MinValue)
                        Comando.Parameters["@" + Params[i].ToString()].Value = System.DBNull.Value;
                    else
                        Comando.Parameters["@" + Params[i].ToString()].Value = Params[i + 1];
                }

                Comando.Parameters["@" + Params[i].ToString()].Direction = ParameterDirection.Input;
                Comando.Parameters["@" + Params[i].ToString()].SourceColumn = Params[i].ToString();
            }
        }


        public static void CrearUpdateQuery(SqlCommand Comando, string Tabla, string Where)
        {

            string sQuery = "UPDATE [" + Tabla + "]" + "\r\n" + "SET ";

            foreach (SqlParameter unParametro in Comando.Parameters)
            {
                if (unParametro.Direction == ParameterDirection.Input)
                {
                    sQuery = sQuery + unParametro.ParameterName.Substring(1) + "=" + unParametro.ParameterName + ",";
                }
            }

            sQuery = sQuery.Remove(sQuery.Length - 1, 1);
            sQuery = sQuery + "\r\n" + "WHERE " + Where;

            Comando.CommandType = CommandType.Text;
            Comando.CommandText = sQuery;
        }

        public static void CrearInsertQuery(SqlCommand Comando, string Tabla)
        {

            string sQuery = "INSERT INTO [" + Tabla + "]" + "\r\n" + "(";

            foreach (SqlParameter unParametro in Comando.Parameters)
            {
                if (unParametro.Direction == ParameterDirection.Input)
                {
                    sQuery = sQuery + unParametro.ParameterName.Substring(1) + ",";
                }
            }

            sQuery = sQuery.Remove(sQuery.Length - 1, 1);
            sQuery = sQuery + " )" + "\r\n" + "VALUES ( ";

            foreach (SqlParameter unParametro in Comando.Parameters)
            {
                if (unParametro.Direction == ParameterDirection.Input)
                {
                    sQuery = sQuery + unParametro.ParameterName + ",";
                }
            }

            sQuery = sQuery.Remove(sQuery.Length - 1, 1);
            sQuery = sQuery + " )";

            Comando.CommandType = CommandType.Text;
            Comando.CommandText = sQuery;
        }

        public static string CrearSelectQuery(string Tabla, string Where)
        {
            if (System.String.IsNullOrEmpty(Where))
                return "SELECT * FROM " + Tabla;
            else
                return "SELECT * FROM " + Tabla + " WHERE " + Where;
        }

        public static string CrearSelectQueryUnique(string Tabla, string Where)
        {
         return "SELECT TOP 1 * FROM " + Tabla + " WHERE " + Where;
        }
        
        public static string CrearSelectQuery(string Tabla, string Where, string OrderBy)
        {
            if (System.String.IsNullOrEmpty(OrderBy))
            {
                if (System.String.IsNullOrEmpty(Where))
                    return "SELECT * FROM " + Tabla;
                else
                    return "SELECT * FROM " + Tabla + " WHERE " + Where;
            }
            else
            {
                if (System.String.IsNullOrEmpty(Where))
                    return "SELECT * FROM " + Tabla + " ORDER BY " + OrderBy;
                else
                    return "SELECT * FROM " + Tabla + " WHERE " + Where + " ORDER BY " + OrderBy;
            }
        }

        public static void CrearSelectQuery(SqlCommand Comando, string Tabla, string Where)
        {
            if (System.String.IsNullOrEmpty(Where))
            {
                Comando.CommandText = "SELECT * FROM " + Tabla;
            }
            else
            {
                Comando.CommandText = "SELECT * FROM " + Tabla + " WHERE " + Where;
            }
        }

        public static void CrearSelectQuery(SqlCommand Comando, string Tabla, string Where, string OrderBy)
        {
            if (System.String.IsNullOrEmpty(Where))
            {
                Comando.CommandText = "SELECT * FROM " + Tabla + " ORDER BY " + OrderBy;
            }
            else
            {
                Comando.CommandText = "SELECT * FROM " + Tabla + " WHERE " + Where + " ORDER BY " + OrderBy;
            }
        }

        public static void CrearDeleteQuery(SqlCommand Comando, string Tabla, string Where)
        {
            if (System.String.IsNullOrEmpty(Where))
            {
                Comando.CommandText = "DELETE FROM " + Tabla;
            }
            else
            {
                Comando.CommandText = "DELETE FROM " + Tabla + " WHERE " + Where;
            }
        }

        public static System.Data.SqlDbType ObtenerTipo(Type tipo)
        {
            if (tipo == typeof(bool) || tipo == typeof(bool?))
            {
                return SqlDbType.Bit;
            }
            else if (tipo == typeof(string))
            {
                return SqlDbType.NVarChar;
            }
            else if (tipo == typeof(int) || tipo == typeof(int?))
            {
                return SqlDbType.Int;
            }
            else if (tipo == typeof(long) || tipo == typeof(long?))
            {
                return SqlDbType.BigInt;
            }
            else if (tipo == typeof(DateTime) || tipo == typeof(DateTime?))
            {
                return SqlDbType.DateTime;
            }
            else if (tipo == typeof(TimeSpan) || tipo == typeof(TimeSpan?))
            {
                return SqlDbType.Time;
            }
            else if (tipo == typeof(float) || tipo == typeof(float?))
            {
                return SqlDbType.Float;
            }
            else if (tipo == typeof(decimal) || tipo == typeof(decimal?))
            {
                return SqlDbType.Decimal;
            }
            else if (tipo == typeof(double) || tipo == typeof(double?))
            {
                return SqlDbType.Real;
            }
            else if (tipo == typeof(byte[]))
            {
                return SqlDbType.Image;
            }
            else throw new Exception("Err: Tipo de tipo Desconocido!");
        }

        #region "DataReader"

        internal static int GetValor(object dReaderItem, int ValorDefault)
        {
            if (dReaderItem is System.DBNull)
            {
                return ValorDefault;
            }
            else
            {
                return ((int)dReaderItem);
            }
        }

        internal static string GetValor(object dReaderItem, string ValorDefault)
        {
            if (dReaderItem is System.DBNull)
            {
                return ValorDefault;
            }
            else
            {
                return ((string)dReaderItem);
            }
        }

        internal static DateTime GetValor(object dReaderItem, DateTime ValorDefault)
        {
            if (dReaderItem is System.DBNull)
            {
                return ValorDefault;
            }
            else
            {
                return ((DateTime)dReaderItem);
            }
        }

        internal static float GetValor(object dReaderItem, float ValorDefault)
        {
            if (dReaderItem is System.DBNull)
            {
                return ValorDefault;
            }
            else
            {
                return ((float)dReaderItem);
            }
        }

        #endregion


        public static void PoblarObjetoDesdeDataRow(object obj, System.Data.DataRow dr)
        {
            //Obtengo el Tipo/Clase del objeto.
            int i;
            Type Objeto = obj.GetType();
            PropertyInfo[] myPropertyInfo = Objeto.GetProperties((BindingFlags.Public | BindingFlags.Instance));

            for (i = 0; i < myPropertyInfo.Length; i++)
            {
                PropertyInfo myPropInfo = ((PropertyInfo)myPropertyInfo[i]);

                //Solo voy a tener en cuenta las propiedades que no sean ReadOnly.
                //Así ignoro la 'TABLA'.
                if (myPropInfo.CanWrite)
                {
                    if (dr[myPropInfo.Name] is Single && (myPropInfo.PropertyType == typeof(decimal) || myPropInfo.PropertyType == typeof(decimal?)))
                    {
                        //Seteo el valor de la propiedad SOLO si no es nulo.
                        if (dr[myPropInfo.Name] != System.DBNull.Value)
                            myPropInfo.SetValue(obj, Convert.ToDecimal(dr[myPropInfo.Name]), null);
                    }
                    else if (dr[myPropInfo.Name] is TimeSpan && (myPropInfo.PropertyType == typeof(DateTime) || myPropInfo.PropertyType == typeof(DateTime?)))
                    {
                        if (dr[myPropInfo.Name] != System.DBNull.Value)
                            myPropInfo.SetValue(obj, (new DateTime().Add((TimeSpan)dr[myPropInfo.Name])), null);
                    }
                    else
                    {
                        //Seteo el valor de la propiedad SOLO si no es nulo.
                        if (dr[myPropInfo.Name] != System.DBNull.Value)
                            myPropInfo.SetValue(obj, dr[myPropInfo.Name], null);
                    }
                }
            }
        }

        public static object[] CrearParametrosDesdeObjeto(object obj)
        {
            //Obtengo el Tipo/Clase del objeto.
            Type Objeto = obj.GetType();

            int i;

            ArrayList Aux = new ArrayList();
            PropertyInfo[] myPropertyInfo = Objeto.GetProperties((BindingFlags.Public | BindingFlags.Instance));

            Aux.Clear();

            for (i = 0; i < myPropertyInfo.Length; i++)
            {
                PropertyInfo myPropInfo = ((PropertyInfo)myPropertyInfo[i]);

                //Solo voy a tener en cuenta las propiedades que no sean ReadOnly.
                //Así ignoro la 'TABLA'.
                if (myPropInfo.CanWrite)
                {
                    Aux.Add(myPropInfo.Name);
                    Aux.Add(myPropInfo.GetValue(obj, null));
                    Aux.Add(myPropInfo.PropertyType);
                }
            }
            return Aux.ToArray();
        }

        public static bool PrepararConexion(SqlCommand cmd)
        {
            try
            {
                if (trans != null && trans.Connection != null)
                {
                    cmd.Transaction = trans;
                    cmd.Connection = trans.Connection;
                }
                else
                {
                    if (conexionSingleton == null)
                        conexionSingleton = new SqlConnection(ConnStr);

                    if (conexionSingleton.State != ConnectionState.Open)
                        conexionSingleton.Open();

                    cmd.Connection = conexionSingleton;
                }

                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool PrepararConexion(SqlCommand cmd, SqlTransaction tran)
        {
            try
            {
                if (tran == null)
                    tran = IniciarTransaccion();

                if (tran.Connection.State != ConnectionState.Open)
                    tran.Connection.Open();

                cmd.Connection = tran.Connection;
                cmd.Transaction = tran;

                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void LiberarConexion()
        {
            if (conexionSingleton != null)
                conexionSingleton.Dispose();
                conexionSingleton = null;
        }

        public static SqlTransaction IniciarTransaccion()
        {
            try
            {
                if (conexionSingleton == null)
                {
                    conexionSingleton = new SqlConnection(ConnStr);
                    if (conexionSingleton.State != ConnectionState.Open)
                        conexionSingleton.Open();
                }

                //Si la transaccion ya esta activa, no puedo iniciar una nueva. Lanzo una Excepcion.
                if (trans != null && trans.Connection != null)
                    throw new Exception("Error: Se intentó iniciar una nueva transacción, con una transacción actualmente en ejecución.");

                trans = conexionSingleton.BeginTransaction();
                return trans;
            }
            catch (Exception ex)
            { throw new Exception("IniciarTransaccion - DataLibrary", ex); }
        }

        public static void CommitTransaccion(SqlTransaction tran)
        {
            try
            {
                tran.Commit();
            }
            catch (Exception ex)
            { throw new Exception("CommitTransaccion - DataLibrary", ex); }
        }

        public static void RollbackTransaccion(SqlTransaction tran)
        {
            try
            {
                tran.Rollback();
            }
            catch (Exception ex)
            { throw new Exception("RollbackTransaccion - DataLibrary", ex); }
        }

        public static decimal StN(string numero)
        {
            //Obtengo la cultura 'regional' actual.
            System.Globalization.CultureInfo c = System.Globalization.CultureInfo.CurrentCulture;

            if (numero.Trim() == "")
                return 0;

            decimal resultado;

            //Reemplazo el separador de grupos por el separador de miles.
            numero = numero.Replace(c.NumberFormat.NumberGroupSeparator, c.NumberFormat.NumberDecimalSeparator);

            if (Decimal.TryParse(numero, out resultado))
                return resultado;
            else
                return 0;
        }

        public static decimal StN(object numero)
        {
            //Obtengo la cultura 'regional' actual.
            System.Globalization.CultureInfo c = System.Globalization.CultureInfo.CurrentCulture;

            if (numero == null)
                return 0;

            if (numero.ToString().Trim() == "")
                return 0;

            decimal resultado;

            //Reemplazo el separador de grupos por el separador de miles.
            numero = numero.ToString().Replace(c.NumberFormat.NumberGroupSeparator, c.NumberFormat.NumberDecimalSeparator);

            if (Decimal.TryParse(numero.ToString(), out resultado))
                return resultado;
            else
                return 0;
        }

        internal static string GetIP()
        {
            var hostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
            return (from addr in hostEntry.AddressList
                    where addr.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork
                    select addr.ToString()
                    ).FirstOrDefault();
        }

        internal static string CrearOrderByDesdeObjeto(object o)
        {
            string orderBy = string.Empty;

            System.Reflection.MemberInfo inf = o.GetType(); // typeof(o);
            object[] attributes;
            attributes =
               inf.GetCustomAttributes(
                    typeof(PautasPublicidad.DTO.SortFieldAttribute), false);

            foreach (Object attribute in attributes)
            {
                orderBy += ((PautasPublicidad.DTO.SortFieldAttribute)attribute).Property + ",";
            }
            if (orderBy.Length > 0)
                orderBy = orderBy.TrimEnd(',');

            return orderBy;
        }
    }
}
