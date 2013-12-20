using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;
using PautasPublicidad.DTO;

namespace PautasPublicidad.DAO
{
    public abstract class DAOBase<T> where T : TablaBase
    {
        private string TableName = DTO.DTOHelper.GetTableNameByType(typeof(T));

        #region Transacciones
        public SqlTransaction IniciarTransaccion()
        {
            return DAOHelper.IniciarTransaccion();
        }
        public void RollbackTransaccion(SqlTransaction tran)
        {
            if (tran != null)
                DAOHelper.RollbackTransaccion(tran);
        }
        public void CommitTransaccion(SqlTransaction tran)
        {
            DAOHelper.CommitTransaccion(tran);
        }
        #endregion

        #region Operaciones CRUD

        public T Create(T dto, SqlTransaction tran)
        {
            try
            {
                //Creo un nuevo comando.
                SqlCommand cmd = new SqlCommand();

                //Abro una nueva conexion asignada al comando.
                DAOHelper.PrepararConexion(cmd, tran);

                //Si aún no tenemos el Id del nuevo Item, lo calculamos.
                dto.RecId = DAOHelper.GetNextId(TableName, "RecId", tran);

                //Si aún no tenemos el DatareaId, lo seteamos.
                dto.DatareaId = DAOHelper.DatareaId;

                //Creo los parametros en base a las Propiedades del objeto.
                DAOHelper.CrearParametros(cmd, DAOHelper.CrearParametrosDesdeObjeto(dto));

                //Creo la Query en base a los parametros existentes.
                DAOHelper.CrearInsertQuery(cmd, TableName);

                //Ejecuto la query.
                DAOHelper.EjecutarNonQuery(cmd);

                return dto;
            }
            catch (Exception e)
            {
                //ToDo: Motor de Loggin.
                throw e;
            }
        }

        public T Read(int id)
        {
            return Read(" RecId=" + id.ToString());
        }


        public T ReadUnique(string sWhere)
        {
            try
            {
               sWhere        = string.Format("DatareaId={0}", DAOHelper.DatareaId) + " AND " + sWhere;
                T o          = DTO.DTOHelper.InstanciarObjetoPorNombreDeTabla(TableName) as T        ;
                DataTable dt = GetTable(DAOHelper.CrearSelectQueryUnique(TableName, sWhere))         ;

                if (dt.Rows.Count >= 1)
                    DAOHelper.PoblarObjetoDesdeDataRow(o, dt.Rows[0]);
                else
                    o = null;

                return o;
            }
            catch (Exception e)
            {
                //ToDo: Motor de Loggin.
                throw e;
            }
        }

        public T Read(string sWhere)
        {
            try
            {
                //Agrego el filtro por empresa:
                if (sWhere.Trim() == string.Empty)
                    sWhere = string.Format(" DatareaId={0}", DAOHelper.DatareaId);
                else
                    sWhere += string.Format(" AND DatareaId={0}", DAOHelper.DatareaId);

                T o = DTO.DTOHelper.InstanciarObjetoPorNombreDeTabla(TableName) as T;
                DataTable dt = GetTable(DAOHelper.CrearSelectQuery(TableName, sWhere));

                if (dt.Rows.Count >= 1)
                    DAOHelper.PoblarObjetoDesdeDataRow(o, dt.Rows[0]);
                else
                    o = null;

                return o;
            }
            catch (Exception e)
            {
                //ToDo: Motor de Loggin.
                throw e;
            }
        }

        public List<T> ReadAll()
        {
            return ReadAll(string.Empty);
        }

        public List<T> ReadAll(string sWhere)
        {
            try
            {
                //Lista de Entidades a retornar.
                List<T> r = new List<T>();
                //T o;

                T o = DTO.DTOHelper.InstanciarObjetoPorNombreDeTabla(TableName) as T;
                string orderBy = DAOHelper.CrearOrderByDesdeObjeto(o);

                //Agrego el filtro por empresa:
                if (sWhere.Trim() == string.Empty)
                    sWhere = string.Format(" DatareaId={0}", DAOHelper.DatareaId);
                else
                    sWhere += string.Format(" AND DatareaId={0}", DAOHelper.DatareaId);

                //Recorro cada registro, instancio el objeto,
                //cargo sus propiedades, lo agrego a la coleccion 'ret'.
                foreach (System.Data.DataRow unDr in GetTable(DAOHelper.CrearSelectQuery(TableName, sWhere, orderBy)).Rows)
                {
                    o = DTO.DTOHelper.InstanciarObjetoPorNombreDeTabla(TableName) as T;
                    DAOHelper.PoblarObjetoDesdeDataRow(o, unDr);
                    r.Add(o);
                }
                return r;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void Update(T dto, int id, SqlTransaction tran)
        {
            Update(dto, " RecId=" + id.ToString(), tran);
        }

        public void Update(T dto, string sWhere, SqlTransaction tran)
        {
            try
            {
                //Creo un nuevo comando.
                SqlCommand cmd = new SqlCommand();

                //Abro una nueva conexion asignada al comando.
                DAOHelper.PrepararConexion(cmd, tran);

                //Si aún no tenemos el Id del nuevo Item, lo calculamos.
                if (dto.RecId <= 0)
                    dto.RecId = DAOHelper.GetNextId(TableName, "RecId", tran);

                //Si aún no tenemos el DatareaId, lo seteamos.
                if (dto.DatareaId == 0)
                    dto.DatareaId = DAOHelper.DatareaId;

                //Creo los parametros en base a las Propiedades del objeto.
                DAOHelper.CrearParametros(cmd, DAOHelper.CrearParametrosDesdeObjeto(dto));

                ////Elimino el Id (PK de la Tabla) de los parametros.
                //cmd.Parameters.RemoveAt("@Id");

                //Agrego el filtro por empresa:
                if (sWhere.Trim() == string.Empty)
                    sWhere = string.Format(" DatareaId={0}", DAOHelper.DatareaId);
                else
                    sWhere += string.Format(" AND DatareaId={0}", DAOHelper.DatareaId);

                //Creo la Query en base a los parametros existentes.
                DAOHelper.CrearUpdateQuery(cmd, TableName, sWhere);

                //Ejecuto la query.
                DAOHelper.EjecutarNonQuery(cmd);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void Delete(int id, SqlTransaction tran)
        {
            Delete(" RecId=" + id.ToString(), tran);
        }

        public void Delete(string sWhere, SqlTransaction tran)
        {
            try
            {
                //Creo un nuevo comando.
                SqlCommand cmd = new SqlCommand();

                //Abro una nueva conexion asignada al comando.
                DAOHelper.PrepararConexion(cmd, tran);

                //Agrego el filtro por empresa:
                if (sWhere.Trim() == string.Empty)
                    sWhere = string.Format(" DatareaId={0}", DAOHelper.DatareaId);
                else
                    sWhere += string.Format(" AND DatareaId={0}", DAOHelper.DatareaId);

                //Creo la Query en base a los parametros existentes.
                DAOHelper.CrearDeleteQuery(cmd, TableName, sWhere);

                //Ejecuto la query.
                DAOHelper.EjecutarNonQuery(cmd);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        #region Metodos Internos

        internal DataTable GetTable(string sQuery)
        {
            try
            {
                //Creo un nuevo comando.
                SqlCommand cmd = new SqlCommand();

                //Preparo el DataAdapter y DataTable para retornar los datos.
                SqlDataAdapter da = new SqlDataAdapter();
                DataTable dt = new DataTable();

                da.SelectCommand = cmd;

                //Abro una nueva conexion asignada al comando.
                DAOHelper.PrepararConexion(cmd);

                cmd.CommandText = sQuery;

                da.Fill(dt);

                return dt;
            }
            catch (Exception)
            {
                throw;
            }
        }

        internal int GetLastId()
        {
            object aux = DAOHelper.EjecutarScalar("SELECT MAX(Id) FROM " + TableName);
            if (aux is DBNull)
                return 1;
            else
                return Convert.ToInt32(aux);
        }

        #endregion
    }
}
