using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class UsuariosDAO : DAOBase<UsuariosDTO>
    {
        public void SetConnectionString(string connectionString)
        {
            DAOHelper.ConnStr = connectionString;
        }

        public void SetDatareaId(int datareaId)
        {
            DAOHelper.DatareaId = datareaId;
        }

        public string GetSqlVersion()
        {
            return Convert.ToString(DAOHelper.EjecutarScalar("SELECT SERVERPROPERTY('productversion')"));
        }

        public string GetConnectionString()
        {
            return DAOHelper.ConnStr;
        }
    
    }
}
