using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DAO;
using PautasPublicidad.DTO;
using System.Data;
using System.Data.SqlClient;

namespace PautasPublicidad.Business
{
    public static class Conexion
    {
        static UsuariosDAO dao    = DAOFactory.Get<UsuariosDAO>();
        static EmpresaDAO daoEmpr = DAOFactory.Get<EmpresaDAO>();

        static public void SetConnectionString(string connStr)
        {
            dao.SetConnectionString(connStr);
        }

        static public void SetDatareaId(int datareaId)
        {
            dao.SetDatareaId(datareaId);
        }

        static public List<EmpresaDTO> GetEmpresas()
        {
            return daoEmpr.GetEmpresas();
        }

        static public bool TestConnection(out string error)
        {
            try 
	        {
                dao.GetSqlVersion();
                error = "";
                return true;
	        }
	        catch (Exception ex)
	        {
                error = ex.Message;
		        return false;
	        }
        }

        static public bool Login(string userName, string password, int datareaId, out UsuariosDTO user)
        {
            try
            {
                user = dao.Read(string.Format("UserName='{0}' AND Password='{1}' AND DatareaId = {2}", userName, password, datareaId));
                return (user != null);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
