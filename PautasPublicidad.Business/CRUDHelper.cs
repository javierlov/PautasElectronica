using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using PautasPublicidad.DTO;

namespace PautasPublicidad.Business
{
    public static class CRUDHelper
    {
        static public void Create(dynamic entidad, dynamic dao)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Create(entidad, tran);
                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        static public void Delete(int id, dynamic dao)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Delete(id, tran);
                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        static public void Update(dynamic entidad, dynamic dao)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(entidad, entidad.RecId, tran);
                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        static public dynamic ReadAll(string filtro, dynamic dao)
        {
            return dao.ReadAll(filtro);
        }

        static public dynamic Read(int id, dynamic dao)
        {
            return dao.Read(id);
        }

        static public dynamic Read(string filtro, dynamic dao)
        {
            return dao.Read(filtro);
        }
    }
}
