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
    public static class Frecuencias
    {
        static FrecuenciaDAO dao           = DAOFactory.Get<FrecuenciaDAO>();
        static FrecuenciaDetDAO daoDetalle = DAOFactory.Get<FrecuenciaDetDAO>();

        static public void Create(FrecuenciaDTO frecuencia, List<FrecuenciaDetDTO> frecuenciaDetList)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    frecuencia = dao.Create(frecuencia, tran);

                    foreach (FrecuenciaDetDTO frecuenciaDet in frecuenciaDetList)
                    {
                        frecuenciaDet.RecId             = 0;
                        frecuenciaDet.DatareaId         = frecuencia.DatareaId;
                        frecuenciaDet.IdentifFrecuencia = frecuencia.IdentifFrecuencia;

                        daoDetalle.Create(frecuenciaDet, tran);
                    }

                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        public static void Update(FrecuenciaDTO frecuencia, List<FrecuenciaDetDTO> frecuenciaDetList)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(frecuencia, frecuencia.RecId, tran);

                    //Elimino todos los atencion y los re-creo.
                    daoDetalle.Delete(string.Format("IdentifFrecuencia = '{0}' AND DatareaId = {1}", frecuencia.IdentifFrecuencia, frecuencia.DatareaId), tran);

                    foreach (FrecuenciaDetDTO frecuenciaDet in frecuenciaDetList)
                    {
                        frecuenciaDet.RecId             = 0;
                        frecuenciaDet.DatareaId         = frecuencia.DatareaId;
                        frecuenciaDet.IdentifFrecuencia = frecuencia.IdentifFrecuencia;
                        daoDetalle.Create(frecuenciaDet, tran);
                    }

                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        public static void Delete(int recId)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    //Obtengo el registro a eliminar.
                    var frecuencia = dao.Read(recId);

                    //Elimino todos los atencion.
                    daoDetalle.Delete( string.Format("IdentifFrecuencia = '{0}' AND DatareaId = {1}", frecuencia.IdentifFrecuencia, frecuencia.DatareaId), tran);

                    //Elimino el registro cabecera.
                    dao.Delete(recId, tran);


                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }
    }
}
