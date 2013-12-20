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
    public static class Avisos
    {
        static AvisosDAO dao              = DAOFactory.Get<AvisosDAO>();
        static AvisosIdAtenDAO daoDetalle = DAOFactory.Get<AvisosIdAtenDAO>();

        static public List<AvisosDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public AvisosDTO Read(string sWhere)
        {
            return dao.Read(sWhere);
        }

        static public AvisosDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public void Create(AvisosDTO aviso, List<AvisosIdAtenDTO> atencionList)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    aviso = dao.Create(aviso, tran);

                    foreach (AvisosIdAtenDTO atencion in atencionList)
                    {
                        atencion.RecId           = 0;
                        atencion.DatareaId       = aviso.DatareaId;
                        atencion.IdentifAviso    = aviso.IdentifAviso;
                        atencion.IdentifIdentAte = atencion.IdentifIdentAte;
                        daoDetalle.Create(atencion, tran);
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

        public static void Update(AvisosDTO aviso, List<AvisosIdAtenDTO> atencionList)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                { 
                    dao.Update(aviso, aviso.RecId, tran);

                    //Elimino todos los atencion y los re-creo.
                    daoDetalle.Delete(string.Format("IdentifAviso = '{0}' AND DatareaId = {1}", aviso.IdentifAviso, aviso.DatareaId), tran);

                    foreach (AvisosIdAtenDTO atencion in atencionList)
                    {
                        atencion.RecId        = 0;
                        atencion.DatareaId    = aviso.DatareaId;
                        atencion.IdentifAviso = aviso.IdentifAviso;
                        daoDetalle.Create(atencion, tran);
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

        static public void Delete(int recId)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    //Obtengo el registro a eliminar.
                    var aviso = dao.Read(recId);

                    //Elimino todos los detalles.
                    daoDetalle.Delete(string.Format("IdentifAviso = '{0}' AND DatareaId = {1}", aviso.IdentifAviso, aviso.DatareaId), tran);

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

        public static List<AvisosIdAtenDTO> ReadAllAtencion(string identifAviso)
        {
            return daoDetalle.ReadAll(string.Format("identifAviso = '{0}'", identifAviso));
        }
    }
}
