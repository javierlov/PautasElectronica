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
    public static class PiezasArte
    {
        static PiezasArteDAO dao = DAOFactory.Get<PiezasArteDAO>();
        static PiezasArteSKUDAO daoDetalle = DAOFactory.Get<PiezasArteSKUDAO>();

        static public List<PiezasArteDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public PiezasArteDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public void Create(PiezasArteDTO piezaArte, List<PiezasArteSKUDTO> piezasArteSKU)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    piezaArte = dao.Create(piezaArte, tran);

                    foreach (PiezasArteSKUDTO piezaArteSKU in piezasArteSKU)
                    {
                        piezaArteSKU.RecId        = 0;
                        piezaArteSKU.DatareaId    = piezaArte.DatareaId;
                        piezaArteSKU.IdentifPieza = piezaArte.IdentifPieza;

                        if (piezaArteSKU.Coeficiente == 0)
                            piezaArteSKU.Coeficiente = null;

                        daoDetalle.Create(piezaArteSKU, tran);
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

        public static void Update(PiezasArteDTO piezaArte, List<PiezasArteSKUDTO> piezasArteSKU)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(piezaArte, piezaArte.RecId, tran);

                    //Elimino todos los productos y los re-creo.
                    daoDetalle.Delete(string.Format("IdentifPieza = '{0}' AND DatareaId = {1}", piezaArte.IdentifPieza, piezaArte.DatareaId), tran);

                    foreach (PiezasArteSKUDTO piezaArteSKU in piezasArteSKU)
                    {
                        piezaArteSKU.RecId = 0;
                        piezaArteSKU.DatareaId = piezaArte.DatareaId;
                        piezaArteSKU.IdentifPieza = piezaArte.IdentifPieza;

                        if (piezaArteSKU.Coeficiente == 0)
                            piezaArteSKU.Coeficiente = null;

                        daoDetalle.Create(piezaArteSKU, tran);
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
                    var piezaArte = dao.Read(recId);

                    //Elimino todos los detalles.
                    daoDetalle.Delete(string.Format("IdentifPieza = '{0}' AND DatareaId = {1}", piezaArte.IdentifPieza, piezaArte.DatareaId), tran);

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

        public static List<PiezasArteSKUDTO> ReadAllProductos(string identifPieza)
        {
            return daoDetalle.ReadAll(string.Format("IdentifPieza = '{0}'", identifPieza));
        }
    }
}
