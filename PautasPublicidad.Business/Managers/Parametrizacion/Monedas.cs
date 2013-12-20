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
    public static class Monedas
    {
        static MonedasDAO dao = DAOFactory.Get<MonedasDAO>();
        static TipoCambioDAO daoDetalle = DAOFactory.Get<TipoCambioDAO>();

        static public List<MonedasDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public MonedasDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public void Create(MonedasDTO moneda, List<TipoCambioDTO> tiposDeCambio)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    moneda = dao.Create(moneda, tran);

                    foreach (TipoCambioDTO tipoCambio in tiposDeCambio)
                    {
                        tipoCambio.RecId      = 0;
                        tipoCambio.DatareaId  = moneda.DatareaId;
                        tipoCambio.IdentifMon = moneda.IdentifMon;
                        daoDetalle.Create(tipoCambio, tran);
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

        public static void Update(MonedasDTO moneda, List<TipoCambioDTO> tiposDeCambio)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(moneda, moneda.RecId, tran);

                    //Elimino todos los atencion y los re-creo.
                    daoDetalle.Delete(
                        string.Format("identifMon = '{0}' AND DatareaId = {1}",
                        moneda.IdentifMon, moneda.DatareaId),
                        tran);

                    foreach (TipoCambioDTO tipoCambio in tiposDeCambio)
                    {
                        tipoCambio.RecId       = 0;
                        tipoCambio.DatareaId   = moneda.DatareaId;
                        tipoCambio.IdentifMon  = moneda.IdentifMon;
                        daoDetalle.Create(tipoCambio, tran);
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
                    var moneda = dao.Read(recId);

                    //Elimino todos los detalles.
                    daoDetalle.Delete(
                        string.Format("identifMon = '{0}' AND DatareaId = {1}",
                        moneda.IdentifMon, moneda.DatareaId),
                        tran);

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

        public static List<TipoCambioDTO> ReadAllTiposCambio(string identifMon)
        {
            return daoDetalle.ReadAll(string.Format("identifMon = '{0}'", identifMon));
        }

        public static decimal GetTipoCambioValor(string identifMon)
        {
            var tiposCambio = daoDetalle.ReadAll(string.Format("IdentifMon='{0}'", identifMon));

            tiposCambio.Sort((x, y) => y.FechaInicio.CompareTo(x.FechaInicio));
            
            if (tiposCambio.Count > 0)
                return tiposCambio[0].Valor;
            else
                return 1;
        }
    }
}
