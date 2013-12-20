using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DAO;
using PautasPublicidad.DTO;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Collections;

namespace PautasPublicidad.Business
{
    public static class Estimados
    {
        static EstimadoCabDAO dao        = DAOFactory.Get<EstimadoCabDAO>();
        static EstimadoDetDAO daoDetalle = DAOFactory.Get<EstimadoDetDAO>();
        static EstimadoSKUDAO daoSKU     = DAOFactory.Get<EstimadoSKUDAO>();

        static EstimadoCabVersionDAO daoVer        = DAOFactory.Get<EstimadoCabVersionDAO>();
        static EstimadoDetVersionDAO daoDetalleVer = DAOFactory.Get<EstimadoDetVersionDAO>();
        static EstimadoSKUVersionDAO daoSKUVer     = DAOFactory.Get<EstimadoSKUVersionDAO>();

        static OrdenadoCabDAO daoOrdenadoCab = DAOFactory.Get<OrdenadoCabDAO>();

        static public EstimadoCabDTO Buscar(string identifEspacio, int año, int mes)
        {
            int anoMes = Convert.ToInt32(año.ToString() + mes.ToString("00"));
            return dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1}", identifEspacio, anoMes));
        }

        static public EstimadoCabDTO Buscar(string paitaId)
        {
            return dao.Read(string.Format("PautaId = '{0}'", paitaId));
        }

        static private void Create(Estimado estimado, SqlTransaction tran)
        {
            try
            {
                dao.Create(estimado.Cabecera, tran);

                foreach (var linea in estimado.Lineas)
                {
                    linea.RecId     = 0;
                    linea.DatareaId = estimado.Cabecera.DatareaId;
                    linea.PautaId   = estimado.Cabecera.PautaId;

                    daoDetalle.Create(linea, tran);
                }

                foreach (var sku in estimado.SKUs)
                {
                    sku.RecId     = 0;
                    sku.DatareaId = estimado.Cabecera.DatareaId;
                    sku.PautaId   = estimado.Cabecera.PautaId;

                    daoSKU.Create(sku, tran);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        static private void Create(Business.EstimadoVersion estimadoVer, SqlTransaction tran)
        {
            try
            {
                daoVer.Create(estimadoVer.Cabecera, tran);

                foreach (var linea in estimadoVer.Lineas)
                {
                    linea.RecId     = 0;
                    linea.DatareaId = estimadoVer.Cabecera.DatareaId;
                    linea.PautaId   = estimadoVer.Cabecera.PautaId;

                    daoDetalleVer.Create(linea, tran);
                }

                foreach (var sku in estimadoVer.SKUs)
                {
                    sku.RecId     = 0;
                    sku.DatareaId = estimadoVer.Cabecera.DatareaId;
                    sku.PautaId   = estimadoVer.Cabecera.PautaId;

                    daoSKUVer.Create(sku, tran);
                }
            }
            catch (Exception)
            {
                throw;
            }            
        }

        public static void CierreOrdenado(Estimado estimado, EstimadoVersion estimadoVer, OrdenadoCabDTO ordenadoCab)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    Create(estimado, tran);
                    Create(estimadoVer, tran);
                    daoOrdenadoCab.Update(ordenadoCab, ordenadoCab.RecId, tran);

                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        public static Estimado Get(int recId)
        {
            var estimado = new Estimado();

            estimado.Cabecera = Read(recId);
            estimado.Lineas   = daoDetalle.ReadAll(string.Format("PautaId='{0}'", estimado.Cabecera.PautaId));
            estimado.SKUs     = daoSKU.ReadAll(string.Format("PautaId='{0}'", estimado.Cabecera.PautaId));

            return estimado;
        }

        public static EstimadoCabDTO Read(int recId)
        {
            return dao.Read(recId);
        }

        public static DataTable BuildAllSKU(List<EstimadoDetDTO> lineas)
        {
            DataTable dt = daoSKU.GetSKUByPiezasArte(lineas);

            foreach (DataRow dr in dt.Rows)
                dr["CantSalidas"] = lineas.FindAll(x => x.IdentifAviso == (string)dr["IdentifAviso"]).Count;

            return dt;
        }

        public static Estimado Confirmar(EstimadoCabDTO estimadoCab, List<EstimadoDetDTO> lineas, List<EstimadoSKUDTO> skus)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    Estimado estimado = new Estimado();
                    estimado.Cabecera = estimadoCab;
                    estimado.Lineas   = lineas;
                    estimado.SKUs     = skus;


                    dao.Update(estimado.Cabecera, estimado.Cabecera.RecId, tran);

                    EstimadoCabVersionDTO cabVer = new EstimadoCabVersionDTO(estimado.Cabecera);
                    cabVer.RecId = 0;
                    daoVer.Create(cabVer, tran);


                    //Elimino toda las lineas del estimado y las re-creo.
                    daoDetalle.Delete(
                        string.Format("PautaId = '{0}'",
                        estimado.Cabecera.PautaId),
                        tran);

                    foreach (var det in estimado.Lineas)
                    {
                        det.RecId     = 0;
                        det.DatareaId = estimadoCab.DatareaId;
                        det.PautaId   = estimadoCab.PautaId;
                        daoDetalle.Create(det, tran);

                        EstimadoDetVersionDTO detVer = new EstimadoDetVersionDTO(det);
                        detVer.Version = estimado.Cabecera.Version;
                        detVer.RecId = 0;
                        daoDetalleVer.Create(detVer, tran);
                    }

                    //Elimino toda las lineas del estimado y las re-creo.
                    daoSKU.Delete(
                        string.Format("PautaId = '{0}'",
                        estimado.Cabecera.PautaId),
                        tran);

                    var dtSKU = BuildAllSKU(lineas);
                    foreach (System.Data.DataRow dr in dtSKU.Rows)
                    {
                        var sku = new EstimadoSKUDTO();

                        sku.RecId     = 0;
                        sku.DatareaId = estimado.Cabecera.DatareaId;
                        sku.PautaId   = estimado.Cabecera.PautaId;

                        if (dr["Duracion"] != DBNull.Value)
                            sku.Duracion    = Convert.ToDecimal(dr["Duracion"]);
                        else
                            sku.Duracion = null;

                        sku.CantSalidas  = Convert.ToDecimal(dr["CantSalidas"]);
                        sku.IdentifAviso = Convert.ToString(dr["IdentifAviso"]);
                        sku.IdentifSKU   = Convert.ToString(dr["IdentifSKU"]);

                        daoSKU.Create(sku, tran);

                        EstimadoSKUVersionDTO skuVer = new EstimadoSKUVersionDTO(sku);
                        skuVer.Version               = estimado.Cabecera.Version;
                        skuVer.RecId                 = 0;
                        daoSKUVer.Create(skuVer, tran);

                    }

                    dao.CommitTransaccion(tran);

                    return estimado;
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        public static void CalcularCosto(EstimadoCabDTO EstimadoCAB, CostosDTO costoCab, List<EstimadoDetDTO> lineas, string usuario)
        {
            decimal costo = 0;
            decimal costoAcum = 0;

            //Lockeo la seccion para que otro thread no me descontrole los calculos.
            lock (typeof(Estimado))
            {
                using (SqlTransaction tran = dao.IniciarTransaccion())
                {
                    try
                    {
                        //La recargo para asegurarme de tener la version mas reciente...
                        EstimadoCAB = dao.Read(string.Format("PautaId='{0}'", EstimadoCAB.PautaId));

                        if (costoCab == null)
                        {
                            throw new Exception("No hay costos confirmados para realizar el costeo de la Pauta: " + EstimadoCAB.PautaId.ToString());
                        }

                        if (EstimadoCAB.UsuCierre != "") //no calcular costo si ya esta cerrado. 
                        {
                            throw new Exception("");
                        }

                        //Inicializo en '0' los registros de trabajo.
                        dao.MoverCeros(EstimadoCAB.PautaId, tran);
                        daoDetalle.MoverCeros(EstimadoCAB.PautaId, tran);
                        daoSKU.MoverCeros(EstimadoCAB.PautaId, tran);

                        //•	Seleccionar la Tabla CostoProveedorVersion 
                        var costosProveedorVersion = Costos.ReadAllProveedorVersiones(EstimadoCAB.IdentifEspacio, costoCab.VigDesde, costoCab.VigHasta, costoCab.Version.Value);

                        //Por cada registro seleccionado:
                        foreach (var costoProveedorVersion in costosProveedorVersion)
                        {
                            //•	Calcular Tipo de cambio
                            var tipoCambioValor = Monedas.GetTipoCambioValor(costoProveedorVersion.IdentifMon);
                            var costoCur        = costoProveedorVersion.Costo * tipoCambioValor;

                            //•	Calcular GrossingUp
                            var costoGros = costoCur * costoProveedorVersion.GrossingUp;

                            //•	Actualizar Tabla DET
                            var estimadoDetalles = daoDetalle.ReadAll(string.Format("PautaId='{0}'", EstimadoCAB.PautaId));
                            foreach (var estimadoDET in estimadoDetalles)
                            {
                                if (costoProveedorVersion.TipoCosto == "FIJO_MENSUAL")
                                {
                                    if (EstimadoCAB.DuracionTot > 0)
                                    {
                                        if (estimadoDET.Duracion.Value > 0)
                                        {
                                            costo = (costoGros / EstimadoCAB.DuracionTot) * estimadoDET.Duracion.Value;
                                        }
                                        else
                                        {
                                            decimal? divisor = 0;

                                            costo = (costoGros / EstimadoCAB.DuracionTot) * divisor.Value;
                                        }

                                    }
                                    else
                                    {
                                        costo = costoGros;
                                    }

                                }
                                else if (costoProveedorVersion.TipoCosto == "SEGUNDO_FIJO")
                                {
                                    costo = costoGros * estimadoDET.Duracion.Value;
                                }
                                else if (costoProveedorVersion.TipoCosto == "SALIDA")
                                {
                                    costo = costoGros;
                                }
                                else if (costoProveedorVersion.TipoCosto == "UNIDAD_PAUTADA")
                                {
                                    costo = costoGros;
                                }
                                else
                                {
                                    throw new Exception("Proveedor.TipoCosto Desconocido.");
                                }

                                estimadoDET.Costo += costo;

                                if (estimadoDET.Duracion > 0)
                                {
                                    estimadoDET.CostoUni = estimadoDET.Costo / estimadoDET.Duracion.Value;
                                }
                                else
                                {
                                    estimadoDET.CostoUni = estimadoDET.Costo;
                                }
                                

                                if (costoProveedorVersion.IncluidoOP)
                                {
                                    estimadoDET.CostoOp += costo;

                                    if (estimadoDET.Duracion > 0)
                                    {
                                        estimadoDET.CostoOpUni = estimadoDET.CostoOp / estimadoDET.Duracion.Value;
                                    }
                                    else
                                    {
                                        estimadoDET.CostoOpUni = estimadoDET.CostoOp;
                                    }

                                }

                                costoAcum += costo;

                                //o	Actualizar la Tabla DET
                                daoDetalle.Update(estimadoDET, estimadoDET.RecId, tran);
                            }

                            //•	Actualizar Tabla SKU
                            foreach (var estimadoDET in estimadoDetalles)
                            {
                                //o	Seleccionar la Tabla SKU con SKU.PautaID = PautaId enviado y SKU.IdentifAviso = DET.IdentifAviso
                                var ordenadoSKUs = daoSKU.ReadAll(string.Format("PautaId='{0}' AND IdentifAviso='{1}'", EstimadoCAB.PautaId, estimadoDET.IdentifAviso));

                                //o	Seleccionar la Tabla Avisos con IdentifAviso... >>
                                var aviso              = Avisos.Read(string.Format("IdentifAviso='{0}'", estimadoDET.IdentifAviso));
                                var productosPiezaArte = PiezasArte.ReadAllProductos(aviso.IdentifPieza);
                                foreach (var ordenadoSKU in ordenadoSKUs)
                                {
                                    //o	>> ... y luego la tabla PiezasArteSKU con IdentifPieza, TipoProd = “Primario” y IdentifSKU = SKU.IdentifSKU
                                    var productoPiezaArte = productosPiezaArte.Find(x => x.TipoProd.Trim().ToUpper() == "PRIMARIO" && x.IdentifSKU == ordenadoSKU.IdentifSKU);

                                    ordenadoSKU.Costo += (productoPiezaArte.Coeficiente.Value * costo);

                                    if (ordenadoSKU.Duracion > 0)
                                    {
                                        ordenadoSKU.CostoUni = (ordenadoSKU.Costo / ordenadoSKU.Duracion.Value);
                                    }
                                    else
                                    {
                                        ordenadoSKU.CostoUni = (ordenadoSKU.Costo);
                                    }


                                    //•	Si campo CostoProveedorVersion.IncluidoOP = “Si”
                                    if (costoProveedorVersion.IncluidoOP)
                                    {
                                        ordenadoSKU.CostoOp += (productoPiezaArte.Coeficiente.Value * costo);

                                        if (ordenadoSKU.Duracion > 0)
                                        {
                                            ordenadoSKU.CostoOpUni = (ordenadoSKU.CostoOp / ordenadoSKU.Duracion.Value);
                                        }
                                        else
                                        {
                                            ordenadoSKU.CostoOpUni = (ordenadoSKU.CostoOp);
                                        }

                                    }

                                    //o	Actualizar Tabla SKU
                                    daoSKU.Update(ordenadoSKU, ordenadoSKU.RecId, tran);
                                }
                            }

                            EstimadoCAB.Costo = costoAcum;

                            if (EstimadoCAB.DuracionTot > 0)
                            {
                                EstimadoCAB.CostoUni = EstimadoCAB.Costo / EstimadoCAB.DuracionTot;
                            }
                            else
                            {
                                EstimadoCAB.CostoUni = EstimadoCAB.Costo;
                            }

                            if (costoProveedorVersion.IncluidoOP)
                            {
                                EstimadoCAB.CostoOp = EstimadoCAB.Costo;

                                if (EstimadoCAB.DuracionTot != 0)
                                {
                                    EstimadoCAB.CostoOpUni = EstimadoCAB.CostoOp / EstimadoCAB.DuracionTot;
                                }
                                else
                                {
                                    EstimadoCAB.CostoOpUni = EstimadoCAB.CostoOp;
                                }

                            }

                            EstimadoCAB.VersionCosto = costoCab.Version.Value;
                            EstimadoCAB.VigDesde     = costoCab.VigDesde;
                            EstimadoCAB.VigHasta     = costoCab.VigHasta;
                            EstimadoCAB.FecCosto     = DateTime.Now;
                            EstimadoCAB.UsuCosto     = usuario;

                            dao.Update(EstimadoCAB, EstimadoCAB.RecId, tran);
                        }

                        dao.CommitTransaccion(tran);
                    }
                    catch (Exception ex)
                    {
                        dao.RollbackTransaccion(tran);
                        throw new Exception("CalcularCosto", ex);
                    }
                }
            }
        }

        public static List<EstimadoCabVersionDTO> GetVersiones(string pautaId)
        {
            return daoVer.ReadAll("PautaId = " + pautaId);
        }

        public static List<EstimadoDetVersionDTO> GetDetalles(string pautaId, int version)
        {
            return daoDetalleVer.ReadAll(string.Format("PautaId = {0} AND Version = {1}", pautaId, version));
        }

        public static List<EstimadoSKUVersionDTO> GetSKUs(string pautaId, int version)
        {
            return daoSKUVer.ReadAll(string.Format("PautaId = {0} AND Version = {1}", pautaId, version));
        }

        public static List<EstimadoCabDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        public static List<EstimadoSKUDTO> ReadAllSKUs(EstimadoCabDTO estimado)
        {
            return daoSKU.ReadAll(string.Format("PautaId = '{0}'", estimado.PautaId));
        }

        public static List<EstimadoDetDTO> ReadAllLineas(EstimadoCabDTO estimado)
        {
            return daoDetalle.ReadAll(string.Format("PautaId = '{0}'", estimado.PautaId));
        }

        internal static EstimadoCabDTO Read(string pautaId)
        {
            return dao.Read(string.Format("PautaId = '{0}'", pautaId));
        }
    }
}
