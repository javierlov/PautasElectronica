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
using System.Reflection;

namespace PautasPublicidad.Business
{
    public static class Ordenados
    {
        static OrdenadoCabDAO dao = DAOFactory.Get<OrdenadoCabDAO>();
        static OrdenadoDetDAO daoDetalle = DAOFactory.Get<OrdenadoDetDAO>();
        static OrdenadoSKUDAO daoSKU = DAOFactory.Get<OrdenadoSKUDAO>();
        static CostosDAO daoCosto = DAOFactory.Get<CostosDAO>();
        static SetUpDAO daoSetUp = DAOFactory.Get<SetUpDAO>();

        public enum eSemanaMes
        {
            SEMANA,
            MES
        }

        static public List<OrdenadoCabDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public OrdenadoCabDTO Read(string pautaId)
        {
            return dao.Read(string.Format("PautaId = '{0}'", pautaId));
        }

        static public OrdenadoCabDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public DataView VistaOrdenados()
        {
            return dao.VistaOrdenados();
        }

        static public DataView VistaCostosConfirmados()
        {
            return dao.VistaCostosConfirmados();
        }

        static public string AnoMesCierreOrd()
        {
            return dao.AnoMesCierreOrd();
        }

        static public OrdenadoCabDTO Read(string identifEspacio, int año, int mes)
        {
            DateTime desde = new DateTime(año, mes, 1);
            DateTime hasta = new DateTime(año, mes, DateTime.DaysInMonth(año, mes));

            return dao.Read(
                string.Format("IdentifEspacio = '{0}' AND AnoMes = {1}",
                    identifEspacio, Convert.ToInt32(año.ToString() + mes.ToString("00"))));
        }

        /// <summary>
        /// Encuentro el registro de costo, tal que el 'año/mes' buscado que mando por parámetro,
        /// caiga dentro del período comprendido por la VigDesde - VigHasta del costo.
        /// </summary>
        /// <param name="identifEspacio"></param>
        /// <param name="año"></param>
        /// <param name="mes"></param>
        /// <returns></returns>
        static public CostosDTO FindCosto(string identifEspacio, int año, int mes)
        {
            DateTime desde = new DateTime(año, mes, 1);
            DateTime hasta = new DateTime(año, mes, DateTime.DaysInMonth(año, mes));

            string sDesde = desde.Year.ToString() + "-" + desde.Month.ToString("00") + "-" + desde.Day.ToString("00");
            string sHasta = hasta.Year.ToString() + "-" + hasta.Month.ToString("00") + "-" + hasta.Day.ToString("00");

            return daoCosto.Read(string.Format("IdentifEspacio = '{0}' AND '{1}' >= VigDesde AND '{2}' <= VigHasta AND IsNull(Confirmado,'') <> ''", identifEspacio, sDesde, sHasta));
        }

        static public decimal SaveNextPautaId(SqlTransaction tran)
        {
                try
                {
                    SetUpDTO s = daoSetUp.Read("");
                    s.NumPauta += 1;
                    daoSetUp.Update(s,s.RecId, tran);
                    return s.NumPauta;
                }
                catch (Exception)
                {
                    throw;
                }
        }

        static public int GetAñoMesCierreOrd()
        {
            try
            {
                SetUpDTO s = daoSetUp.Read("");
                return Convert.ToInt32(s.AnoMesCierreOrd);
            }
            catch (Exception)
            {
                throw;
            }
        }

        static public void SetAñoMesCierreEst(int añoMesEst, SqlTransaction tran)
        {
            using ( tran = dao.IniciarTransaccion())
            {
                try
                {
                    SetUpDTO s        = daoSetUp.Read("");
                    s.AnoMesCierreEst = añoMesEst;
                    daoSetUp.Update(s, s.RecId, tran);

                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        static public void SetAñoMesCierreOrd(int añoMesOrd, SqlTransaction tran)
        {
            using (tran = dao.IniciarTransaccion())
            {
                try
                {
                    SetUpDTO s        = daoSetUp.Read("");
                    s.AnoMesCierreOrd = añoMesOrd;
                    daoSetUp.Update(s, s.RecId, tran);

                    dao.CommitTransaccion(tran);

                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        static public int GetAñoMesCierreEst()
        {
            try
            {
                SetUpDTO s = daoSetUp.Read("");
                return Convert.ToInt32(s.AnoMesCierreEst);
            }
            catch (Exception)
            {
                throw;
            }
        }

        static public decimal GetNextPautaId()
        {
            try
            {
                SetUpDTO s = daoSetUp.Read("");
                s.NumPauta += 1;
                return s.NumPauta;
            }
            catch (Exception)
            {
                throw;
            }
        }

        static public void Create(OrdenadoCabDTO ordenado, List<OrdenadoDetDTO> lineas)
        {
            OrdenadoSKUDTO sku;

            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    ordenado.PautaId = SaveNextPautaId(tran).ToString();
                    ordenado         = dao.Create(ordenado, tran);

                    foreach (OrdenadoDetDTO linea in lineas)
                    {
                        linea.RecId     = 0;
                        linea.DatareaId = ordenado.DatareaId;
                        linea.PautaId   = ordenado.PautaId;

                        daoDetalle.Create(linea, tran);
                    }

                    var dtSKU = BuildAllSKU(lineas);
                    foreach (System.Data.DataRow dr in dtSKU.Rows)
                    {
                        sku = new OrdenadoSKUDTO();
                        
                        sku.RecId     = 0;
                        sku.DatareaId = ordenado.DatareaId;
                        sku.PautaId   = ordenado.PautaId;

                        if (dr["Duracion"] != DBNull.Value)
                            sku.Duracion    = Convert.ToDecimal(dr["Duracion"]);
                        else
                            sku.Duracion = null;

                        sku.CantSalidas  = Convert.ToDecimal(dr["CantSalidas"]);
                        sku.IdentifAviso = Convert.ToString(dr["IdentifAviso"]);
                        sku.IdentifSKU   = Convert.ToString(dr["IdentifSKU"]);
                        
                        daoSKU.Create(sku, tran);
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

        public static void Update(OrdenadoCabDTO ordenado, List<OrdenadoDetDTO> lineas)
        {
            OrdenadoSKUDTO sku;

            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(ordenado, ordenado.RecId, tran);

                    //Elimino toda las lineas del ordenado y las re-creo.
                    daoDetalle.Delete(
                        string.Format("PautaId = '{0}'",
                        ordenado.PautaId),
                        tran);

                    foreach (OrdenadoDetDTO linea in lineas)
                    {
                        linea.RecId = 0;
                        linea.DatareaId = ordenado.DatareaId;
                        linea.PautaId  = ordenado.PautaId;

                        daoDetalle.Create(linea, tran);
                    }

                    //Elimino toda las lineas del ordenado y las re-creo.
                    daoSKU.Delete(
                        string.Format("PautaId = '{0}'",
                        ordenado.PautaId),
                        tran);

                    var dtSKU = BuildAllSKU(lineas);
                    foreach (System.Data.DataRow dr in dtSKU.Rows)
                    {
                        sku = new OrdenadoSKUDTO();

                        sku.RecId = 0;
                        sku.DatareaId = ordenado.DatareaId;
                        sku.PautaId = ordenado.PautaId;

                        if (dr["Duracion"] != DBNull.Value)
                            sku.Duracion = Convert.ToDecimal(dr["Duracion"]);
                        else
                            sku.Duracion = null;

                        sku.CantSalidas = Convert.ToDecimal(dr["CantSalidas"]);
                        sku.IdentifAviso = Convert.ToString(dr["IdentifAviso"]);
                        sku.IdentifSKU = Convert.ToString(dr["IdentifSKU"]);

                        daoSKU.Create(sku, tran);
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

        static public void Delete(OrdenadoCabDTO ordenado) //(int id)
        {   
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {   //Elimino toda las lineas del ordenado detalle
                    daoDetalle.Delete(
                        string.Format("PautaId = '{0}'",
                        ordenado.PautaId),
                        tran);
                    //Elimino toda las lineas del ordenado sku
                    daoSKU.Delete(
                        string.Format("PautaId = '{0}'",
                        ordenado.PautaId),
                        tran);
                    //eliminio ordenado.
                    dao.Delete(ordenado.RecId, tran);

                    dao.CommitTransaccion(tran);
                   
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }

        private static bool ExisteEnArray(string valor, string[] valores)
        {
            foreach (string s in valores)
            {
                if (s.Trim().ToUpper() == valor.Trim().ToUpper())
                    return true;
            }
            return false;
        }

        public static List<DateTime> GetDatesByDayNames(int year, int month, List<string> dayNamesList)
        {
            string dayName;
            DateTime date;
            List<DateTime> datesList = new List<DateTime>();

            for (int day = 1; day <= DateTime.DaysInMonth(year, month); day++)
            {
                date = new DateTime(year, month, day);
                dayName = date.ToString("dddd", new CultureInfo("es-ES")).ToUpper().Trim();

                if (ExisteEnArray(dayName.Replace('é', 'e').Replace('á', 'a').Replace('É', 'E').Replace('Á', 'A'), dayNamesList.ToArray()))
                    datesList.Add(date);
            }
            return datesList;
        }

        public static List<DateTime> GetDatesByDayNumbers(int year, int month, List<string> dayNumbersList)
        {
            int day;
            List<DateTime> datesList = new List<DateTime>();

            foreach (string dayNumber in dayNumbersList)
            {
                if (int.TryParse(dayNumber, out day))
                    datesList.Add(new DateTime(year, month, day));
            }
            return datesList;
        }

        public static List<OrdenadoDetDTO> ReadAllLineas(OrdenadoCabDTO ordenado)
        {
            return daoDetalle.ReadAll(string.Format("PautaId = '{0}'", ordenado.PautaId));
        }

        public static List<OrdenadoSKUDTO> ReadAllSKUs(OrdenadoCabDTO ordenado)
        {
            return daoSKU.ReadAll(string.Format("PautaId = '{0}'", ordenado.PautaId));
        }

        public static DataTable BuildAllSKU(List<OrdenadoDetDTO> lineas)
        {
            DataTable dt = daoSKU.GetSKUByPiezasArte(lineas);

            foreach (DataRow dr in dt.Rows)
                dr["CantSalidas"] = lineas.FindAll(x => x.IdentifAviso == (string)dr["IdentifAviso"]).Count;

            return dt;
        }

        public static void CalcularCosto(OrdenadoCabDTO ordenadoCAB, CostosDTO costoCab, List<OrdenadoDetDTO> lineas, string usuario)
        {
            decimal costo = 0;
            decimal costoAcum = 0;
            
            //Lockeo la seccion para que otro thread no me descontrole los calculos.
            lock (typeof(Ordenado))
            {
                using (SqlTransaction tran = dao.IniciarTransaccion())
                {
                    try
                    {
                        //La recargo para asegurarme de tener la version mas reciente...
                        ordenadoCAB = dao.Read(string.Format("PautaId='{0}'", ordenadoCAB.PautaId));

                        if(ordenadoCAB == null)
                            throw new Exception("");

                        if (costoCab == null)
                            throw new Exception("No hay costos confirmados para realizar el costeo de la Pauta: " + ordenadoCAB.PautaId.ToString());

                        if (ordenadoCAB.UsuCierre != "" && ordenadoCAB.UsuCierre != null)
                            throw new Exception("El ordenado se encuentra cerrado");

                        //Inicializo en '0' los registros de trabajo.
                        dao.MoverCeros(ordenadoCAB.PautaId, tran);
                        daoDetalle.MoverCeros(ordenadoCAB.PautaId, tran);
                        daoSKU.MoverCeros(ordenadoCAB.PautaId, tran);

                        //•	Seleccionar la Tabla CostoProveedorVersion 
                        var costosProveedorVersion = Costos.ReadAllProveedorVersiones(ordenadoCAB.IdentifEspacio, costoCab.VigDesde, costoCab.VigHasta, costoCab.Version.Value);
                    
                        //Por cada registro seleccionado:
                        foreach (var costoProveedorVersion in costosProveedorVersion)
                        {
                            var tipoCambioValor = Monedas.GetTipoCambioValor(costoProveedorVersion.IdentifMon); //•	Calcular Tipo de cambio y moneda
                            var costoCur        = costoProveedorVersion.Costo * tipoCambioValor;                //•	Calcular Tipo de cambio y moneda
                            var costoGros       = costoCur * costoProveedorVersion.GrossingUp;                  //•	Calcular Impuestos

                            //•	Actualizar Tabla DET
                            var ordenadoDetalles = daoDetalle.ReadAll(string.Format("PautaId='{0}'", ordenadoCAB.PautaId));

                            foreach (var ordenadoDET in ordenadoDetalles)
                            {
                                if (costoProveedorVersion.TipoCosto == "FIJO_MENSUAL")
                                    if (ordenadoCAB.DuracionTot > 0)
                                        if (ordenadoDET.Duracion.Value > 0)
                                            costo = (costoGros / ordenadoCAB.DuracionTot) * ordenadoDET.Duracion.Value;
                                        else
                                            costo = 0;
                                    else
                                        costo = costoGros;
                                else if (costoProveedorVersion.TipoCosto == "SEGUNDO_FIJO")
                                    costo = costoGros * ordenadoDET.Duracion.Value;
                                else if (costoProveedorVersion.TipoCosto == "SALIDA" || costoProveedorVersion.TipoCosto == "UNIDAD_PAUTADA")
                                    costo = costoGros;
                                else
                                    throw new Exception("Proveedor.TipoCosto Desconocido.");

                                ordenadoDET.Costo += costo;

                                if (ordenadoDET.Duracion ==null)
                                    ordenadoDET.CostoUni = ordenadoDET.Costo / 1;
                                else if (ordenadoDET.Duracion == 0)
                                        ordenadoDET.CostoUni = ordenadoDET.Costo / 1;
                                     else if (ordenadoDET.Duracion!=0)
                                        ordenadoDET.CostoUni = ordenadoDET.Costo / ordenadoDET.Duracion.Value;
                                
                                if (costoProveedorVersion.IncluidoOP)
                                {
                                    ordenadoDET.CostoOp += costo;
                                    if(ordenadoDET.Duracion ==null)
                                        ordenadoDET.CostoOpUni = ordenadoDET.CostoOp / 1; ///fek hay q validar que se pueda usar valor 1.
                                    else if (ordenadoDET.Duracion == 0)
                                            ordenadoDET.CostoOpUni = ordenadoDET.CostoOp / 1; ///fek hay q validar que se pueda usar valor 1.
                                         else if (ordenadoDET.Duracion!=0)
                                            ordenadoDET.CostoOpUni = ordenadoDET.CostoOp / ordenadoDET.Duracion.Value;
                                }

                                costoAcum += costo;

                                //o	Actualizar la Tabla DET
                                daoDetalle.Update(ordenadoDET, ordenadoDET.RecId, tran);
                            }

                            //•	Actualizar Tabla SKU
                            foreach (var ordenadoDET in ordenadoDetalles)
                            {
                                //o	Seleccionar la Tabla SKU con SKU.PautaID = PautaId enviado y SKU.IdentifAviso = DET.IdentifAviso
                                var ordenadoSKUs = daoSKU.ReadAll(string.Format("PautaId='{0}' AND IdentifAviso='{1}'", ordenadoCAB.PautaId, ordenadoDET.IdentifAviso));

                                //o	Seleccionar la Tabla Avisos con IdentifAviso... >>
                                var aviso = Avisos.Read(string.Format("IdentifAviso='{0}'", ordenadoDET.IdentifAviso));
                                var productosPiezaArte = PiezasArte.ReadAllProductos(aviso.IdentifPieza);
                                
                              
                                foreach (var ordenadoSKU in ordenadoSKUs)
                                {
                                    //o	>> ... y luego la tabla PiezasArteSKU con IdentifPieza, TipoProd = “Primario” y IdentifSKU = SKU.IdentifSKU
                                    var productoPiezaArte = productosPiezaArte.Find(x => x.TipoProd.Trim().ToUpper() == "PRIMARIO" && x.IdentifSKU == ordenadoSKU.IdentifSKU);

                                    ordenadoSKU.Costo += (productoPiezaArte.Coeficiente.Value * costo);
                                    if (ordenadoSKU.Duracion == null)
                                        ordenadoSKU.CostoUni = (ordenadoSKU.Costo / 1); 
                                    else if(ordenadoSKU.Duracion == 0)
                                            ordenadoSKU.CostoUni = (ordenadoSKU.Costo / 1); 
                                         else if (ordenadoSKU.Duracion!=0)
                                            ordenadoSKU.CostoUni = (ordenadoSKU.Costo / ordenadoSKU.Duracion.Value);
                                            

                                    //•	Si campo CostoProveedorVersion.IncluidoOP = “Si”
                                    if (costoProveedorVersion.IncluidoOP)
                                    {
                                        ordenadoSKU.CostoOp += (productoPiezaArte.Coeficiente.Value * costo);
                                        if (ordenadoSKU.Duracion == null)
                                           ordenadoSKU.CostoOpUni = (ordenadoSKU.CostoOp / 1);
                                        else if( ordenadoSKU.Duracion ==0)
                                                ordenadoSKU.CostoOpUni = (ordenadoSKU.CostoOp / 1);   
                                             else if (ordenadoSKU.Duracion!= 0)
                                                ordenadoSKU.CostoOpUni = (ordenadoSKU.CostoOp / ordenadoSKU.Duracion.Value);
                                    }

                                    //o	Actualizar Tabla SKU
                                    daoSKU.Update(ordenadoSKU, ordenadoSKU.RecId, tran);
                                }
                            }

                            ordenadoCAB.Costo = costoAcum;
                            if (ordenadoCAB.DuracionTot == 0)
                                ordenadoCAB.CostoUni = ordenadoCAB.Costo / 1;     ///fek hay q validar que se pueda usar valor 1.
                            else if (ordenadoCAB.DuracionTot == 0)
                                    ordenadoCAB.CostoUni = ordenadoCAB.Costo / 1;     ///fek hay q validar que se pueda usar valor 1.
                                 else if (ordenadoCAB.DuracionTot != 0)
                                    ordenadoCAB.CostoUni = ordenadoCAB.Costo / ordenadoCAB.DuracionTot;
                                
                               
                            if (costoProveedorVersion.IncluidoOP)
                            {
                                ordenadoCAB.CostoOp = ordenadoCAB.Costo;
                                if (ordenadoCAB.DuracionTot == 0)
                                    ordenadoCAB.CostoOpUni = ordenadoCAB.CostoOp / 1; ///fek hay q validar que se pueda usar valor 1.
                                else if (ordenadoCAB.DuracionTot == 0)
                                        ordenadoCAB.CostoOpUni = ordenadoCAB.CostoOp / 1;
                                     else if (ordenadoCAB.DuracionTot!=0)
                                        ordenadoCAB.CostoOpUni = ordenadoCAB.CostoOp / ordenadoCAB.DuracionTot;
                            }

                            ordenadoCAB.VersionCosto = costoCab.Version.Value;
                            ordenadoCAB.VigDesde     = costoCab.VigDesde;
                            ordenadoCAB.VigHasta     = costoCab.VigHasta;
                            ordenadoCAB.FecCosto     = DateTime.Now;
                            ordenadoCAB.UsuCosto     = usuario;

                            dao.Update(ordenadoCAB, ordenadoCAB.RecId, tran);
                        }

                        dao.CommitTransaccion(tran);
                    }
                    catch (Exception ex)
                    {
                        dao.RollbackTransaccion(tran);
                        throw new Exception("Error al Calcular Costo", ex);
                    }
                }
            }
        }
    }
}
