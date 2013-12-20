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
using System.Windows.Forms;
using System.Web.UI;

namespace PautasPublicidad.Business
{
    public static class Certificados
    {
        static CertificadoCabDAO dao         = DAOFactory.Get<CertificadoCabDAO>();
        static CertificadoDetDAO daoDetalle  = DAOFactory.Get<CertificadoDetDAO>();
        static CertificadoSKUDAO daoSKU      = DAOFactory.Get<CertificadoSKUDAO>();
        static CostosDAO daoCosto            = DAOFactory.Get<CostosDAO>();
        static EstimadoCabDAO daoEstimadoCab = DAOFactory.Get<EstimadoCabDAO>();
        static SetUpDAO daoSetUp             = DAOFactory.Get<SetUpDAO>();

        static public CertificadoCabDTO Buscar(string identifEspacio, int año, int mes, string identifOrigen)
        {
            CertificadoCabDTO cDummy = null;

            int anoMes = Convert.ToInt32(año.ToString() + mes.ToString("00"));

            var c = dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1} AND IdentifOrigen = '{2}'", identifEspacio, anoMes, identifOrigen));

            if (c != null)
            {
                return c;
            }
            else
            {
                //Lo busco sin origen, y construyo los registros con origen.
                var d = dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1}", identifEspacio, anoMes));

                if (identifOrigen != null && identifOrigen.Trim() != "")
                {
                    DialogResult dr = new DialogResult();

                    dr = MessageBox.Show("No existe un certificado para ese Origen. Desea crearlo?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

                    if (dr == DialogResult.Yes)
                    {
                        CopyToNewOrigen(Get(d.RecId), identifOrigen);
                        return dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1} AND IdentifOrigen = '{2}'", identifEspacio, anoMes, identifOrigen));
                    }
                    else
                    {
                        return cDummy;
                    }
                }
                return cDummy;
            }
        }

        static public string AnoMesCierreOrd()
        {
            return dao.AnoMesCierreOrd();
        }

        static public string BuscarCertificado(string Espacio, string Anio, string Mes, string pautaId, string identifOrigen)
        {
            string retVal = string.Empty;

            CertificadoCabDTO c = null;

            int anoMes = Convert.ToInt32(Anio + Mes.PadLeft(2,'0'));

            if (Espacio != string.Empty && Espacio != null)
            {
                c = dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1} AND IdentifOrigen = '{2}'", Espacio, anoMes, identifOrigen));

                if (c == null)
                {
                    var d = dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1}", Espacio, anoMes));

                    if (d == null)
                    {
                        retVal = "Certificado NULL";
                    }
                    else
                    {
                        retVal = "Origen NULL";
                    }
                }
                else
                {
                    retVal = "Certificado OK";
                }
            }
            else
            {
                c = dao.Read(string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", pautaId, identifOrigen));

                if (c == null)
                {
                    var d = dao.Read(string.Format("PautaId = '{0}' AND IdentifOrigen IS NULL", pautaId));

                    if (d == null)
                    {
                        retVal = "Certificado NULL";
                    }
                    else
                    {
                        retVal = "Origen NULL";
                    }
                }
                else
                {
                    retVal = "Certificado OK";
                }
            }

            return retVal;

        }

        static public CertificadoCabDTO Crear(string pautaId, string identifOrigen)
        {
            {
                var d = dao.Read(string.Format("PautaId = '{0}'", pautaId));
                CopiarANuevoOrigen(Get(d.RecId), identifOrigen);
                return dao.Read(string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", pautaId, identifOrigen));
            }
        }

        static public CertificadoCabDTO Buscar(string pautaId, string identifOrigen)
        {
            var c = dao.Read(string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", pautaId, identifOrigen));

            if (c != null)
            {
                return c;
            }
            else
            {
                var d = dao.Read(string.Format("PautaId = '{0}'", pautaId));

                if (identifOrigen != null && identifOrigen.Trim() != "" && d != null)
                {
                    CopyToNewOrigen(Get(d.RecId), identifOrigen);
                    return dao.Read(string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", pautaId, identifOrigen));
                }
                else
                {
                    return dao.Read(string.Format("PautaId = '{0}' AND IdentifOrigen IS NULL", pautaId));
                }
            }
        }

        static public void Create(CertificadoCabDTO certificado, List<CertificadoDetDTO> lineas)
        {
            CertificadoSKUDTO sku;

            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    certificado.PautaId = SaveNextPautaId(tran).ToString();
                    certificado         = dao.Create(certificado, tran);

                    foreach (CertificadoDetDTO linea in lineas)
                    {
                        linea.RecId     = 0;
                        linea.DatareaId = certificado.DatareaId;
                        linea.PautaId   = certificado.PautaId;

                        daoDetalle.Create(linea, tran);
                    }

                    var dtSKU = BuildAllSKU(lineas);
                    foreach (System.Data.DataRow dr in dtSKU.Rows)
                    {
                        sku = new CertificadoSKUDTO();

                        sku.RecId     = 0;
                        sku.DatareaId = certificado.DatareaId;
                        sku.PautaId   = certificado.PautaId;

                        if (dr["Duracion"] != DBNull.Value)
                            sku.Duracion = Convert.ToDecimal(dr["Duracion"]);
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
        
        
        
        static public void Create(Certificado certificado, SqlTransaction tran)
        {
            try
            {
                dao.Create(certificado.Cabecera, tran);

                foreach (var linea in certificado.Lineas)
                {
                    linea.RecId     = 0;
                    linea.DatareaId = certificado.Cabecera.DatareaId;
                    linea.PautaId   = certificado.Cabecera.PautaId;

                    daoDetalle.Create(linea, tran);
                }

                foreach (var sku in certificado.SKUs)
                {
                    sku.RecId     = 0;
                    sku.DatareaId = certificado.Cabecera.DatareaId;
                    sku.PautaId   = certificado.Cabecera.PautaId;

                    daoSKU.Create(sku, tran);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }


        static public void CopiarANuevoOrigen(Certificado certificado, string identifOrigen)
        {
            if (identifOrigen.Trim() != "" && (certificado.Cabecera.IdentifOrigen == null || certificado.Cabecera.IdentifOrigen.Trim() == ""))
            {
                using (SqlTransaction tran = dao.IniciarTransaccion())
                {
                    try
                    {
                        //Genero nuevos registros de certificado, con el origen seleccionado.
                        certificado.Cabecera.RecId         = 0;
                        certificado.Cabecera.IdentifOrigen = identifOrigen;
                        certificado.Lineas.ForEach((o) => o.IdentifOrigen = identifOrigen);
                        certificado.SKUs.ForEach((s) => s.IdentifOrigen = identifOrigen);

                        Create(certificado, tran);

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



        internal static void CopyToNewOrigen(Certificado certificado, string identifOrigen)
        {
            if (identifOrigen.Trim() != "" && (certificado.Cabecera.IdentifOrigen == null || certificado.Cabecera.IdentifOrigen.Trim() == ""))
            {
                using (SqlTransaction tran = dao.IniciarTransaccion())
                {
                    try
                    {
                        //Genero nuevos registros de certificado, con el origen seleccionado.
                        certificado.Cabecera.RecId         = 0;
                        certificado.Cabecera.IdentifOrigen = identifOrigen;
                        certificado.Lineas.ForEach((o) => o.IdentifOrigen = identifOrigen);
                        certificado.SKUs.ForEach((s) => s.IdentifOrigen = identifOrigen);

                        Create(certificado, tran);

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

        internal static void CierreEstimado(Certificado certificado, EstimadoCabDTO estimadoCab)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    Create(certificado, tran);
                    
                    daoEstimadoCab.Update(estimadoCab, estimadoCab.RecId, tran);

                    dao.CommitTransaccion(tran);
                }
                catch (Exception)
                {
                    dao.RollbackTransaccion(tran);
                    throw;
                }
            }
        }


        static public CostosDTO FindCosto(string identifEspacio, int año, int mes)
        {
            DateTime desde = new DateTime(año, mes, 1);
            DateTime hasta = new DateTime(año, mes, DateTime.DaysInMonth(año, mes));

            return daoCosto.Read(string.Format("IdentifEspacio = '{0}' AND VigDesde <= '{1}' AND VigHasta >= '{2}' AND IsNull(Confirmado,'') <> ''", identifEspacio, hasta.ToString("yyyyMMdd"), desde.ToString("yyyyMMdd")));
        }

        public static Certificado Get(int recId)
        {
            var certificado = new Certificado();

            certificado.Cabecera = Read(recId);

            if (certificado.Cabecera.IdentifOrigen != null)
            {
                certificado.Lineas = daoDetalle.ReadAll(string.Format("PautaId='{0}' AND IdentifOrigen='{1}'", certificado.Cabecera.PautaId, certificado.Cabecera.IdentifOrigen));
                certificado.SKUs = daoSKU.ReadAll(string.Format("PautaId='{0}' AND IdentifOrigen='{1}'", certificado.Cabecera.PautaId, certificado.Cabecera.IdentifOrigen));
            }
            else
            {
                certificado.Lineas = daoDetalle.ReadAll(string.Format("PautaId='{0}' AND IdentifOrigen IS NULL", certificado.Cabecera.PautaId));
                certificado.SKUs = daoSKU.ReadAll(string.Format("PautaId='{0}' AND IdentifOrigen IS NULL", certificado.Cabecera.PautaId));
            }

            return certificado;
        }

        public static CertificadoCabDTO Read(int recId)
        {
            return dao.Read(recId);
        }

        static public CertificadoCabDTO Read(string PautaId, string IdentifOrigen)
        {
            return dao.Read("PAUTAID = '" + PautaId + "' AND IDENTIFORIGEN = '" + IdentifOrigen + "'");

        }
        static public CertificadoCabDTO Read(string identifEspacio, int año, int mes, string identifOrigen)
        {
            DateTime desde = new DateTime(año, mes, 1);
            DateTime hasta = new DateTime(año, mes, DateTime.DaysInMonth(año, mes));

            return dao.Read(string.Format("IdentifEspacio = '{0}' AND AnoMes = {1} AND IdentifOrigen = '{2}'", identifEspacio, Convert.ToInt32(año.ToString() + mes.ToString("00")),identifOrigen));
        }

        static public List<CertificadoCabDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        public static List<CertificadoDetDTO> ReadAllLineas(CertificadoCabDTO certificado)
        {

            if (certificado.IdentifOrigen != "NULL")
            {
                List<CertificadoDetDTO> mycert = daoDetalle.ReadAll(string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", certificado.PautaId,certificado.IdentifOrigen));

                return mycert;
            }
            return daoDetalle.ReadAll(string.Format("PautaId = '{0}' AND IdentifOrigen != 'NULL'", certificado.PautaId));
        }

        static public DataView VistaCertificados()
        {
            return dao.VistaCertificados();
        }

        public static DataTable BuildAllSKU(List<CertificadoDetDTO> lineas)
        {
            DataTable dt = daoSKU.GetSKUByPiezasArte(lineas);

            foreach (DataRow dr in dt.Rows)
                dr["CantSalidas"] = lineas.FindAll(x => x.IdentifAviso == (string)dr["IdentifAviso"]).Count;

            return dt;
        }

        static public decimal SaveNextPautaId(SqlTransaction tran)
        {
            try
            {
                SetUpDTO s = daoSetUp.Read("");
                s.NumPauta += 1;
                daoSetUp.Update(s, s.RecId, tran);
                return s.NumPauta;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void Update(CertificadoCabDTO certificado, string sWhere)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(certificado, sWhere, tran);

                    dao.CommitTransaccion(tran);

                }
                catch
                {
                    dao.RollbackTransaccion(tran);
                }

            }

        }

        public static void Update(CertificadoCabDTO certificado, List<CertificadoDetDTO> lineas)
        {
            CertificadoSKUDTO sku;

            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(certificado, certificado.RecId, tran);

                    //Elimino toda las lineas del ordenado y las re-creo.
                    daoDetalle.Delete(
                        string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'",
                        certificado.PautaId,
                        certificado.IdentifOrigen),
                        tran);

                    foreach (CertificadoDetDTO linea in lineas)
                    {
                        linea.RecId = 0;
                        linea.DatareaId = certificado.DatareaId;
                        linea.PautaId = certificado.PautaId;
                        linea.IdentifOrigen = certificado.IdentifOrigen;
                        
                        daoDetalle.Create(linea, tran);
                    }

                    //Elimino toda las lineas del ordenado y las re-creo.
                    daoSKU.Delete(
                        string.Format("PautaId = '{0}'",
                        certificado.PautaId),
                        tran);

                    var dtSKU = BuildAllSKU(lineas);
                    foreach (System.Data.DataRow dr in dtSKU.Rows)
                    {
                        sku = new CertificadoSKUDTO();

                        sku.RecId = 0;
                        sku.DatareaId = certificado.DatareaId;
                        sku.PautaId = certificado.PautaId;
                        sku.IdentifOrigen = certificado.IdentifOrigen;

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

        public static void CalcularCosto(CertificadoCabDTO certificadoCAB, CostosDTO costoCab, List<CertificadoDetDTO> lineas, string usuario)
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
                        certificadoCAB = dao.Read(string.Format("PautaId='{0}'", certificadoCAB.PautaId));

                        if (costoCab == null)
                        {
                            throw new Exception("No hay costos confirmados para realizar el costeo de la Pauta: " + certificadoCAB.PautaId.ToString());
                        }

                        //Inicializo en '0' los registros de trabajo.
                        dao.MoverCeros(certificadoCAB.PautaId, tran);
                        daoDetalle.MoverCeros(certificadoCAB.PautaId, tran);
                        daoSKU.MoverCeros(certificadoCAB.PautaId, tran);

                        //•	Seleccionar la Tabla CostoProveedorVersion 
                        var costosProveedorVersion = Costos.ReadAllProveedorVersiones(certificadoCAB.IdentifEspacio, costoCab.VigDesde, costoCab.VigHasta, costoCab.Version.Value);

                        //Por cada registro seleccionado:
                        foreach (var costoProveedorVersion in costosProveedorVersion)
                        {
                            //•	Calcular Tipo de cambio
                            var tipoCambioValor = Monedas.GetTipoCambioValor(costoProveedorVersion.IdentifMon);
                            var costoCur = costoProveedorVersion.Costo * tipoCambioValor;

                            //•	Calcular GrossingUp
                            var costoGros = costoCur * costoProveedorVersion.GrossingUp;

                            //•	Actualizar Tabla DET
                            var certificadoDetalles = daoDetalle.ReadAll(string.Format("PautaId='{0}'", certificadoCAB.PautaId));
                            foreach (var certificadoDET in certificadoDetalles)
                            {
                                if (costoProveedorVersion.TipoCosto == "FIJO_MENSUAL")
                                {
                                    if (certificadoCAB.DuracionTot > 0)
                                    {
                                        if (certificadoDET.Duracion != null)
                                        {
                                            costo = (costoGros / certificadoCAB.DuracionTot) * certificadoDET.Duracion.Value;
                                        }
                                        else
                                        {
                                            decimal? divisor = 0;

                                            costo = (costoGros / certificadoCAB.DuracionTot) * divisor.Value;
                                        }

                                    }
                                    else
                                    {
                                        costo = costoGros;
                                    }

                                }
                                else if (costoProveedorVersion.TipoCosto == "SEGUNDO_FIJO")
                                {
                                    costo = costoGros * certificadoDET.Duracion.Value;
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

                                certificadoDET.Costo += costo;

                                if (certificadoDET.Duracion != null)
                                {

                                    if (certificadoDET.Duracion > 0)
                                    {
                                        certificadoDET.CostoUni = certificadoDET.Costo / certificadoDET.Duracion.Value;
                                    }
                                    else
                                    {
                                        certificadoDET.CostoUni = certificadoDET.Costo;
                                    }

                                }
                                else
                                {
                                    certificadoDET.CostoUni = certificadoDET.Costo;
                                }
                                

                                if (costoProveedorVersion.IncluidoOP)
                                {
                                    certificadoDET.CostoOp += costo;

                                    if (certificadoDET.Duracion != null)
                                    {
                                        certificadoDET.CostoOpUni = certificadoDET.CostoOp / certificadoDET.Duracion.Value;
                                    }
                                    else
                                    {
                                        certificadoDET.CostoOpUni = certificadoDET.CostoOp;
                                    }

                                }

                                costoAcum += costo;

                                //o	Actualizar la Tabla DET
                                daoDetalle.Update(certificadoDET, certificadoDET.RecId, tran);
                            }

                            //•	Actualizar Tabla SKU
                            foreach (var certificadoDET in certificadoDetalles)
                            {
                                //o	Seleccionar la Tabla SKU con SKU.PautaID = PautaId enviado y SKU.IdentifAviso = DET.IdentifAviso
                                var ordenadoSKUs = daoSKU.ReadAll(string.Format("PautaId='{0}' AND IdentifAviso='{1}'", certificadoCAB.PautaId, certificadoDET.IdentifAviso));

                                //o	Seleccionar la Tabla Avisos con IdentifAviso... >>
                                var aviso = Avisos.Read(string.Format("IdentifAviso='{0}'", certificadoDET.IdentifAviso));
                                var productosPiezaArte = PiezasArte.ReadAllProductos(aviso.IdentifPieza);
                                foreach (var ordenadoSKU in ordenadoSKUs)
                                {
                                    //o	>> ... y luego la tabla PiezasArteSKU con IdentifPieza, TipoProd = “Primario” y IdentifSKU = SKU.IdentifSKU
                                    var productoPiezaArte = productosPiezaArte.Find(x => x.TipoProd.Trim().ToUpper() == "PRIMARIO" && x.IdentifSKU == ordenadoSKU.IdentifSKU);

                                    ordenadoSKU.Costo += (productoPiezaArte.Coeficiente.Value * costo);

                                    if (ordenadoSKU.Duracion != null)
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

                                        if (ordenadoSKU.Duracion != null)
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

                            certificadoCAB.Costo = costoAcum;

                            if (certificadoCAB.DuracionTot > 0)
                            {
                                certificadoCAB.CostoUni = certificadoCAB.Costo / certificadoCAB.DuracionTot;
                            }
                            else
                            {
                                certificadoCAB.CostoUni = certificadoCAB.Costo;
                            }


                            if (costoProveedorVersion.IncluidoOP)
                            {
                                certificadoCAB.CostoOp = certificadoCAB.Costo;

                                if (certificadoCAB.DuracionTot > 0)
                                {
                                    certificadoCAB.CostoOpUni = certificadoCAB.CostoOp / certificadoCAB.DuracionTot;
                                }
                                else
                                {
                                    certificadoCAB.CostoOpUni = certificadoCAB.CostoOp;
                                }

                            }

                            certificadoCAB.VersionCosto = costoCab.Version.Value;
                            certificadoCAB.VigDesde = costoCab.VigDesde;
                            certificadoCAB.VigHasta = costoCab.VigHasta;
                            certificadoCAB.FecCosto = DateTime.Now;
                            certificadoCAB.UsuCosto = usuario;

                            dao.Update(certificadoCAB, certificadoCAB.RecId, tran);
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

        public static List<CertificadoCabDTO> GetOrigenes(string pautaId)
        {
            return dao.ReadAll("PautaId = " + pautaId);
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

        static public void Delete(CertificadoCabDTO certificado) //(int id)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {   //Elimino toda las lineas del certificado detalle
                    daoDetalle.Delete( string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", certificado.PautaId,certificado.IdentifOrigen), tran);
                    //Elimino toda las lineas del certificado sku
                    daoSKU.Delete( string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", certificado.PautaId,certificado.IdentifOrigen), tran);
                    //elimino certificado.
                    dao.Delete(string.Format("PautaId = '{0}' AND IdentifOrigen = '{1}'", certificado.PautaId, certificado.IdentifOrigen), tran);

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
        
        
        
        public static List<CertificadoDetDTO> GetDetalles(string pautaId, string identifOrigen)
        {
            return daoDetalle.ReadAll(string.Format("PautaId = {0} AND IdentifOrigen = '{1}'", pautaId, identifOrigen));
        }

        public static List<CertificadoSKUDTO> GetSKUs(string pautaId, string identifOrigen)
        {
            return daoSKU.ReadAll(string.Format("PautaId = {0} AND IdentifOrigen = '{1}'", pautaId, identifOrigen));
        }
    }
}
