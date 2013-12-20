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
    public static class Costos
    {
        static CostosDAO dao                     = DAOFactory.Get<CostosDAO>();
        static CostosFrecuenciaDAO daoFrecuencia = DAOFactory.Get<CostosFrecuenciaDAO>();
        static CostosProveedorDAO daoProveedor   = DAOFactory.Get<CostosProveedorDAO>();

        static CostoVersionDAO daoVer                      = DAOFactory.Get<CostoVersionDAO>();
        static CostosFrecuenciaVersionDAO daoFrecuenciaVer = DAOFactory.Get<CostosFrecuenciaVersionDAO>();
        static CostosProveedorVersionDAO daoProveedorVer   = DAOFactory.Get<CostosProveedorVersionDAO>();

        static public List<CostosDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public CostosDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public List<CostoVersionDTO> ReadAllVersiones(CostosDTO costo)
        {
            return daoVer.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}'", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd")));
        }

        static public List<CostosProveedorVersionDTO> ReadAllProveedorVersiones(string sWhere)
        {
            return daoProveedorVer.ReadAll(sWhere);
        }

        static public List<CostosFrecuenciaVersionDTO> ReadAllFrecuenciaVersiones(string sWhere)
        {
            return daoFrecuenciaVer.ReadAll(sWhere);
        }

        static public void Create(CostosDTO costo, List<CostosFrecuenciaDTO> frecuenciaList, List<CostosProveedorDTO> proveedorList)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    costo = dao.Create(costo, tran);

                    foreach (CostosFrecuenciaDTO frecuencia in frecuenciaList)
                    {
                        frecuencia.RecId          = 0;
                        frecuencia.DatareaId      = costo.DatareaId;
                        frecuencia.IdentifEspacio = costo.IdentifEspacio;
                        frecuencia.VigDesde       = costo.VigDesde;
                        frecuencia.VigHasta       = costo.VigHasta;

                        daoFrecuencia.Create(frecuencia, tran);
                    }

                    foreach (CostosProveedorDTO proveedor in proveedorList)
                    {
                        proveedor.RecId          = 0;
                        proveedor.DatareaId      = costo.DatareaId;
                        proveedor.IdentifEspacio = costo.IdentifEspacio;
                        proveedor.VigDesde       = costo.VigDesde;
                        proveedor.VigHasta       = costo.VigHasta;

                        daoProveedor.Create(proveedor, tran);

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

        public static void Update(CostosDTO costo, List<CostosFrecuenciaDTO> frecuenciaList, List<CostosProveedorDTO> proveedorList)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    dao.Update(costo, costo.RecId, tran);

                    //Elimino todos los frecuencia y proveedores y los re-creo.
                    daoFrecuencia.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);

                    daoProveedor.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);

                    foreach (CostosFrecuenciaDTO frecuencia in frecuenciaList)
                    {
                        frecuencia.RecId          = 0                   ;
                        frecuencia.DatareaId      = costo.DatareaId     ;
                        frecuencia.IdentifEspacio = costo.IdentifEspacio;
                        frecuencia.VigDesde       = costo.VigDesde      ;
                        frecuencia.VigHasta       = costo.VigHasta      ;

                        daoFrecuencia.Create(frecuencia, tran);
                    }

                    foreach (CostosProveedorDTO proveedor in proveedorList)
                    {
                        proveedor.RecId          = 0                    ;
                        proveedor.DatareaId      = costo.DatareaId      ;
                        proveedor.IdentifEspacio = costo.IdentifEspacio ;
                        proveedor.VigDesde       = costo.VigDesde       ;
                        proveedor.VigHasta       = costo.VigHasta       ;

                        daoProveedor.Create(proveedor, tran);

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
                    var costo = dao.Read(recId);

                    //Elimino todos los detalles (frecuencias y proveedores).
                    daoFrecuencia.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);
                    daoFrecuenciaVer.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);
                    daoProveedor.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);
                    daoProveedorVer.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);
                    daoVer.Delete(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND DatareaId = {3}", costo.IdentifEspacio, costo.VigDesde.ToString("yyyyMMdd"), costo.VigHasta.ToString("yyyyMMdd"), costo.DatareaId), tran);

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

        public static List<CostosProveedorDTO> ReadAllProveedor(string identifEspacio, DateTime vigDesde, DateTime vigHasta)
        {
            return daoProveedor.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}'", identifEspacio, vigDesde.ToString("yyyyMMdd"), vigHasta.ToString("yyyyMMdd")));
        }

        public static List<CostosFrecuenciaDTO> ReadAllFrecuencia(string identifEspacio, DateTime vigDesde, DateTime vigHasta)
        {
            return daoFrecuencia.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}'", identifEspacio, vigDesde.ToString("yyyyMMdd"), vigHasta.ToString("yyyyMMdd")));
        }

        public static List<CostosProveedorVersionDTO> ReadAllProveedorVersiones(string identifEspacio, DateTime vigDesde, DateTime vigHasta, decimal version)
        {
            return daoProveedorVer.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND Version = {3}", identifEspacio, vigDesde.ToString("yyyyMMdd"), vigHasta.ToString("yyyyMMdd"), version));
        }

        public static List<CostosFrecuenciaVersionDTO> ReadAllFrecuenciaVersiones(string identifEspacio, DateTime vigDesde, DateTime vigHasta, decimal version)
        {
            return daoFrecuenciaVer.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}' AND Version = {3}", identifEspacio, vigDesde.ToString("yyyyMMdd"), vigHasta.ToString("yyyyMMdd"), version));
        }

        public static List<CostosProveedorDTO> ReadAllProveedor(CostosDTO costos)
        {
            return daoProveedor.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}'", costos.IdentifEspacio, costos.VigDesde.ToString("yyyyMMdd"), costos.VigHasta.ToString("yyyyMMdd")));
        }

        public static List<CostosFrecuenciaDTO> ReadAllFrecuencia(CostosDTO costos)
        {
            return daoFrecuencia.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}'", costos.IdentifEspacio, costos.VigDesde.ToString("yyyyMMdd"), costos.VigHasta.ToString("yyyyMMdd")));
        }

        public static decimal GetLastVersion(string identifEspacio, DateTime vigDesde, DateTime vigHasta)
        {
            decimal lastVer = 0;
            List<CostoVersionDTO> versiones = daoVer.ReadAll(string.Format("IdentifEspacio = '{0}' AND VigDesde = '{1}' AND VigHasta = '{2}'", identifEspacio, vigDesde.ToString("yyyyMMdd"), vigHasta.ToString("yyyyMMdd")));

            if (versiones != null && versiones.Count>0)
            {
                foreach (var item in versiones)
	            {
                    if (item.Version > lastVer)
                        lastVer = item.Version;
	            }
                return lastVer;
            }
            else
            {
                return 0;
            }
        }

        public static void Commit(int idCosto, UsuariosDTO usuario)
        {
            using (SqlTransaction tran = dao.IniciarTransaccion())
            {
                try
                {
                    CostosDTO costo = dao.Read(idCosto);

                    costo.FecConfirmado = DateTime.Now;
                    costo.Confirmado    = usuario.UserName;
                    costo.Version       = GetLastVersion(costo.IdentifEspacio, costo.VigDesde, costo.VigHasta) + 1;

                    //Incremento version, y actualizo datos de confirmacion en costo.
                    dao.Update(costo, costo.RecId, tran);

                    //Creo el nuevoregistro de Costo Version.
                    daoVer.Create(new CostoVersionDTO(costo, costo.Version.Value), tran);

                    //Creo todos los registros de FrecuenciaVersion.
                    foreach (CostosFrecuenciaDTO f in ReadAllFrecuencia(costo))
                    {
                        CostosFrecuenciaVersionDTO fv = new CostosFrecuenciaVersionDTO(f, costo.Version.Value);
                        daoFrecuenciaVer.Create(fv, tran);
                    }

                    //Creo todos los registros de ProveedorVersion.
                    foreach (CostosProveedorDTO p in ReadAllProveedor(costo))
                    {
                        CostosProveedorVersionDTO pv = new CostosProveedorVersionDTO(p, costo.Version.Value);
                        daoProveedorVer.Create(pv, tran);
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
    }
}
