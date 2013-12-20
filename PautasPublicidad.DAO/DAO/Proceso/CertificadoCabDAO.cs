using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class CertificadoCabDAO : DAOBase<CertificadoCabDTO>
    {        
        //Methods...

        public DataView VistaCertificados()
        {
            DataTable dt = new DataTable();

            string sQuery = @"SELECT   oc.RecId                                                 ,
                                        oc.PautaId                                              ,
                                        oc.IdentifOrigen                                        ,
                                        oc.IdentifEspacio                                       ,
                                        IdentifFrecuencia                                       ,
                                        IdentifIntervalo                                        ,
                                        CONVERT(VARCHAR,oc.AnoMes) AS AnoMes                    ,
                                        SUBSTRING((CAST(oc.AnoMes AS VARCHAR(6))),1,4) AS 'año' ,
                                        SUBSTRING((CAST(oc.AnoMes AS VARCHAR(6))),5,2) AS 'mes' ,
                                        oc.HoraInicio,oc.HoraFin,oc.VigDesde                    ,
                                        oc.VigHasta                                             ,
                                        oc.UsuCosto                                             ,
                                        oc.CantSalidas                                          ,      
                                        oc.CertValido                                           ,       
                                        oc.FecCertValido 
                                   FROM dbo.CertificadoCab oc 
                                  WHERE oc.IdentifOrigen != 'NULL' 
                               ORDER BY oc.RecId DESC";

            DAOHelper.LlenarDataTable(ref dt, sQuery);
            return dt.DefaultView;
        }

        public string AnoMesCierreOrd()
        {
            return Convert.ToString(DAOHelper.EjecutarScalar("SELECT AnoMesCierreOrd FROM dbo.SetUp "));
        }

        public void LlenarTabla(DataTable Tabla, string sQuery)
        {
            DAOHelper.LlenarDataTable(ref Tabla, sQuery);
        }

        public void MoverCeros(string pautaId, SqlTransaction tran)
        {
            DAOHelper.EjecutarNonQuery(
                string.Format(
                "UPDATE [CertificadoCab] SET Costo=0, CostoOP=0, CostoUni=0, CostoOPUni=0 WHERE PautaId = '{0}'",
                pautaId), tran);
        }
    }
}
