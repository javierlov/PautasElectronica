using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class OrdenadoCabDAO : DAOBase<OrdenadoCabDTO>
    {
        public void LlenarTabla(DataTable Tabla, string sQuery)
        {
             DAOHelper.LlenarDataTable(ref Tabla,sQuery);
        }

        public DataView VistaOrdenados()
        {
            DataTable dt = new DataTable();
            DAOHelper.LlenarDataTable(ref dt,
                "SELECT  oc.RecId ,oc.PautaId ,oc.IdentifEspacio ,IdentifFrecuencia,IdentifIntervalo,CONVERT(VARCHAR,oc.AnoMes) AS AnoMes ,SUBSTRING((CAST(oc.AnoMes AS VARCHAR(6))),1,4) AS 'año',SUBSTRING((CAST(oc.AnoMes AS VARCHAR(6))),5,2) AS 'mes',oc.HoraInicio,oc.HoraFin,oc.VigDesde,oc.VigHasta,oc.UsuCosto,oc.CantSalidas,oc.UsuCierre,oc.FecCierre FROM    dbo.OrdenadoCab oc ORDER BY oc.RecId DESC");
                return dt.DefaultView;
        }
        public DataView VistaCostosConfirmados()
        {
            DataTable dtconf = new DataTable();
            DAOHelper.LlenarDataTable(ref dtconf,
                "SELECT DISTINCT ct.RecId ,ct.IdentifEspacio ,ct.IdentifFrecuencia ,SUBSTRING(( CAST(ct.VigDesde AS VARCHAR(8)) ), 1, 4) AS 'año' , SUBSTRING(( CAST(ct.VigDesde AS VARCHAR(8)) ), 6, 2) AS 'mes' ,ct.VigDesde ,ct.VigHasta ,ct.Confirmado ,ct.FecConfirmado FROM    dbo.Costos ct , dbo.OrdenadoCab oc WHERE   ct.Confirmado != '' AND ( CONVERT(VARCHAR, ct.IdentifEspacio) + CONVERT(VARCHAR, ct.VigDesde) + CONVERT(VARCHAR, ct.VigHasta) ) NOT IN ( SELECT  ( CONVERT(VARCHAR, oc.IdentifEspacio) + CONVERT(VARCHAR, oc.VigDesde) + CONVERT(VARCHAR, oc.VigHasta) ) FROM    dbo.OrdenadoCab oc , Costos ct WHERE   ct.VigDesde = oc.VigDesde AND ct.VigHasta = oc.VigHasta AND ct.IdentifEspacio = oc.IdentifEspacio )");
            return dtconf.DefaultView;
        
        }

        public DateTime FechaServer
        {
            get
            {
                return (DateTime)DAOHelper.EjecutarScalar("SELECT GETDATE()");
            }
        }


        ///ToDo
        public string AnoMesCierreOrd()
        {
            return Convert.ToString(DAOHelper.EjecutarScalar("SELECT AnoMesCierreOrd FROM dbo.SetUp "));
        }

        //Methods...
        public void MoverCeros(string pautaId, SqlTransaction tran)
        {
            DAOHelper.EjecutarNonQuery(
                string.Format(
                "UPDATE [OrdenadoCab] SET Costo=0, CostoOP=0, CostoUni=0, CostoOPUni=0 WHERE PautaId = '{0}' AND DatareaId = {1}", 
                pautaId, DAOHelper.DatareaId), tran);
        }
    }
}

/*
        public int RecId { get; set; }
        public int DatareaId { get; set; }
        public string PautaId { get; set; }
        public string IdentifEspacio { get; set; }
        public decimal AnoMes { get; set; }
        public string IdentifFrecuencia { get; set; }
        public DateTime HoraInicio { get; set; }
        public DateTime HoraFin { get; set; }
        public string IdentifIntervalo { get; set; }
        public decimal CantSalidas { get; set; }
        public decimal DuracionTot { get; set; }
        public decimal Costo { get; set; }
        public decimal CostoOp { get; set; }
        public decimal CostoUni { get; set; }
        public decimal CostoOpUni { get; set; }
        public DateTime VigDesde { get; set; }
        public DateTime VigHasta { get; set; }
        public decimal VersionCosto { get; set; }
        public DateTime FecCosto { get; set; }
        public string UsuCosto { get; set; }
        public DateTime FecCierre { get; set; }
        public string UsuCierre { get; set; }

*/