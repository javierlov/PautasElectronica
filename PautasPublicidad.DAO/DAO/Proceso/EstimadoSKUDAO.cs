using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class EstimadoSKUDAO : DAOBase<EstimadoSKUDTO>
    {
        //Methods...
        public DataTable GetSKUByPiezasArte(List<EstimadoDetDTO> lineas)
        {
            var dt = new DataTable();
            string strIn = "";
            string last = "";
            string qry = @"SELECT *, 0 as CantSalidas  
                            FROM SKU  
                            JOIN PiezasArteSKU ON SKU.IdentifSKU=PiezasArteSKU.IdentifSKU
                            JOIN Avisos ON PiezasArteSKU.IdentifPieza=Avisos.IdentifPieza
                            WHERE 
                            PiezasArteSKU.TipoProd = 'PRIMARIO'
                            AND IdentifAviso IN ({0})"; //'av1', 'AV_1'

            //Si no tengo linas de detalle, tampoco productos.
            if (lineas.Count == 0)
                return dt;

            lineas.Sort((x, y) => x.IdentifAviso.CompareTo(y.IdentifAviso));
            lineas.ForEach((x) =>
            {
                if (x.IdentifAviso != last)
                {
                    last = x.IdentifAviso;
                    strIn += string.Format("'{0}',", x.IdentifAviso);
                }
            });


            qry = string.Format(qry, strIn.Substring(0, strIn.Length - 1)) + " AND SKU.DatareaId = " + DAOHelper.DatareaId;
            DAOHelper.LlenarDataTable(ref dt, qry);

            return dt;
        }

        public void MoverCeros(string pautaId, SqlTransaction tran)
        {
            DAOHelper.EjecutarNonQuery(
                string.Format(
                "UPDATE [EstimadoSKU] SET Costo=0, CostoOP=0, CostoUni=0, CostoOPUni=0 WHERE PautaId = '{0}' AND DatareaId = {1}",
                pautaId, DAOHelper.DatareaId), tran);
        }
    }
}

/*
        public int RecId { get; set; }
        public int DatareaId { get; set; }
        public string PautaId { get; set; }
        public string IdentifAviso { get; set; }
        public string IdentifSKU { get; set; }
        public decimal Duracion { get; set; }
        public decimal Costo { get; set; }
        public decimal CostoOp { get; set; }
        public decimal CostoUni { get; set; }
        public decimal CostoOpUni { get; set; }
        public decimal CantSalidas { get; set; }

*/