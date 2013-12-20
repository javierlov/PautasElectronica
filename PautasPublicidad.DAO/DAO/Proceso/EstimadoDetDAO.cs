using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class EstimadoDetDAO : DAOBase<EstimadoDetDTO>
    {
        //Methods...
        public void MoverCeros(string pautaId, SqlTransaction tran)
        {
            DAOHelper.EjecutarNonQuery(
                string.Format(
                "UPDATE [EstimadoDet] SET Costo=0, CostoOP=0, CostoUni=0, CostoOPUni=0 WHERE PautaId = '{0}' AND DatareaId = {1}",
                pautaId, DAOHelper.DatareaId), tran);
        }

        public int GetLastRecId()
        {
            object aux = DAOHelper.EjecutarScalar("SELECT MAX(RECID) FROM ESTIMADODET");
            if (aux is DBNull)
                return 0;
            else
                return Convert.ToInt32(aux);
        }

    }
}

/*
        public int RecId { get; set; }
        public int DatareaId { get; set; }
        public string PautaId { get; set; }
        public decimal Dia { get; set; }
        public string DiaSemana { get; set; }
        public DateTime Hora { get; set; }
        public decimal Salida { get; set; }
        public decimal Duracion { get; set; }
        public string IdentifAviso { get; set; }
        public decimal CostoOp { get; set; }
        public decimal Costo { get; set; }
        public decimal CostoUni { get; set; }
        public decimal CostoOpUni { get; set; }

*/