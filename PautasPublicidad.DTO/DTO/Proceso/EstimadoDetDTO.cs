using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class EstimadoDetDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string PautaId { get; set; }
        public DateTime Fecha { get; set; }
        public decimal Dia { get; set; }
        public string DiaSemana { get; set; }
        public TimeSpan Hora { get; set; }
        public decimal Salida { get; set; }
        public decimal? Duracion { get; set; }
        public string IdentifAviso { get; set; }        
        public decimal Costo { get; set; }
        public decimal CostoOp { get; set; }
        public decimal CostoUni { get; set; }
        public decimal CostoOpUni { get; set; }

        public EstimadoDetDTO()
        {
        }

        public EstimadoDetDTO(OrdenadoDetDTO ordDet)
        {
            DTOHelper.FillObjectByObject(ordDet, this);
        }
    }

    public class MiEstimadoDetDTO : EstimadoDetDTO
    {
        public string CodigoAviso { get; set; }
    }
}
