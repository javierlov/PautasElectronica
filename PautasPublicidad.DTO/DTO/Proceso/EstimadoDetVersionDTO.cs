using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class EstimadoDetVersionDTO : TablaBase
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
        public decimal Version { get; set; }

        public EstimadoDetVersionDTO()
        {
        }

        public EstimadoDetVersionDTO(EstimadoDetDTO estDet)
        {
            DTOHelper.FillObjectByObject(estDet, this);
        }
    }
}
