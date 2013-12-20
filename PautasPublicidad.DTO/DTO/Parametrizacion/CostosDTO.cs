using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class CostosDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifEspacio { get; set; }
        public DateTime VigDesde { get; set; }
        public DateTime VigHasta { get; set; }
        public string Frecuencia { get; set; }
        public string IdentifFrecuencia { get; set; }
        public string Horario { get; set; }
        public string Confirmado { get; set; }
        public DateTime? FecConfirmado { get; set; }
        public decimal? Version { get; set; }

    }
}
