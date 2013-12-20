using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class CostosFrecuenciaDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifEspacio { get; set; }
        public DateTime VigDesde { get; set; }
        public DateTime VigHasta { get; set; }
        public string DiaSemana { get; set; }
        public decimal? Dia { get; set; }
        public TimeSpan? HoraDesde { get; set; }
        public TimeSpan? HoraHasta { get; set; }
        public decimal Costo { get; set; }

    }
}
