using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class PiezasArteDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifPieza { get; set; }
        public string Name { get; set; }
        public string IdentifAnun { get; set; }
        public decimal? Duracion { get; set; }
        public string Extension { get; set; }
        public string IdentifTipoPieza { get; set; }
        public string OrdenProd { get; set; }
        public DateTime? VigDesde { get; set; }
        public DateTime? VigHasta { get; set; }

    }
}
