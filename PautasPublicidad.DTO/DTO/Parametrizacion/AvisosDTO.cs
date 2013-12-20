using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class AvisosDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifAviso { get; set; }
        public string Name { get; set; }
        public string IdentifEspacio { get; set; }
        public string IdentifFormAviso { get; set; }
        public string IdentifPieza { get; set; }
        public decimal? Duracion { get; set; }
        public string EtiquetaProd { get; set; }
        public string Zocalo { get; set; }
        public string NroIngesta { get; set; }
        public DateTime VigDesde { get; set; }
        public DateTime VigHasta { get; set; }

    }
}
