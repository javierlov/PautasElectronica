using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class IdentAtencionDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifIdentAte { get; set; }
        public string Name { get; set; }
        public string TipIdentif { get; set; }
        public string TipoCDN { get; set; }
        public string Telefono { get; set; }
        public decimal? DNIS { get; set; }
        public bool Estado { get; set; }
    }
}
