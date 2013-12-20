using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class FrecuenciaDetDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifFrecuencia { get; set; }
        public string DiaSemana { get; set; }
        public int? Dia { get; set; }

    }
}
