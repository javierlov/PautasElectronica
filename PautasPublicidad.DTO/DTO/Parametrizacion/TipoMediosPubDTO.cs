using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class TipoMediosPubDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifTipo { get; set; }
        public string Name { get; set; }
        public string IdentifTecno { get; set; }

    }
}
