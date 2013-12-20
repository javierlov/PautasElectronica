using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [SortField(Property = "IdentifTipoEsp")]
    [Serializable]
    public class TipoEspacioDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifTipoEsp { get; set; }
        public string Name { get; set; }
        public bool Frecuencia { get; set; }
        public bool Hora { get; set; }
        public bool Intervalo { get; set; }

    }
}
