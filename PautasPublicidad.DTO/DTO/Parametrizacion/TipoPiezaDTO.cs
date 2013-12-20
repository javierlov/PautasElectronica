using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [SortField(Property = "IdentifTipoPieza")]
    [Serializable]
    public class TipoPiezaDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifTipoPieza { get; set; }
        public string Name { get; set; }
        public bool Duracion { get; set; }

    }
}
