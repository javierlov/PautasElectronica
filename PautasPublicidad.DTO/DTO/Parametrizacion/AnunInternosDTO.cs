using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [SortField(Property = "IdentifAnun")]
    [Serializable]
    public class AnunInternosDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifAnun { get; set; }
        public string Name { get; set; }
        public string Entorno { get; set; }

    }
}
