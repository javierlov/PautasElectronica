using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class ProveedorDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifProv { get; set; }
        public string Name { get; set; }

    }
}
