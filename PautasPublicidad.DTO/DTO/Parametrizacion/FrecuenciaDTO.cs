using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class FrecuenciaDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifFrecuencia { get; set; }
        public string Name { get; set; }
        public string SemMes { get; set; }

    }
}
