using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class EmpresaDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifEmpresa { get; set; }
        public string Name { get; set; }
        public string Leyenda { get; set; }
        public string Logo { get; set; }

    }
}
