using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class sysdiagramsDTO : TablaBase
    {
        //Properties...
        public string name { get; set; }
        public int principal_id { get; set; }
        public int diagram_id { get; set; }
        public int version { get; set; }
        public object definition { get; set; }

    }
}
