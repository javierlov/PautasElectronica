using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class UsuariosDTO : TablaBase
    {
        public string UserName { get; set; }
    }
}
