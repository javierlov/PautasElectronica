using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class TablaBase
    {
        ////Campos genéricos para todas las tablas.
        public int DatareaId { get; set; } //ToDo: Cambiar a string!
        public int RecId { get; set; }
    }
}
