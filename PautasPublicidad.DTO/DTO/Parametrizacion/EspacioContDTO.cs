using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class EspacioContDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifEspacio { get; set; }
        public string Name { get; set; }
        public string IdentifMedio { get; set; }
        public string IdentifTipoEsp { get; set; }
        public string IdentifFrecuencia { get; set; }
        public TimeSpan? HoraInicio { get; set; }
        public TimeSpan? HoraFin { get; set; }
        public string IdentifIntervalo { get; set; }
        public string FormatoOP { get; set; }
        public string Responsable { get; set; }
        public string Contacto { get; set; }
        public string Email { get; set; }
        public string Direccion { get; set; }
        public string Telefono { get; set; }


    }
}
