using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
   [Serializable]
    public class SetUpDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public decimal NumPauta { get; set; }
        public string InventIDOC { get; set; }
        public string Sector { get; set; }
        public decimal PorcIVA { get; set; }
        public decimal AnoMesCierreEst { get; set; }
        public decimal AnoMesCierreOrd { get; set; }
    }
}
