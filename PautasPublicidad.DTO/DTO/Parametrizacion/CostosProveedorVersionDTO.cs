using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class CostosProveedorVersionDTO : TablaBase
    {
        private CostosProveedorDTO p;

        public CostosProveedorVersionDTO() { }

        public CostosProveedorVersionDTO(CostosProveedorDTO p, decimal version)
        {
            this.p = p;

            IdentifEspacio = p.IdentifEspacio;
            VigDesde = p.VigDesde;
            VigHasta = p.VigHasta;
            IdentifProv = p.IdentifProv;
            Categoria = p.Categoria;
            IncluidoOP = p.IncluidoOP;
            Estimado = p.Estimado;
            TipoCosto = p.TipoCosto;
            IdentifMon = p.IdentifMon;
            GrossingUp = p.GrossingUp;
            Costo = p.Costo;
            GeneraOC = p.GeneraOC;
            Version = version;
        }
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string IdentifEspacio { get; set; }
        public DateTime VigDesde { get; set; }
        public DateTime VigHasta { get; set; }
        public string IdentifProv { get; set; }
        public string Categoria { get; set; }
        public bool IncluidoOP { get; set; }
        public bool Estimado { get; set; }
        public string TipoCosto { get; set; }
        public string IdentifMon { get; set; }
        public decimal GrossingUp { get; set; }
        public decimal Costo { get; set; }
        public bool? GeneraOC { get; set; }
        public decimal Version { get; set; }

    }
}
