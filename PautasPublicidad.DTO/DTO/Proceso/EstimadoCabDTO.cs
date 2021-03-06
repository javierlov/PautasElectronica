using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace PautasPublicidad.DTO
{
    [Serializable]
    public class EstimadoCabDTO : TablaBase
    {
        //Properties...
        //public int RecId { get; set; }
        //public int DatareaId { get; set; }
        public string PautaId { get; set; }
        public string IdentifEspacio { get; set; }
        public decimal AnoMes { get; set; }
        public string IdentifFrecuencia { get; set; }
        public TimeSpan HoraInicio { get; set; }
        public TimeSpan HoraFin { get; set; }
        public string IdentifIntervalo { get; set; }
        public decimal CantSalidas { get; set; }
        public decimal DuracionTot { get; set; }
        public decimal Costo { get; set; }
        public decimal CostoOp { get; set; }
        public decimal CostoUni { get; set; }
        public decimal CostoOpUni { get; set; }
        public DateTime VigDesde { get; set; }
        public DateTime VigHasta { get; set; }
        public decimal VersionCosto { get; set; }
        public DateTime FecCosto { get; set; }
        public string UsuCosto { get; set; }
        public DateTime? FecCierre { get; set; }
        public string UsuCierre { get; set; }
        public string Confirmado { get; set; }
        public DateTime FecConfirma { get; set; }
        public decimal Version { get; set; }
        public DateTime FecUltModif { get; set; }

        public EstimadoCabDTO()
        {
        }

        public EstimadoCabDTO(OrdenadoCabDTO ordenadoCab)
        {
            DTOHelper.FillObjectByObject(ordenadoCab, this);
            FecCierre = ordenadoCab.FecCierre.Value;
        }
    }
}
