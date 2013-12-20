using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;

namespace PautasPublicidad.Business
{
    [Serializable]
    public class EstimadoVersion
    {
        public EstimadoCabVersionDTO Cabecera { get; set; }
        public List<EstimadoDetVersionDTO> Lineas { get; set; }
        public List<EstimadoSKUVersionDTO> SKUs { get; set; }

        public EstimadoVersion(Estimado estimado, decimal version)
        {
            EstimadoDetVersionDTO estDetVer;
            EstimadoSKUVersionDTO estSKUVer;

            Cabecera         = new EstimadoCabVersionDTO(estimado.Cabecera);
            Cabecera.Version = version;

            Lineas = new List<EstimadoDetVersionDTO>();

            foreach (var estDet in estimado.Lineas)
            {
                estDetVer         = new EstimadoDetVersionDTO(estDet);
                estDetVer.Version = version;
                estDetVer.RecId   = 0;
                Lineas.Add(estDetVer);
            }

            SKUs = new List<EstimadoSKUVersionDTO>();

            foreach (var estSKU in estimado.SKUs)
            {
                estSKUVer         = new EstimadoSKUVersionDTO(estSKU);
                estSKUVer.Version = version;
                estSKUVer.RecId   = 0;
                SKUs.Add(estSKUVer);
            }
        }
    }
}
