using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;

namespace PautasPublicidad.Business
{
    [Serializable]
    public class Estimado
    {
        public EstimadoCabDTO Cabecera { get; set; }
        public CostosDTO Costo { get; set; }
        public List<EstimadoDetDTO> Lineas { get; set; }
        public List<EstimadoSKUDTO> SKUs { get; set; }

        internal Estimado()
        {

        }

        public void ProcesoDeCierreEstimado(string usuario)
        {
            //Lockeo la seccion para que otro thread no me descontrole los calculos.
            lock (this)
            {
                try
                {
                    CalcularCosto(usuario);
                    ArmarCertificado(usuario);
                }
                catch (Exception ex)
                {
                    throw new Exception("ProcesoDeCierreOrdenado", ex);
                }
            }
        }

        private void ArmarCertificado(string usuario)
        {
            try
            {
                //Ordenado.
                Cabecera.FecCierre      = DateTime.Now;
                Cabecera.UsuCierre      = usuario;
                Certificado certificado = new Certificado(this, usuario);

                Certificados.CierreEstimado(certificado, Cabecera);
            }
            catch (Exception ex)
            {
                throw new Exception("ArmarEstimado", ex);
            }
        }

        private void CalcularCosto(string usuario)
        {
            Estimados.CalcularCosto(Cabecera, Costo, Lineas, usuario);
            FillProperties(Cabecera.PautaId);
        }  

        public Estimado(Ordenado ordenado, string usuario)
        {
            EstimadoDetDTO estDet;
            EstimadoSKUDTO estSKU;

            Cabecera             = new EstimadoCabDTO(ordenado.Cabecera);
            Cabecera.Confirmado  = usuario;
            Cabecera.FecConfirma = DateTime.Now;
            Cabecera.Version     = 1;
            Cabecera.FecUltModif = DateTime.Now;
            Cabecera.FecCierre   = null;
            Cabecera.UsuCierre   = string.Empty;

            Lineas = new List<EstimadoDetDTO>();
            foreach (var ordDet in ordenado.Lineas)
            {
                estDet       = new EstimadoDetDTO(ordDet);
                estDet.RecId = 0;
                Lineas.Add(estDet);
            }

            SKUs = new List<EstimadoSKUDTO>();
            foreach (var ordSKU in ordenado.SKUs)
            {
                estSKU       = new EstimadoSKUDTO(ordSKU);
                estSKU.RecId = 0;
                SKUs.Add(estSKU);
            }
        }

        public Estimado(string pautaId)
        {
            FillProperties(pautaId);
        }

        private void FillProperties(string pautaId)
        {
            Cabecera = Estimados.Read(pautaId);
            Lineas   = Estimados.ReadAllLineas(Cabecera);
            SKUs     = Estimados.ReadAllSKUs(Cabecera);
            Costo    = Business.Ordenados.FindCosto(Cabecera.IdentifEspacio, Convert.ToInt32(Cabecera.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(Cabecera.AnoMes.ToString().Substring(4, 2)));
        }
    }
}
