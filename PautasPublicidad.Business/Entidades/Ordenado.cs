using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;

namespace PautasPublicidad.Business
{
    [Serializable]
    public class Ordenado
    {
        public OrdenadoCabDTO Cabecera { get; set; }
        public CostosDTO Costo { get; set; }
        public List<OrdenadoDetDTO> Lineas { get; set; }
        public List<OrdenadoSKUDTO> SKUs { get; set; }

        public Ordenado(string pautaId)
        {
            FillProperties(pautaId);
        }
        
        public void ProcesoDeCierreOrdenado(string usuario)
        {
            //Lockeo la seccion para que otro thread no me descontrole los calculos.
            lock (this)
            {
                try
                {
                    CalcularCosto(usuario);
                    GenerarOC(usuario);
                    ArmarEstimado(usuario);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en proceso de cierre.", ex);
                }
            }
        }

        private void CalcularCosto(string usuario)
        {
            Ordenados.CalcularCosto(Cabecera, Costo, Lineas, usuario);
            FillProperties(Cabecera.PautaId);
        }  

        private void FillProperties(string pautaId)
        {
            Cabecera = Ordenados.Read(pautaId);
            Lineas   = Ordenados.ReadAllLineas(Cabecera);
            SKUs     = Ordenados.ReadAllSKUs(Cabecera);
            Costo    = Business.Ordenados.FindCosto(Cabecera.IdentifEspacio, Convert.ToInt32(Cabecera.AnoMes.ToString().Substring(0, 4)), Convert.ToInt32(Cabecera.AnoMes.ToString().Substring(4, 2)));
        }

        private void GenerarOC(string usuario)
        {
            //ToDo.
        }

        private void ArmarEstimado(string usuario)
        {
            try
            {
                //Ordenado.
                Cabecera.FecCierre          = DateTime.Now;
                Cabecera.UsuCierre          = usuario;
                Estimado estimado           = new Estimado(this, usuario);
                EstimadoVersion estimadoVer = new EstimadoVersion(estimado, 1);
                
                Estimados.CierreOrdenado(estimado, estimadoVer, Cabecera);
            }
            catch (Exception ex)
            {
                throw new Exception("ArmarEstimado", ex);
            }
        }        
    }
}
