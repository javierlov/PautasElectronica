using System                    ;
using System.Collections.Generic;
using System.Linq               ;
using System.Text               ;
using PautasPublicidad.DAO      ;
using PautasPublicidad.DTO      ;
using System.Data.SqlClient     ;

namespace PautasPublicidad.Business
{
    [Serializable]
    public class Certificado
    {
        public CertificadoCabDTO Cabecera     { get; set; }
        public CostosDTO Costo                { get; set; }
        public List<CertificadoDetDTO> Lineas { get; set; }
        public List<CertificadoSKUDTO> SKUs   { get; set; }

        static CertificadoCabDAO dao        = DAOFactory.Get<CertificadoCabDAO>();
        static CertificadoDetDAO daoDetalle = DAOFactory.Get<CertificadoDetDAO>();
        static CertificadoSKUDAO daoSKU     = DAOFactory.Get<CertificadoSKUDAO>();
        static CostosDAO daoCosto           = DAOFactory.Get<CostosDAO>();
        static SetUpDAO daoSetUp            = DAOFactory.Get<SetUpDAO>();

        public enum eSemanaMes
        {
            SEMANA,
            MES
        }

        internal Certificado()
        {

        }

        static public List<CertificadoCabDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        public Certificado(Estimado estimado, string usuario)
        {
            CertificadoDetDTO estDet;
            CertificadoSKUDTO estSKU;

            Cabecera = new CertificadoCabDTO(estimado.Cabecera);
            Lineas   = new List<CertificadoDetDTO>();

            foreach (var ordDet in estimado.Lineas)
            {
                estDet       = new CertificadoDetDTO(ordDet);
                estDet.RecId = 0;

                Lineas.Add(estDet);
            }

            SKUs = new List<CertificadoSKUDTO>();

            foreach (var ordSKU in estimado.SKUs)
            {
                estSKU       = new CertificadoSKUDTO(ordSKU);
                estSKU.RecId = 0;

                SKUs.Add(estSKU);
            }
        }
    }
}
