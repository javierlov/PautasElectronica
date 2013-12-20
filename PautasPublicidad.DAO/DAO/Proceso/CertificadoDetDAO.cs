using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class CertificadoDetDAO : DAOBase<CertificadoDetDTO>
    {
        //Methods...
        public void MoverCeros(string pautaId, SqlTransaction tran)
        {
            DAOHelper.EjecutarNonQuery(
                string.Format(
                "UPDATE [CertificadoDet] SET Costo=0, CostoOP=0, CostoUni=0, CostoOPUni=0 WHERE PautaId = '{0}'",
                pautaId), tran);
        }

        public int GetLastRecId()
        {
            object aux = DAOHelper.EjecutarScalar("SELECT MAX(RECID) FROM CERTIFICADODET");
            if (aux is DBNull)
                return 0;
            else
                return Convert.ToInt32(aux);
        }

    }
}
