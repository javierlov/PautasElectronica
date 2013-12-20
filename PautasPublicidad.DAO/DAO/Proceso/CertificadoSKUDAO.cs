using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class CertificadoSKUDAO : DAOBase<CertificadoSKUDTO>
    {
        //Methods...
        public DataTable GetSKUByPiezasArte(List<CertificadoDetDTO> lineas)
        {
            try
            {
                var dt = new DataTable();
                string strIn = "";
                string last = "";
                string qry = @"SELECT *, 0 as CantSalidas  
                            FROM SKU  
                            JOIN PiezasArteSKU ON SKU.IdentifSKU=PiezasArteSKU.IdentifSKU
                            JOIN Avisos ON PiezasArteSKU.IdentifPieza=Avisos.IdentifPieza
                            WHERE 
                            PiezasArteSKU.TipoProd = 'PRIMARIO'
                            AND IdentifAviso IN ({0});"; //'av1', 'AV_1'

                //Si no tengo linas de detalle, tampoco productos.
                if (lineas.Count == 0)
                    return dt;

                lineas.Sort((x, y) => x.IdentifAviso.CompareTo(y.IdentifAviso));
                lineas.ForEach((x) =>
                {
                    if (x.IdentifAviso != last)
                    {
                        last = x.IdentifAviso;
                        strIn += string.Format("'{0}',", x.IdentifAviso);
                    }
                });


                qry = string.Format(qry, strIn.Substring(0, strIn.Length - 1));
                DAOHelper.LlenarDataTable(ref dt, qry);

                return dt;
            }
            catch (Exception ex)
            {
                string a = ex.ToString();

                DataTable dtx = new DataTable();

                return dtx;

                //throw new Exception("Debe seleccionar un aviso");
            }
        }

        public void MoverCeros(string pautaId, SqlTransaction tran)
        {
            DAOHelper.EjecutarNonQuery(
                string.Format(
                "UPDATE [CertificadoSKU] SET Costo=0, CostoOP=0, CostoUni=0, CostoOPUni=0 WHERE PautaId = '{0}'",
                pautaId), tran);
        }
    }
}
