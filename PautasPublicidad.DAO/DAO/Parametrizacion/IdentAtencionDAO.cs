using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class IdentAtencionDAO : DAOBase<IdentAtencionDTO>
    {
        //Methods...
        public DataTable FindPaginado(string filter, Int64 startIndex, Int64 endIndex)
        {
            string qry = @"SELECT [IdentifIdentAte]
                            , [Name]
                            , IsNull((SELECT COUNT(1)
	                                FROM AvisosIdAten 
                                    JOIN Avisos ON AvisosIdAten.IdentifAviso = Avisos.IdentifAviso 
	                                WHERE AvisosIdAten.IdentifIdentAte = st.IdentifIdentAte 
                                      AND VigHasta > GETDATE() AND Avisos.DatareaId = " + DAOHelper.DatareaId + @" 
                                    ), 0) as [Asignado]
                             FROM (select [IdentifIdentAte], [Name], row_number()over(order by t.[IdentifIdentAte]) as [rn] 
                                from [IdentAtencion] as t 
                                where (([IdentifIdentAte] + ' ' + [Name]) LIKE '%{0}%' AND Estado = 1 AND t.DatareaId = " + DAOHelper.DatareaId + @")) as st 
                            WHERE st.[rn] between {1} and {2}";

            var dt = new DataTable();
            qry = string.Format(qry, filter, startIndex, endIndex);
            DAOHelper.LlenarDataTable(ref dt, qry);

            return dt;
        }

        public DataTable FindValue(string IdentifIdentAtencion)
        {
            string qry = @"SELECT [IdentifIdentAte]
                            , [Name] 
                            , IsNull((SELECT COUNT(1)
	                                FROM AvisosIdAten 
                                    JOIN Avisos ON AvisosIdAten.IdentifAviso = Avisos.IdentifAviso 
	                                WHERE AvisosIdAten.IdentifIdentAte = IdentAtencion.IdentifIdentAte 
                                      AND VigHasta > GETDATE()
                                    ), 0) as [Asignado]
                        FROM [IdentAtencion] 
                        where [IdentifIdentAte] = '{0}' AND Estado = 1 AND IdentAtencion.DatareaId = " + DAOHelper.DatareaId;
            
            var dt = new DataTable();
            qry = string.Format(qry, IdentifIdentAtencion);
            DAOHelper.LlenarDataTable(ref dt, qry);

            return dt;
        }

        public bool ExisteAvisoEnVigencia(string IdentifIdentAtencion)
        {
            string qry = @"SELECT COUNT(1)
	                        FROM AvisosIdAten 
                            JOIN Avisos ON AvisosIdAten.IdentifAviso = Avisos.IdentifAviso 
	                        WHERE AvisosIdAten.IdentifIdentAte = '{0}' 
                            AND Avisos.VigHasta > GETDATE() AND Avisos.DatareaId = " + DAOHelper.DatareaId;

            qry = string.Format(qry, IdentifIdentAtencion);
            return Convert.ToInt32(DAOHelper.EjecutarScalar(qry)) > 0; 
        }
    }
}
/*
        public int RecId { get; set; }
        public int DatareaId { get; set; }
        public string IdentifIdentAte { get; set; }
        public string Name { get; set; }
        public string TipIdentif { get; set; }
        public string TipoCDN { get; set; }
        public string Telefono { get; set; }
        public decimal DNIS { get; set; }

*/
