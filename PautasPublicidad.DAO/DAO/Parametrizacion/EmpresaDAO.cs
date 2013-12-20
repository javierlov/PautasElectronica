using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Data.SqlClient;
using System.Data;

namespace PautasPublicidad.DAO
{
    public class EmpresaDAO : DAOBase<EmpresaDTO>
    {
        //Methods...
        public List<EmpresaDTO> GetEmpresas()
        {
            List<EmpresaDTO> empresas = new List<EmpresaDTO>();
            EmpresaDTO empresa;
            var dt = new DataTable();
            string qry = @"SELECT * FROM Empresa";

            DAOHelper.LlenarDataTable(ref dt, qry);

            foreach (DataRow dr in dt.Rows)
            {
                empresa = new EmpresaDTO();
                DAOHelper.PoblarObjetoDesdeDataRow(empresa, dr);
                empresas.Add(empresa);
            }

            return empresas;
        }
    }
}
