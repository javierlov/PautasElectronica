using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DAO;
using PautasPublicidad.DTO;
using System.Data;
using System.Data.SqlClient;

namespace PautasPublicidad.Business
{
    public static class TecnologiasSoporte
    {
        static TecnoSoporteDAO dao = DAOFactory.Get<TecnoSoporteDAO>();

        static public List<TecnoSoporteDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public TecnoSoporteDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public void Create(TecnoSoporteDTO grupoMedios)
        {
            CRUDHelper.Create(grupoMedios, dao);
        }

        static public void Update(TecnoSoporteDTO grupoMedios)
        {
            CRUDHelper.Update(grupoMedios, dao);
        }

        static public void Delete(int idTecnoSoporte)
        {
            CRUDHelper.Delete(idTecnoSoporte, dao);
        }

        public static void Create(string name, string identifTecno)
        {
            var o = new TecnoSoporteDTO()
            {
                Name         = name.Trim(),
                IdentifTecno = identifTecno.Trim()
            };
            Create(o);
        }
    }
}
