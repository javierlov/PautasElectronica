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
    public static class MediosPublicitarios
    {
        static MediosPubDAO dao           = DAOFactory.Get<MediosPubDAO>();
        static GrupoMediosPubDAO daoGrupo = DAOFactory.Get<GrupoMediosPubDAO>();
        static TipoMediosPubDAO daoTipo   = DAOFactory.Get<TipoMediosPubDAO>();

        #region Grupo Medios Publicitarios
        static public List<GrupoMediosPubDTO> ReadAllGrupo(string sWhere)
        {
            return daoGrupo.ReadAll(sWhere);
        }

        static public GrupoMediosPubDTO ReadGrupo(int id)
        {
            return daoGrupo.Read(id);
        }

        static public void CreateGrupo(GrupoMediosPubDTO grupoMedios)
        {
            CRUDHelper.Create(grupoMedios, daoGrupo);
        }

        static public void UpdateGrupo(GrupoMediosPubDTO grupoMedios)
        {
            CRUDHelper.Update(grupoMedios, daoGrupo);
        }

        static public void DeleteGrupo(int idGrupoMediosPub)
        {
            CRUDHelper.Delete(idGrupoMediosPub, daoGrupo);
        }

        public static void CreateGrupo(string name, string identifGrupo)
        {
            var o = new GrupoMediosPubDTO()
            {
                Name         = name.Trim(),
                IdentifGrupo = identifGrupo.Trim()
            };
            CreateGrupo(o);
        }
        #endregion

        #region Tipo Medios Publicitarios
        static public List<TipoMediosPubDTO> ReadAllTipo(string sWhere)
        {
            return daoTipo.ReadAll(sWhere);
        }

        static public TipoMediosPubDTO ReadTipo(int id)
        {
            return daoTipo.Read(id);
        }

        static public void CreateTipo(TipoMediosPubDTO grupoMedios)
        {
            CRUDHelper.Create(grupoMedios, daoTipo);
        }

        static public void UpdateTipo(TipoMediosPubDTO grupoMedios)
        {
            CRUDHelper.Update(grupoMedios, daoTipo);
        }

        static public void DeleteTipo(int idGrupoMediosPub)
        {
            CRUDHelper.Delete(idGrupoMediosPub, daoTipo);
        }

        public static void CreateTipo(string identifTipo, string name, string identifTecno)
        {
            var o = new TipoMediosPubDTO()
            {
                IdentifTipo  = identifTipo.Trim(),
                Name         = name.Trim(),
                IdentifTecno = identifTecno.Trim()
            };
            CreateTipo(o);
        }
        #endregion

        #region Medios Publicitarios
        static public List<MediosPubDTO> ReadAll(string sWhere)
        {
            return dao.ReadAll(sWhere);
        }

        static public MediosPubDTO Read(int id)
        {
            return dao.Read(id);
        }

        static public void Create(MediosPubDTO mediosPub)
        {
            CRUDHelper.Create(mediosPub, dao);
        }

        static public void Update(MediosPubDTO mediosPub)
        {
            CRUDHelper.Update(mediosPub, dao);
        }

        static public void Delete(int idMediosPub)
        {
            CRUDHelper.Delete(idMediosPub, dao);
        }

        public static void Create(string identifMedio, string name, string identifGrupo, string identifTipo)
        {
            var o = new MediosPubDTO()
            {  
                IdentifMedio = identifMedio.Trim(), 
                Name         = name.Trim(),
                IdentifGrupo = identifGrupo.Trim(),
                IdentifTipo  = identifTipo.Trim()
            };
            Create(o);
        }
        #endregion
    }
}
