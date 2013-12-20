using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PautasPublicidad.DTO;
using System.Collections;

namespace PautasPublicidad.DAO
{
    public static class DAOFactory
    {
        //Utilizo una instancia singleton de cada dao.
        static ArrayList daoList = new ArrayList() 
        { 
            new GrupoMediosPubDAO(),
            new TecnoSoporteDAO(),
            new TipoMediosPubDAO(),
            new MediosPubDAO(),
            new AnunInternosDAO(),
            new AvisosDAO(),
            new AvisosIdAtenDAO(),
            new CostosDAO(),
            new CostosFrecuenciaDAO(),
            new CostosFrecuenciaVersionDAO(),
            new CostosProveedorDAO(),
            new CostosProveedorVersionDAO(),
            new CostoVersionDAO(),
            new EspacioContDAO(),
            new FormAvisoDAO(),
            new FrecuenciaDAO(),
            new FrecuenciaDetDAO(),
            new IdentAtencionDAO(),
            new IntervaloDAO(),
            new MonedasDAO(),
            new PiezasArteDAO(),
            new PiezasArteSKUDAO(),
            new ProveedorDAO(),
            new SetUpDAO(),
            new SKUDAO(),
            new TipoCambioDAO(),
            new TipoEspacioDAO(),
            new TipoPiezaDAO(),
            new UsuariosDAO(),
            new OrigenDAO(),
            new EntornoDAO(),
            new OrdenadoCabDAO(),
            new OrdenadoDetDAO(),
            new OrdenadoSKUDAO(),
            new EstimadoCabDAO(),
            new EstimadoCabVersionDAO(),
            new EstimadoDetDAO(),
            new EstimadoDetVersionDAO(),
            new EstimadoSKUDAO(),
            new EstimadoSKUVersionDAO(),
            new CertificadoCabDAO(),
            new CertificadoDetDAO(),
            new CertificadoSKUDAO(),
            new EmpresaDAO()
        };

        public static object Get(string typeName)
        {
            return Get(Type.GetType("PautasPublicidad.DAO." + typeName));
        }

        //Retorno el dao correspondiente según el tipo solicitado.
        public static T Get<T>()
        {
            foreach (var item in daoList)
                if (item is T)
                    return (T)item;
            return default(T);
        }

        /// <summary>
        /// Obtener el DAO correspondiente a la ClaseDAO enviada por parametro.
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public static object Get(Type t)
        {
            foreach (var item in daoList)
                if (item.GetType() == t)
                    return item;
            return null;
        }
    }
}
