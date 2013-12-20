using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using PautasPublicidad.DTO;
using PautasPublicidad.DAO;
using System.Xml;

namespace PautasPublicidad.Business
{
    public static class BusinessMapper
    {
        public static string MapperInfoXmlPath { get; set; }
        public static string AbmConfigXmlPath { get; set; }

        private static Dictionary<string, MapInfo> mapsInfo;

        public enum eEntities
        {
            GrupoMediosPub,
            TecnoSoporte,
            TipoMediosPub,
            MediosPub,
            AnunInternos,
            Avisos,
            AvisosIdAten,
            Costos,
            CostosFrecuencia,
            CostosFrecuenciaVersion,
            CostosProveedor,
            CostosProveedorVersion,
            CostoVersion,
            EspacioCont,
            FormAviso,
            Frecuencia,
            FrecuenciaDet,
            IdentAtencion,
            Intervalo,
            Monedas,
            PiezasArte,
            PiezasArteSKU,
            Proveedor,
            SetUp,
            SKU,
            TipoCambio,
            TipoEspacio,
            TipoPieza,
            Usuarios,
            Origen,
            Entorno,
            Empresa
        }

        public static T StringToEnum<T>(string name)
        {
            return (T)Enum.Parse(typeof(T), name);
        }

        public static eEntities GetEntityByName(string entityName)
        {
            return StringToEnum<eEntities>(entityName);
        }

        public static dynamic GetDaoByEntity(eEntities entity)
        {
            return DAOFactory.Get(entity.ToString() + "DAO");
        }

        public static MapInfo GetMapInfo(string entityName)
        {
            if (mapsInfo == null)
                FillMapInfo();

            return mapsInfo[entityName];
        }

        private static void FillMapInfo()
        {
            if (MapperInfoXmlPath == null || MapperInfoXmlPath == string.Empty)
                throw new Exception("Path del archivo MapperInfo.xml sin definir.");

            mapsInfo = new Dictionary<string, MapInfo>();
            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(MapperInfoXmlPath);

            foreach (XmlNode node in xDoc.SelectNodes("Mappings/MapInfo"))
            {
                mapsInfo.Add(node.Attributes["EntityName"].Value,
                             new MapInfo(node.Attributes["EntityName"].Value,
                             node.Attributes["TextField"].Value,
                             node.Attributes["ValueField"].Value,
                             DAOFactory.Get(node.Attributes["EntityName"].Value + "DAO"),
                             node.Attributes["AbmUrl"].Value));
            }
        }
    }

    [Serializable]
    public class MapInfo
    {
        public MapInfo(string entityName, string textField, string valueField, object daoHandler, string abmUrl)
        {
            EntityName       = entityName;
            EntityTextField  = textField;
            EntityValueField = valueField;
            DAOHandler       = daoHandler;
            ABMUrl           = abmUrl;
        }
        public MapInfo(string entityName, string textField, string valueField, object daoHandler)
        {
            EntityName       = entityName;
            EntityTextField  = textField;
            EntityValueField = valueField;
            DAOHandler       = daoHandler;
        }

        public string EntityName { get; set; }
        public string EntityTextField { get; set; }
        public string EntityValueField { get; set; }
        public dynamic DAOHandler { get; set; }
        public string ABMUrl { get; set; }
    }

    [global::System.Serializable]
    public class AbmConfigXMLException : Exception
    {
        public AbmConfigXMLException() { }
        public AbmConfigXMLException(string message) : base(message) { }
        public AbmConfigXMLException(string message, Exception inner) : base(message, inner) { }
        protected AbmConfigXMLException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }
}
