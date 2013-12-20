using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PautasPublicidad.DTO;
using System.Globalization;
using System.Web.UI;

namespace PautasPublicidad.Web
{
    public static class ProcesoHelper
    {
        internal static int NextTempRecId(StateBag viewState)
        {
            
            if (viewState["TempRecId"] != null)
            {
                int RecId = Convert.ToInt32(viewState["TempRecId"]) + 1;

                viewState.Add("TempRecId", RecId);
                return RecId;
            }
            else
            {
                viewState.Add("TempRecId", 1);
                return 1;
            }
        }

        internal static List<EstimadoDetDTO> CopiarPeriodos(DateTime fOrigenDesde, DateTime fOrigenHasta,
    DateTime fDestinoDesde, List<EstimadoDetDTO> lineas, StateBag viewState)
        {
            //Armo lista de elemtnos que SI reemplazo.
            //Busco en la coleccion, todas las lineas en el periodo, y con el aviso seleccionado.
            var lineasACopiar = lineas.FindAll(
                (x) =>
                    (x.Fecha >= fOrigenDesde
                    && x.Fecha <= fOrigenHasta));


            //Si encontre líneas a copiar...
            if (lineasACopiar.Count > 0)
            {
                EstimadoDetDTO nuevaLinea;
                DateTime fechaTmp;
                List<EstimadoDetDTO> lineasTmp = new List<EstimadoDetDTO>();
                TimeSpan diasEnElFuturo = fDestinoDesde.Subtract(fOrigenDesde);

                //Por cada linea que encontre, genero una nueva e igual, x dias en el futuro.
                foreach (var linea in lineasACopiar)
                {
                    nuevaLinea           = new EstimadoDetDTO();
                    nuevaLinea.RecId     = NextTempRecId(viewState);
                    nuevaLinea.DatareaId = linea.DatareaId;

                    //Avanzo la fecha tantos dias como corresponda...
                    fechaTmp                 = linea.Fecha.Add(diasEnElFuturo)  ;
                    nuevaLinea.Fecha         = fechaTmp                         ; 
                    nuevaLinea.Dia           = fechaTmp.Day                     ; 
                    nuevaLinea.DiaSemana     = fechaTmp.ToString("dddd", new CultureInfo("es-ES")).ToUpper().Trim(); 
                    nuevaLinea.Costo         = linea.Costo                      ;
                    nuevaLinea.CostoOp       = linea.CostoOp                    ;
                    nuevaLinea.CostoOpUni    = linea.CostoOpUni                 ;
                    nuevaLinea.CostoUni      = linea.CostoUni                   ;
                    nuevaLinea.Duracion      = linea.Duracion                   ;
                    nuevaLinea.Hora          = linea.Hora                       ;
                    nuevaLinea.IdentifAviso  = linea.IdentifAviso               ;
                    nuevaLinea.PautaId       = linea.PautaId                    ;
                    nuevaLinea.Salida        = linea.Salida                     ;

                    //Agrego la nueva linea.
                    lineasTmp.Add(nuevaLinea);
                }

                //Junto las dos listas (temporal y la que ya tenia).
                lineasTmp.AddRange(lineas);

                //Ordeno por fecha.
                lineasTmp.Sort(
                    (x, y) => DateTime.Compare(x.Fecha, y.Fecha));

                //Guardo la lista en el Viewstate.
                return lineasTmp;
            }
            else
            {
                return lineas;
            }
        }

        internal static List<EstimadoDetDTO> ReemplazarAvisosTodos(string avisoOrigen, string avisoDestino, List<EstimadoDetDTO> lineas)
        {
            //Por cada linea cuto identifAviso sea igual al 'avisoOrigen', reemplazo el aviso.
            lineas.ForEach(linea =>
            {
                if (linea.IdentifAviso == avisoOrigen)
                    linea.IdentifAviso = avisoDestino;
            });

            return lineas;
        }

        internal static List<CertificadoDetDTO> ReemplazarAvisosTodos(string avisoOrigen, string avisoDestino, List<CertificadoDetDTO> lineas)
        {
            //Por cada linea cuto identifAviso sea igual al 'avisoOrigen', reemplazo el aviso.
            lineas.ForEach(linea =>
            {
                if (linea.IdentifAviso == avisoOrigen)
                    linea.IdentifAviso = avisoDestino;
            });

            return lineas;
        }

        internal static List<EstimadoDetDTO> ReemplazarAvisosSeleccionados(string avisoOrigen, string avisoDestino, List<object> ids, List<EstimadoDetDTO> lineas)
        {
            //Obtengo todos los Ids de los registros seleccionados.
            //List<object> ids = gv.GetSelectedFieldValues("RecId");

            //Busco las línas con los RecId seleccionados, y cuyo aviso, sea igual al 'avisoOrigen',
            //Para cada una de las líneas encontradas, reemplazo el aviso.
            lineas.FindAll(
                (x) => ids.Contains(x.RecId) && x.IdentifAviso == avisoOrigen).ForEach(
                (linea) =>
                    linea.IdentifAviso = avisoDestino);

            return lineas;
        }

        internal static List<CertificadoDetDTO> ReemplazarAvisosSeleccionados(string avisoOrigen, string avisoDestino, List<object> ids, List<CertificadoDetDTO> lineas)
        {
            //Obtengo todos los Ids de los registros seleccionados.
            //List<object> ids = gv.GetSelectedFieldValues("RecId");

            //Busco las línas con los RecId seleccionados, y cuyo aviso, sea igual al 'avisoOrigen',
            //Para cada una de las líneas encontradas, reemplazo el aviso.
            lineas.FindAll(
                (x) => ids.Contains(x.RecId) && x.IdentifAviso == avisoOrigen).ForEach(
                (linea) =>
                    linea.IdentifAviso = avisoDestino);

            return lineas;
        }

        internal static List<CertificadoDetDTO> ReemplazarAvisosPorPeriodo(DateTime fDesde, DateTime fHasta, string avisoOrigen, string avisoDestino, List<CertificadoDetDTO> lineas)
        {
            //Busco las línas con el periodo seleccionado, y cuyo aviso, sea igual al 'avisoOrigen',
            //Para cada una de las líneas encontradas, reemplazo el aviso.
            lineas.FindAll(
                (x) =>
                    (x.Fecha >= fDesde
                    && x.Fecha <= fHasta
                    && x.IdentifAviso == avisoOrigen)).ForEach(
                    (linea) =>
                        linea.IdentifAviso = avisoDestino);

            return lineas;
        }

        internal static List<EstimadoDetDTO> ReemplazarAvisosPorPeriodo(DateTime fDesde, DateTime fHasta, string avisoOrigen, string avisoDestino, List<EstimadoDetDTO> lineas)
        {
            //Busco las línas con el periodo seleccionado, y cuyo aviso, sea igual al 'avisoOrigen',
            //Para cada una de las líneas encontradas, reemplazo el aviso.
            lineas.FindAll(
                (x) =>
                    (x.Fecha >= fDesde
                    && x.Fecha <= fHasta
                    && x.IdentifAviso == avisoOrigen)).ForEach(
                    (linea) =>
                        linea.IdentifAviso = avisoDestino);

            return lineas;
        }

    }
}