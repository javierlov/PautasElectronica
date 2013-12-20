using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PautasPublicidad.Web.Classes;
using PautasPublicidad.Business;
using PautasPublicidad.DTO;
using System.Data;
using System.Drawing;
using ClosedXML.Excel;
using System.IO;

namespace PautasPublicidad.Web.Classes
{
    public class csOP_Helper
    {
        public static List<OrdenadoCabDTO> ordenado;
        public static List<EstimadoCabDTO> estimado;
        public static List<CertificadoCabDTO> certificados;

        static string _Estado          = string.Empty;
        static string _Origen          = string.Empty;
        static string _PautaId         = string.Empty;

        static object _Cabecera        = null;
        static object _Detalle         = null;
        static object _SKUS            = null;

        static EspacioContDTO _Espacio = null;

        //Constructor

        public csOP_Helper(string Estado, string Origen, string PautaId, OrdenadoCabDTO Cabecera, List<OrdenadoDetDTO> Detalles, List<OrdenadoSKUDTO> SKUS, EspacioContDTO Espacio, string filename)
        {
            _Estado   = Estado;
            _Origen   = Origen;
            _PautaId  = PautaId;
            _Cabecera = Cabecera;
            _Detalle  = Detalles;
            _SKUS     = SKUS;
            _Espacio  = Espacio;

            Imprimir(Espacio.FormatoOP, filename);
            ArmarCabecera(Espacio.FormatoOP, filename);
            ArmarDetalle(Espacio.FormatoOP, filename);
        }

        public csOP_Helper(string Estado, string Origen, string PautaId, EstimadoCabDTO Cabecera, List<EstimadoDetDTO> Detalles, List<EstimadoSKUDTO> SKUS, EspacioContDTO Espacio,string filename)
        {
            _Estado   = Estado;
            _Origen   = Origen;
            _PautaId  = PautaId;
            _Cabecera = Cabecera;
            _Detalle  = Detalles;
            _SKUS     = SKUS;
            _Espacio  = Espacio;

            Imprimir(Espacio.FormatoOP, filename);
            ArmarCabecera(Espacio.FormatoOP, filename);
            ArmarDetalle(Espacio.FormatoOP, filename);
        }

        public csOP_Helper(string Estado, string Origen, string PautaId, CertificadoCabDTO Cabecera, List<CertificadoDetDTO> Detalles, List<CertificadoSKUDTO> SKUS, EspacioContDTO Espacio,string filename)
        {
            _Estado   = Estado;
            _Origen   = Origen;
            _PautaId  = PautaId;
            _Cabecera = Cabecera;
            _Detalle  = Detalles;
            _SKUS     = SKUS;
            _Espacio  = Espacio;

            Imprimir(Espacio.FormatoOP, filename);
            ArmarCabecera(Espacio.FormatoOP, filename);
            ArmarDetalle(Espacio.FormatoOP, filename);
        }

        public void Imprimir(string TipoOP,string nomArchivo)
        {
            try {

                switch (TipoOP)
                {
                    case "CALENDARIO_DESCRIPTIVO":
                        {
                            switch (_Estado.ToUpper())
                            {
                                case "ORDENADO":
                                    {
                                        csOP_CALENDARIO_DESCRIPTIVO miOpCalendarioDescriptivo = new csOP_CALENDARIO_DESCRIPTIVO(_Estado, _Origen, _PautaId, (OrdenadoCabDTO)_Cabecera, (List<OrdenadoDetDTO>)_Detalle, (List<OrdenadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpCalendarioDescriptivo.CargarCabecera(_Estado);
                                        miOpCalendarioDescriptivo.CargarItems(_Estado);
                                        miOpCalendarioDescriptivo.CargarSKUS(_Estado);
                                        miOpCalendarioDescriptivo.CargarPie();

                                        miOpCalendarioDescriptivo.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "ESTIMADO":
                                    {
                                        csOP_CALENDARIO_DESCRIPTIVO miOpCalendarioDescriptivo = new csOP_CALENDARIO_DESCRIPTIVO(_Estado, _Origen, _PautaId, (EstimadoCabDTO)_Cabecera, (List<EstimadoDetDTO>)_Detalle, (List<EstimadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpCalendarioDescriptivo.CargarCabecera(_Estado);
                                        miOpCalendarioDescriptivo.CargarItems(_Estado);
                                        miOpCalendarioDescriptivo.CargarSKUS(_Estado);
                                        miOpCalendarioDescriptivo.CargarPie();

                                        miOpCalendarioDescriptivo.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "CERTIFICADO":
                                    {
                                        csOP_CALENDARIO_DESCRIPTIVO miOpCalendarioDescriptivo = new csOP_CALENDARIO_DESCRIPTIVO(_Estado, _Origen, _PautaId, (CertificadoCabDTO)_Cabecera, (List<CertificadoDetDTO>)_Detalle, (List<CertificadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpCalendarioDescriptivo.CargarCabecera(_Estado);
                                        miOpCalendarioDescriptivo.CargarItems(_Estado);
                                        miOpCalendarioDescriptivo.CargarSKUS(_Estado);
                                        miOpCalendarioDescriptivo.CargarPie();

                                        miOpCalendarioDescriptivo.CreatePackage(nomArchivo);

                                        break;
                                    }
                            }

                            break;
                        }

                    case "CALENDARIO_NUMERICO":
                        {
                            switch (_Estado.ToUpper())
                            {
                                case "ORDENADO":
                                    {
                                        csOP_CALENDARIO_NUMERICO miOpCalendarioNumerico = new csOP_CALENDARIO_NUMERICO(_Estado, _Origen, _PautaId, (OrdenadoCabDTO)_Cabecera, (List<OrdenadoDetDTO>)_Detalle, (List<OrdenadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpCalendarioNumerico.CargarCabecera(_Estado);
                                        miOpCalendarioNumerico.CargarItems(_Estado);
                                        miOpCalendarioNumerico.CargarSKUS(_Estado);
                                        miOpCalendarioNumerico.CargarPie();

                                        miOpCalendarioNumerico.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "ESTIMADO":
                                    {
                                        csOP_CALENDARIO_NUMERICO miOpCalendarioNumerico = new csOP_CALENDARIO_NUMERICO(_Estado, _Origen, _PautaId, (EstimadoCabDTO)_Cabecera, (List<EstimadoDetDTO>)_Detalle, (List<EstimadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpCalendarioNumerico.CargarCabecera(_Estado);
                                        miOpCalendarioNumerico.CargarItems(_Estado);
                                        miOpCalendarioNumerico.CargarSKUS(_Estado);
                                        miOpCalendarioNumerico.CargarPie();

                                        miOpCalendarioNumerico.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "CERTIFICADO":
                                    {
                                        csOP_CALENDARIO_NUMERICO miOpCalendarioNumerico = new csOP_CALENDARIO_NUMERICO(_Estado, _Origen, _PautaId, (CertificadoCabDTO)_Cabecera, (List<CertificadoDetDTO>)_Detalle, (List<CertificadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpCalendarioNumerico.CargarCabecera(_Estado);
                                        miOpCalendarioNumerico.CargarItems(_Estado);
                                        miOpCalendarioNumerico.CargarSKUS(_Estado);
                                        miOpCalendarioNumerico.CargarPie();

                                        miOpCalendarioNumerico.CreatePackage(nomArchivo);

                                        break;
                                    }
                            }
                            break;
                        }

                    case "GRAFICA":
                        {
                            switch (_Estado.ToUpper())
                            {
                                case "ORDENADO":
                                    {
                                        csOP_GRAFICA miOpGrafica = new csOP_GRAFICA(_Estado, _Origen, _PautaId, (OrdenadoCabDTO)_Cabecera, (List<OrdenadoDetDTO>)_Detalle, (List<OrdenadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpGrafica.CargarCabecera(_Estado);
                                        miOpGrafica.CargarItems(_Estado);
                                        miOpGrafica.CargarSKUS(_Estado);
                                        miOpGrafica.CargarPie();

                                        miOpGrafica.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "ESTIMADO":
                                    {
                                        csOP_GRAFICA miOpGrafica = new csOP_GRAFICA(_Estado, _Origen, _PautaId, (EstimadoCabDTO)_Cabecera, (List<EstimadoDetDTO>)_Detalle, (List<EstimadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpGrafica.CargarCabecera(_Estado);
                                        miOpGrafica.CargarItems(_Estado);
                                        miOpGrafica.CargarSKUS(_Estado);
                                        miOpGrafica.CargarPie();

                                        miOpGrafica.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "CERTIFICADO":
                                    {
                                        csOP_GRAFICA miOpGrafica = new csOP_GRAFICA(_Estado, _Origen, _PautaId, (CertificadoCabDTO)_Cabecera, (List<CertificadoDetDTO>)_Detalle, (List<CertificadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpGrafica.CargarCabecera(_Estado);
                                        miOpGrafica.CargarItems(_Estado);
                                        miOpGrafica.CargarSKUS(_Estado);
                                        miOpGrafica.CargarPie();

                                        miOpGrafica.CreatePackage(nomArchivo);

                                        break;
                                    }
                            }
                            break;
                        }

                    case "PNT_PRODUCTO":
                        {
                            switch (_Estado.ToUpper())
                            {
                                case "ORDENADO":
                                    {
                                        csOP_PNT_PRODUCTO miOpPntProducto = new csOP_PNT_PRODUCTO(_Estado, _Origen, _PautaId, (OrdenadoCabDTO)_Cabecera, (List<OrdenadoDetDTO>)_Detalle, (List<OrdenadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpPntProducto.CargarCabecera(_Estado);
                                        miOpPntProducto.CargarItems(_Estado);
                                        miOpPntProducto.CargarSKUS(_Estado);
                                        miOpPntProducto.CargarPie();

                                        miOpPntProducto.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "ESTIMADO":
                                    {
                                        csOP_PNT_PRODUCTO miOpPntProducto = new csOP_PNT_PRODUCTO(_Estado, _Origen, _PautaId, (EstimadoCabDTO)_Cabecera, (List<EstimadoDetDTO>)_Detalle, (List<EstimadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpPntProducto.CargarCabecera(_Estado);
                                        miOpPntProducto.CargarItems(_Estado);
                                        miOpPntProducto.CargarSKUS(_Estado);
                                        miOpPntProducto.CargarPie();

                                        miOpPntProducto.CreatePackage(nomArchivo);

                                        break;
                                    }

                                case "CERTIFICADO":
                                    {
                                        csOP_PNT_PRODUCTO miOpPntProducto = new csOP_PNT_PRODUCTO(_Estado, _Origen, _PautaId, (CertificadoCabDTO)_Cabecera, (List<CertificadoDetDTO>)_Detalle, (List<CertificadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                        miOpPntProducto.CargarCabecera(_Estado);
                                        miOpPntProducto.CargarItems(_Estado);
                                        miOpPntProducto.CargarSKUS(_Estado);
                                        miOpPntProducto.CargarPie();

                                        miOpPntProducto.CreatePackage(nomArchivo);

                                        break;
                                    }
                            }
                            break;
                        }

                    case "PNT_SALIDA":
                        {
                            switch (_Estado.ToUpper())
                            {
                                case "ORDENADO": {

                                    csOP_PNT_SALIDA miOpPntSalida = new csOP_PNT_SALIDA(_Estado, _Origen, _PautaId, (OrdenadoCabDTO)_Cabecera, (List<OrdenadoDetDTO>)_Detalle, (List<OrdenadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio, nomArchivo);

                                    miOpPntSalida.CargarCabecera(_Estado);
                                    miOpPntSalida.CargarItems(_Estado);
                                    miOpPntSalida.CargarSKUS(_Estado);
                                    miOpPntSalida.CargarPie();

                                    miOpPntSalida.CreatePackage(nomArchivo);

                                    break; }

                                case "ESTIMADO": {

                                    csOP_PNT_SALIDA miOpPntSalida = new csOP_PNT_SALIDA(_Estado, _Origen, _PautaId, (EstimadoCabDTO)_Cabecera, (List<EstimadoDetDTO>)_Detalle, (List<EstimadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                    miOpPntSalida.CargarCabecera(_Estado);
                                    miOpPntSalida.CargarItems(_Estado);
                                    miOpPntSalida.CargarSKUS(_Estado);
                                    miOpPntSalida.CargarPie();

                                    miOpPntSalida.CreatePackage(nomArchivo);

                                    break; }

                                case "CERTIFICADO": {

                                    csOP_PNT_SALIDA miOpPntSalida = new csOP_PNT_SALIDA(_Estado, _Origen, _PautaId, (CertificadoCabDTO)_Cabecera, (List<CertificadoDetDTO>)_Detalle, (List<CertificadoSKUDTO>)_SKUS, (EspacioContDTO)_Espacio);

                                    miOpPntSalida.CargarCabecera(_Estado);
                                    miOpPntSalida.CargarItems(_Estado);
                                    miOpPntSalida.CargarSKUS(_Estado);
                                    miOpPntSalida.CargarPie();

                                    miOpPntSalida.CreatePackage(nomArchivo);
                                    
                                    break; }
                            }
                            break;
                        }
                }

            }
            catch (Exception ex) {

                string a = ex.Message;
            
            }
            
        }

        public void ArmarCabecera(string TipoOP,string nomArchivo)
        {
            
            switch (_Estado.ToUpper())
            {
                case "ORDENADO":    {
                    switch (TipoOP.ToUpper())
                    {
                        case "CALENDARIO_DESCRIPTIVO": {
                            Cabecera_OP_Calendario_Descriptivo(nomArchivo);
                            break; }

                        case "CALENDARIO_NUMERICO": {
                            Cabecera_OP_Calendario_Numerico(nomArchivo);
                            break; }
                        
                        case "GRAFICA": {
                            Cabecera_OP_Grafica(nomArchivo);
                            break; }
                        
                        case "PNT_PRODUCTO": {
                            Cabecera_OP_PNT_Producto(nomArchivo);
                            break; }
                        
                        case "PNT_SALIDA": {
                            Cabecera_OP_PNT_Salida(nomArchivo);
                            break; }
                    }
                    break; }
                case "ESTIMADO":    {
                    switch (TipoOP.ToUpper())
                    {
                        case "CALENDARIO_DESCRIPTIVO": {
                            Cabecera_OP_Calendario_Descriptivo(nomArchivo);
                            break; }

                        case "CALENDARIO_NUMERICO": {
                            Cabecera_OP_Calendario_Numerico(nomArchivo);
                            break; }
                        
                        case "GRAFICA": {
                            Cabecera_OP_Grafica(nomArchivo);
                            break; }
                        
                        case "PNT_PRODUCTO": {
                            Cabecera_OP_PNT_Producto(nomArchivo);
                            break; }
                        
                        case "PNT_SALIDA": {
                            Cabecera_OP_PNT_Salida(nomArchivo);
                            break; }
                    }
                    break; }

                case "CERTIFICADO":
                    {
                        switch (TipoOP.ToUpper())
                        {
                            case "CALENDARIO_DESCRIPTIVO":
                                {
                                    Cabecera_OP_Calendario_Descriptivo(nomArchivo);
                                    break;
                                }
                            case "CALENDARIO_NUMERICO":
                                {
                                    Cabecera_OP_Calendario_Numerico(nomArchivo);
                                    break;
                                }
                            case "GRAFICA":
                                {
                                    Cabecera_OP_Grafica(nomArchivo);
                                    break;
                                }
                            case "PNT_PRODUCTO":
                                {
                                    Cabecera_OP_PNT_Producto(nomArchivo);
                                    break;
                                }
                            case "PNT_SALIDA":
                                {
                                    Cabecera_OP_PNT_Salida(nomArchivo);
                                    break;
                                }
                        }
                        break;
                    }
            }       
        }

        public void ArmarDetalle(string TipoOP,string nomArchivo)
        {

            switch (_Estado.ToUpper())
            {
                case "ORDENADO":
                    {
                        switch (TipoOP.ToUpper())
                        {
                            case "CALENDARIO_DESCRIPTIVO":
                                {
                                    Detalle_OP_Calendario_Descriptivo(nomArchivo);
                                    break;
                                }
                            case "CALENDARIO_NUMERICO":
                                {
                                    Detalle_OP_Calendario_Numerico(nomArchivo);
                                    break;
                                }
                            case "GRAFICA":
                                {
                                    Detalle_OP_Grafica(nomArchivo);
                                    break;
                                }
                            case "PNT_PRODUCTO":
                                {
                                    Detalle_OP_PNT_Producto(nomArchivo);
                                    break;
                                }
                            case "PNT_SALIDA":
                                {
                                    Detalle_OP_PNT_Salida(nomArchivo);
                                    break;
                                }
                        }
                        break;
                    }
                case "ESTIMADO":
                    {
                        switch (TipoOP.ToUpper())
                        {
                            case "CALENDARIO_DESCRIPTIVO":
                                {
                                    Detalle_OP_Calendario_Descriptivo(nomArchivo);
                                    break;
                                }
                            case "CALENDARIO_NUMERICO":
                                {
                                    Detalle_OP_Calendario_Numerico(nomArchivo);
                                    break;
                                }
                            case "GRAFICA":
                                {
                                    Detalle_OP_Grafica(nomArchivo);
                                    break;
                                }
                            case "PNT_PRODUCTO":
                                {
                                    Detalle_OP_PNT_Producto(nomArchivo);
                                    break;
                                }
                            case "PNT_SALIDA":
                                {
                                    Detalle_OP_PNT_Salida(nomArchivo);
                                    break;
                                }
                        }
                        break;
                    }

                case "CERTIFICADO":
                    {
                        switch (TipoOP.ToUpper())
                        {
                            case "CALENDARIO_DESCRIPTIVO":
                                {
                                    Detalle_OP_Calendario_Descriptivo(nomArchivo);
                                    break;
                                }
                            case "CALENDARIO_NUMERICO":
                                {
                                    Detalle_OP_Calendario_Numerico(nomArchivo);
                                    break;
                                }
                            case "GRAFICA":
                                {
                                    Detalle_OP_Grafica(nomArchivo);
                                    break;
                                }
                            case "PNT_PRODUCTO":
                                {
                                    Detalle_OP_PNT_Producto(nomArchivo);
                                    break;
                                }
                            case "PNT_SALIDA":
                                {
                                    Detalle_OP_PNT_Salida(nomArchivo);
                                    break;
                                }
                        }
                        break;
                    }
            }       
        }

        public void ArmarPie()
        { }
       
#region Cabeceras

        public void Cabecera_OP_Calendario_Descriptivo(string nomarchivo)
        {

            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO": {

                    EmpresaDTO empresa        = new DAO.EmpresaDAO().Read(1);
                    OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;

                    DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));
                    string[] Meses         = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string aniomes         = miCabecera.AnoMes.ToString();


                    //ARMADO DE CALENDARIO - MEJORAR
                    int mes       = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio      = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes    = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);

                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                    for (int Filas = 20; Filas <= 40; Filas = Filas + 5)
                    {
                        for (int Columnas = 2; Columnas <= 8; Columnas++)
                        {
                            DateTime theDate = new DateTime(anio, 1, 1).AddDays(dia-2);

                            oSheet.Cell(Filas, Columnas).Value = theDate;

                            dia++;
                        }
                    }

                    oSheet.Cells("C5").Value  = "Orden de Publicidad - " + miCabecera.IdentifEspacio;
                    oSheet.Cells("B9").Value  = miCabecera.PautaId;
                    oSheet.Cells("B10").Value = medio.Name;
                    oSheet.Cells("B11").Value = _Espacio.Name;
                    oSheet.Cells("B12").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("B13").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cells("A14").Value = "EMAIL: ";
                    oSheet.Cells("B14").Value = _Espacio.Email == null ? "" : _Espacio.Email;
                    FrecuenciaDTO frecuencia  = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", miCabecera.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                    oSheet.Cells("B15").Value = frecuencia.Name + " de " + _Espacio.HoraInicio + " a " + _Espacio.HoraFin + "hs.";

                    break; }
                case "ESTIMADO": {

                    EmpresaDTO empresa        = new DAO.EmpresaDAO().Read(1);
                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;

                    DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));
                    string[] Meses         = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string aniomes         = miCabecera.AnoMes.ToString();

                    //ARMADO DE CALENDARIO - MEJORAR
                    int mes       = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio      = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes    = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);

                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                    for (int Filas = 20; Filas <= 40; Filas = Filas + 5)
                    {
                        for (int Columnas = 2; Columnas <= 8; Columnas++)
                        {
                            DateTime theDate = new DateTime(anio, 1, 1).AddDays(dia - 2);

                            oSheet.Cell(Filas, Columnas).Value = theDate;

                            dia++;
                        }
                    }

                    oSheet.Cells("C5").Value  = "Orden de Publicidad - " + miCabecera.IdentifEspacio;
                    oSheet.Cells("B9").Value  = miCabecera.PautaId;
                    oSheet.Cells("B10").Value = medio.Name;
                    oSheet.Cells("B11").Value = _Espacio.Name;
                    oSheet.Cells("B12").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("B13").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cells("A14").Value = "EMAIL: ";
                    oSheet.Cells("B14").Value = _Espacio.Email == null ? "" : _Espacio.Email;
                    FrecuenciaDTO frecuencia  = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", miCabecera.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                    oSheet.Cells("B15").Value = frecuencia.Name + " de " + _Espacio.HoraInicio + " a " + _Espacio.HoraFin + "hs.";

                    break; }

                case "CERTIFICADO": {

                    EmpresaDTO empresa           = new DAO.EmpresaDAO().Read(1);
                    CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;

                    DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));
                    string[] Meses         = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string aniomes         = miCabecera.AnoMes.ToString();


                    //ARMADO DE CALENDARIO - MEJORAR
                    int mes       = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio      = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);

                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                    for (int Filas = 20; Filas <= 40; Filas = Filas + 5)
                    {
                        for (int Columnas = 2; Columnas <= 8; Columnas++)
                        {

                            DateTime theDate = new DateTime(anio, 1, 1).AddDays(dia - 2);

                            oSheet.Cell(Filas, Columnas).Value = theDate;

                            dia++;
                        }
                    }

                    oSheet.Cells("C5").Value = "Orden de Publicidad - " + miCabecera.IdentifEspacio;
                    oSheet.Cells("B9").Value = miCabecera.PautaId;
                    oSheet.Cells("B10").Value = medio.Name;
                    oSheet.Cells("B11").Value = _Espacio.Name;
                    oSheet.Cells("B12").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("B13").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cells("A14").Value = "EMAIL: ";
                    oSheet.Cells("B14").Value = _Espacio.Email == null ? "" : _Espacio.Email;
                    FrecuenciaDTO frecuencia = CRUDHelper.Read(string.Format("IdentifFrecuencia = '{0}'", miCabecera.IdentifFrecuencia), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.Frecuencia));
                    oSheet.Cells("B15").Value = frecuencia.Name + " de " + _Espacio.HoraInicio + " a " + _Espacio.HoraFin + "hs.";
                    
                    break; }
            }

            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 

        }

        public void Cabecera_OP_Calendario_Numerico(string nomarchivo)
        {
            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);


            string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO":
                    {
                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                        string aniomes = miCabecera.AnoMes.ToString();
                        oSheet.Cells("C8").Value = miCabecera.PautaId;
                        oSheet.Cells("C11").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);

                        break;
                    }
                case "ESTIMADO": {

                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                    string aniomes = miCabecera.AnoMes.ToString();
                    oSheet.Cells("C8").Value = miCabecera.PautaId;
                    oSheet.Cells("C11").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);

                    break;
                }
                case "CERTIFICADO": {

                    CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                    string aniomes = miCabecera.AnoMes.ToString();
                    oSheet.Cells("C8").Value = miCabecera.PautaId;
                    oSheet.Cells("C11").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);

                    break; }
            }

            DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

            oSheet.Cells("C9").Value = medio.Name;
            oSheet.Cells("C10").Value = _Espacio.Name;
            oSheet.Cells("C12").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
            oSheet.Cells("C13").Value = _Espacio.Telefono == null ? "" : "'" + _Espacio.Telefono;
            oSheet.Cells("C14").Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion;
            oSheet.Cells("c15").Value = _Espacio.Email == null ? "" : _Espacio.Email;

            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 
        
        }

        public void Cabecera_OP_Grafica(string nomarchivo)
        {
            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO":
                    {
                        EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string aniomes = miCabecera.AnoMes.ToString();

                        oSheet.Cells("C6").Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                        oSheet.Cells("C7").Value = medio.Name;
                        oSheet.Cells("C8").Value = _Espacio.Name;
                        oSheet.Cells("C9").Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                        oSheet.Cells("C10").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                        oSheet.Cells("D14").Value = "MES: " + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                        oSheet.Cells("F25").Value = _Espacio.Responsable == null ? "" : _Espacio.Responsable.ToUpper();

                        SetUpDTO setup = new SetUpDTO();
                        string sector = setup.Sector;

                        oSheet.Cells("F26").Value = sector;
                        oSheet.Cells("F27").Value = empresa.Name;
                        oSheet.Cells("F28").Value = empresa.Leyenda;

                        break;
                    }
            
                case "ESTIMADO": {

                    EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                    DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string aniomes = miCabecera.AnoMes.ToString();

                    oSheet.Cells("C6").Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("C7").Value = medio.Name;
                    oSheet.Cells("C8").Value = _Espacio.Name;
                    oSheet.Cells("C9").Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                    oSheet.Cells("C10").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cells("D14").Value = "MES: " + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("F25").Value = _Espacio.Responsable == null ? "" : _Espacio.Responsable.ToUpper();

                    SetUpDTO setup = new SetUpDTO();
                    string sector = setup.Sector;

                    oSheet.Cells("F26").Value = sector;
                    oSheet.Cells("F27").Value = empresa.Name;
                    oSheet.Cells("F28").Value = empresa.Leyenda;
                    
                    break; }

                case "CERTIFICADO":
                    {
                        EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                        CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string aniomes = miCabecera.AnoMes.ToString();

                        oSheet.Cells("C6").Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                        oSheet.Cells("C7").Value = medio.Name;
                        oSheet.Cells("C8").Value = _Espacio.Name;
                        oSheet.Cells("C9").Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                        oSheet.Cells("C10").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                        oSheet.Cells("D14").Value = "MES: " + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                        oSheet.Cells("F25").Value = _Espacio.Responsable == null ? "" : _Espacio.Responsable.ToUpper();

                        SetUpDTO setup = new SetUpDTO();
                        string sector = setup.Sector;

                        oSheet.Cells("F26").Value = sector;
                        oSheet.Cells("F27").Value = empresa.Name;
                        oSheet.Cells("F28").Value = empresa.Leyenda;

                        break;
                    }
            }

            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 
        }

        public void Cabecera_OP_PNT_Salida(string nomarchivo)
        {
            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO":
                    {
                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;                        
                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string aniomes = miCabecera.AnoMes.ToString();

                        //formato de referencia es [fila, columna]
                        oSheet.Cells("B10").Value = medio.Name;
                        oSheet.Cells("B11").Value = _Espacio.Name;
                        oSheet.Cells("B12").Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                        oSheet.Cells("B13").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                        oSheet.Cells("B14").Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                        oSheet.Cells("B15").Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion.ToUpper();
                        oSheet.Cells("B16").Value = _Espacio.Email == null ? "" : _Espacio.Email;

                        oSheet.Cells("D5").Value = "ORDEN DE PUBLICIDAD - " + _Espacio.Name;

                        //ARMADO DE CALENDARIO - MEJORAR
                        int mes = Convert.ToInt32(aniomes.Substring(4,2));
                        int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                        int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                        string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                        int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0,4) + "-" + aniomes.Substring(4,2) + "-"  + "01").DayOfWeek);
                        int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);


                        int dia = PrimerDiaMes - PrimerDiaSemana -1;

                        for (int Filas = 20; Filas <= 40; Filas = Filas + 5)
                        {
                            for (int Columnas = 3; Columnas <= 9; Columnas++)
                            {

                                DateTime theDate = new DateTime(anio, 1, 1).AddDays(dia +1);

                                oSheet.Cell(Filas, Columnas).Value = theDate;

                                dia++;
                            }
                        }

                        ////////////////
                        break;
                    }

                case "ESTIMADO": {

                    //formato de referencia es [fila, columna]
                    oSheet.Cells("D5").Value = "ORDEN DE PUBLICIDAD - " + _Espacio.Name;
                    oSheet.Cells("B2").Value = _Espacio.IdentifMedio;
                    oSheet.Cells("B11").Value = _Espacio.Name;

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;

                    string aniomes = miCabecera.AnoMes.ToString();

                    oSheet.Cells("B12").Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("B13").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cells("B14").Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                    oSheet.Cells("B15").Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion.ToUpper();
                    oSheet.Cells("B16").Value = _Espacio.Email == null ? "" : _Espacio.Email;

                    //ARMADO DE CALENDARIO - MEJORAR
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);


                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                    for (int Filas = 20; Filas <= 40; Filas = Filas + 5)
                    {
                        for (int Columnas = 3; Columnas <= 9; Columnas++)
                        {

                            oSheet.Cell(Filas, Columnas).Value = dia;

                            dia++;
                        }
                    }

                    ////////////////
                    break; }

                case "CERTIFICADO": {

                    //formato de referencia es [fila, columna]
                    oSheet.Cells("D5").Value = "ORDEN DE PUBLICIDAD - " + _Espacio.Name;
                    oSheet.Cells("B2").Value = _Espacio.IdentifMedio;
                    oSheet.Cells("B11").Value = _Espacio.Name;

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

                    CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;

                    string aniomes = miCabecera.AnoMes.ToString();

                    oSheet.Cells("B12").Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cells("B13").Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cells("B14").Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                    oSheet.Cells("B15").Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion.ToUpper();
                    oSheet.Cells("B16").Value = _Espacio.Email == null ? "" : _Espacio.Email;

                    //ARMADO DE CALENDARIO - MEJORAR
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);


                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                    for (int Filas = 20; Filas <= 40; Filas = Filas + 5)
                    {
                        for (int Columnas = 3; Columnas <= 9; Columnas++)
                        {

                            oSheet.Cell(Filas, Columnas).Value = dia;

                            dia++;
                        }
                    }

                    ////////////////
                    break;
                }
            }

            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 

        }

        public void Cabecera_OP_PNT_Producto(string nomarchivo)
        {

            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO":
                    {
                        EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string aniomes = miCabecera.AnoMes.ToString();

                        
                        oSheet.Cell(6, 6).Value = "'" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                        oSheet.Cell(7, 6).Value = "'" + miCabecera.PautaId.ToString();
                        oSheet.Cell(10, 3).Value = medio.Name;
                        oSheet.Cell(11, 3).Value = _Espacio.Name;
                        oSheet.Cell(13, 3).Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                        oSheet.Cell(14, 3).Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                        oSheet.Cell(15, 3).Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion.ToUpper();
                        oSheet.Cell(16, 3).Value = _Espacio.Email == null ? "" : _Espacio.Email;

                        break;
                    }
                case "ESTIMADO": 
                    {

                    EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                    DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string aniomes = miCabecera.AnoMes.ToString();

                    oSheet.Cell(6, 6).Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cell(10, 3).Value = medio.Name;
                    oSheet.Cell(11, 3).Value = _Espacio.Name;
                    oSheet.Cell(13, 3).Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cell(14, 3).Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                    oSheet.Cell(15, 3).Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion.ToUpper();
                    oSheet.Cell(16, 3).Value = _Espacio.Email == null ? "" : _Espacio.Email;
                    
                    break; 
                    }
                case "CERTIFICADO": 
                    {

                    EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                    CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                    DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string aniomes = miCabecera.AnoMes.ToString();

                    oSheet.Cell(6, 6).Value = "" + Meses[Convert.ToInt32(aniomes.Substring(4, 2))].ToUpper() + "/" + aniomes.Substring(0, 4);
                    oSheet.Cell(10, 3).Value = medio.Name;
                    oSheet.Cell(11, 3).Value = _Espacio.Name;
                    oSheet.Cell(13, 3).Value = _Espacio.Contacto == null ? "" : _Espacio.Contacto.ToUpper();
                    oSheet.Cell(14, 3).Value = _Espacio.Telefono == null ? "" : _Espacio.Telefono;
                    oSheet.Cell(15, 3).Value = _Espacio.Direccion == null ? "" : _Espacio.Direccion.ToUpper();
                    oSheet.Cell(16, 3).Value = _Espacio.Email == null ? "" : _Espacio.Email;
                    
                    break; 
                    }
            }

            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 

        }

#endregion

#region Detalles

        public void Detalle_OP_Calendario_Descriptivo(string nomArchivo)
        {
            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomArchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                #region Ordenado
                case "ORDENADO":
                    {

                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                        List<OrdenadoDetDTO> miDetalle = (List<OrdenadoDetDTO>)_Detalle;

                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };

                        AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[0].IdentifAviso));

                        var dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten);

                        AvisosIdAtenDTO ident =(AvisosIdAtenDTO)CRUDHelper.Read((string.Format("IdentifAviso = '{0}'", miDetalle[0].IdentifAviso)), dao);

                        dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion);

                        IdentAtencionDTO entity = (DTO.IdentAtencionDTO)CRUDHelper.Read(ident.RecId, dao);

                        if(entity != null)
                        {
                            oSheet.Cells("D16").Value = "Teléfonos " + entity.Telefono;
                        }

                        string aniomes = miCabecera.AnoMes.ToString();

                        List<string> IdentifAvisos = new List<string>();

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {

                            if (IdentifAvisos.Count == 0)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                            }

                            bool lencontrado = false;

                            string scadena = miDetalle[i].IdentifAviso;

                            for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                            {
                                if (IdentifAvisos[j] == scadena)
                                {
                                    lencontrado = true;

                                    break;
                                }
                            }

                            if (!lencontrado)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                                break;
                            }

                        }

                        List<string> Horarios = new List<string>();

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            int cantsal = 0;

                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                {
                                    cantsal++;

                                }

                            }
                        }

                        Horarios.Add("");

                        string sHora = string.Empty;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            sHora = miDetalle[i].Hora.ToString();

                            bool lEncontrado = false;

                            for (int j = 0; j <= Horarios.Count - 1; j++)
                            {
                                if (Horarios[j] == miDetalle[i].Hora.ToString())
                                {
                                    lEncontrado = true;

                                    break;
                                }
                            }

                            if (!lEncontrado)
                            {
                                Horarios.Add(miDetalle[i].Hora.ToString());
                            }
                        }

                        Horarios.RemoveAt(0);

                        Horarios.Sort();

                        string[] Celdas = { "", "", "B", "C", "D", "E", "F", "G", "H" };

                        if (Horarios.Count > 1)
                        {
                            oSheet.Row(41).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(36).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(31).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(26).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(21).InsertRowsBelow(Horarios.Count - 1);

                        }

                        int FilaBase = 21;

                        for (int j = 1; j <= 5; j++) //semanas
                        {

                            for (int k = 2; k <= 8; k++) //columnas
                            {
                                oSheet.Cell(FilaBase + Horarios.Count, k).FormulaA1 = string.Format("=COUNTA({0}{1}:{2}{3})", Celdas[k].ToString(), (FilaBase).ToString(), Celdas[k].ToString(), (FilaBase + Horarios.Count - 1).ToString());
                            }

                            FilaBase += 5 + Horarios.Count - 1;

                        }

                        int anio = Convert.ToInt32(miCabecera.AnoMes.ToString().Substring(0, 4));
                        int mes = Convert.ToInt32(miCabecera.AnoMes.ToString().Substring(4, 2));

                        int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                        for (int x = 0; x <= 4; x++)
                        {
                            FilaBase = 20 + (5 * x);

                            FilaBase += ((Horarios.Count - 1) * x) + 1;

                            for (int i = 0; i <= Horarios.Count - 1; i++)
                            {
                                oSheet.Cell(FilaBase + i, 1).Value = Horarios[i].ToString();

                            }

                            for (int i = 2; i <= 8; i++)
                            {
                                string DiaAnio = oSheet.Cell(FilaBase - 1, i).Value.ToString();

                                for (int j = 0; j <= miDetalle.Count - 1; j++)
                                {
                                    if (miDetalle[j].Fecha.ToString() == DiaAnio)
                                    {
                                        for (int k = 0; k <= Horarios.Count - 1; k++)
                                        {
                                            if (miDetalle[j].Hora.ToString() == oSheet.Cell(FilaBase + k, 1).Value.ToString())
                                            {
                                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[j].IdentifAviso));

                                                oSheet.Cell(FilaBase + k, i).Value = aviso.EtiquetaProd;

                                                break;

                                            }
                                        }
                                    }
                                }
                            }

                            decimal CosOp = 0;

                            for (int i = 0; i <= miDetalle.Count - 1; i++)
                            {
                                CosOp += miDetalle[i].CostoOp;
                            }

                            int CeldaCosOp = 48 + ((Horarios.Count - 1) * 5);

                            oSheet.Cells("H" + CeldaCosOp.ToString()).Value = CosOp;


                            ////firma//
                            oSheet.Cells("H" + (CeldaCosOp + 5).ToString()).Value = _Espacio.Responsable;

                            DAO.SetUpDAO STUP = new DAO.SetUpDAO();
                            SetUpDTO setup = STUP.Read(1);
                            oSheet.Cells("H" + (CeldaCosOp + 6).ToString()).Value = setup.Sector;

                            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                            oSheet.Cells("H" + (CeldaCosOp + 7).ToString()).Value = empresa.Name;

                            oSheet.Name = "O" + aniomes + miCabecera.IdentifEspacio;


                        }

                        break;
                    }
                #endregion
                #region Estimado
                case "ESTIMADO":
                    {

                        EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                        List<EstimadoDetDTO> miDetalle = (List<EstimadoDetDTO>)_Detalle;

                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };

                        AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[0].IdentifAviso));

                        var dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten);

                        AvisosIdAtenDTO ident =(AvisosIdAtenDTO)CRUDHelper.Read((string.Format("IdentifAviso = '{0}'", miDetalle[0].IdentifAviso)), dao);

                        dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion);

                        IdentAtencionDTO entity = (DTO.IdentAtencionDTO)CRUDHelper.Read(ident.RecId, dao);

                        if(entity != null)
                        {
                            oSheet.Cells("D16").Value = "Teléfonos " + entity.Telefono;
                        }

                        string aniomes = miCabecera.AnoMes.ToString();

                        List<string> IdentifAvisos = new List<string>();

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {

                            if (IdentifAvisos.Count == 0)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                            }

                            bool lencontrado = false;

                            string scadena = miDetalle[i].IdentifAviso;

                            for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                            {
                                if (IdentifAvisos[j] == scadena)
                                {
                                    lencontrado = true;

                                    break;
                                }
                            }

                            if (!lencontrado)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                                break;
                            }

                        }

                        List<string> Horarios = new List<string>();

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            int cantsal = 0;

                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                {
                                    cantsal++;

                                }

                            }
                        }

                        Horarios.Add("");

                        string sHora = string.Empty;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            sHora = miDetalle[i].Hora.ToString();

                            bool lEncontrado = false;

                            for (int j = 0; j <= Horarios.Count - 1; j++)
                            {
                                if (Horarios[j] == miDetalle[i].Hora.ToString())
                                {
                                    lEncontrado = true;

                                    break;
                                }
                            }

                            if (!lEncontrado)
                            {
                                Horarios.Add(miDetalle[i].Hora.ToString());
                            }
                        }

                        Horarios.RemoveAt(0);

                        Horarios.Sort();

                        string[] Celdas = { "", "", "B", "C", "D", "E", "F", "G", "H" };

                        if (Horarios.Count > 1)
                        {
                            oSheet.Row(41).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(36).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(31).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(26).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(21).InsertRowsBelow(Horarios.Count - 1);

                        }

                        int FilaBase = 21;

                        for (int j = 1; j <= 5; j++) //semanas
                        {

                            for (int k = 2; k <= 8; k++) //columnas
                            {
                                oSheet.Cell(FilaBase + Horarios.Count, k).FormulaA1 = string.Format("=COUNTA({0}{1}:{2}{3})", Celdas[k].ToString(), (FilaBase).ToString(), Celdas[k].ToString(), (FilaBase + Horarios.Count - 1).ToString());
                            }

                            FilaBase += 5 + Horarios.Count - 1;

                        }

                        int anio = Convert.ToInt32(miCabecera.AnoMes.ToString().Substring(0, 4));
                        int mes = Convert.ToInt32(miCabecera.AnoMes.ToString().Substring(4, 2));

                        int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                        for (int x = 0; x <= 4; x++)
                        {
                            FilaBase = 20 + (5 * x);

                            FilaBase += ((Horarios.Count - 1) * x) + 1;

                            for (int i = 0; i <= Horarios.Count - 1; i++)
                            {
                                oSheet.Cell(FilaBase + i, 1).Value = Horarios[i].ToString();

                            }

                            for (int i = 2; i <= 8; i++)
                            {
                                string DiaAnio = oSheet.Cell(FilaBase - 1, i).Value.ToString();

                                for (int j = 0; j <= miDetalle.Count - 1; j++)
                                {
                                    if (miDetalle[j].Fecha.ToString() == DiaAnio)
                                    {
                                        for (int k = 0; k <= Horarios.Count - 1; k++)
                                        {
                                            if (miDetalle[j].Hora.ToString() == oSheet.Cell(FilaBase + k, 1).Value.ToString())
                                            {
                                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[j].IdentifAviso));

                                                oSheet.Cell(FilaBase + k, i).Value = aviso.EtiquetaProd;

                                                break;

                                            }
                                        }
                                    }
                                }
                            }

                            decimal CosOp = 0;

                            for (int i = 0; i <= miDetalle.Count - 1; i++)
                            {
                                CosOp += miDetalle[i].CostoOp;
                            }

                            int CeldaCosOp = 48 + ((Horarios.Count - 1) * 5);

                            oSheet.Cells("H" + CeldaCosOp.ToString()).Value = CosOp;


                            ////firma//
                            oSheet.Cells("H" + (CeldaCosOp + 5).ToString()).Value = _Espacio.Responsable;

                            DAO.SetUpDAO STUP = new DAO.SetUpDAO();
                            SetUpDTO setup = STUP.Read(1);
                            oSheet.Cells("H" + (CeldaCosOp + 6).ToString()).Value = setup.Sector;

                            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                            oSheet.Cells("H" + (CeldaCosOp + 7).ToString()).Value = empresa.Name;

                            oSheet.Name = "O" + aniomes + miCabecera.IdentifEspacio;


                        }

                        break;
                    }
                #endregion
                #region Certificado
                case "CERTIFICADO":
                    {
                        CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                        List<CertificadoDetDTO> miDetalle = (List<CertificadoDetDTO>)_Detalle;

                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                        DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));

                        string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                        string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };

                        AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[0].IdentifAviso));

                        var dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten);

                        AvisosIdAtenDTO ident = (AvisosIdAtenDTO)CRUDHelper.Read((string.Format("IdentifAviso = '{0}'", miDetalle[0].IdentifAviso)), dao);

                        dao = BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.IdentAtencion);

                        IdentAtencionDTO entity = (DTO.IdentAtencionDTO)CRUDHelper.Read(ident.RecId, dao);

                        if (entity != null)
                        {
                            oSheet.Cells("D16").Value = "Teléfonos " + entity.Telefono;
                        }

                        string aniomes = miCabecera.AnoMes.ToString();

                        List<string> IdentifAvisos = new List<string>();

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {

                            if (IdentifAvisos.Count == 0)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                            }

                            bool lencontrado = false;

                            string scadena = miDetalle[i].IdentifAviso;

                            for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                            {
                                if (IdentifAvisos[j] == scadena)
                                {
                                    lencontrado = true;

                                    break;
                                }
                            }

                            if (!lencontrado)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                                break;
                            }

                        }

                        List<string> Horarios = new List<string>();

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            int cantsal = 0;

                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                {
                                    cantsal++;

                                }

                            }
                        }

                        Horarios.Add("");

                        string sHora = string.Empty;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            sHora = miDetalle[i].Hora.ToString();

                            bool lEncontrado = false;

                            for (int j = 0; j <= Horarios.Count - 1; j++)
                            {
                                if (Horarios[j] == miDetalle[i].Hora.ToString())
                                {
                                    lEncontrado = true;

                                    break;
                                }
                            }

                            if (!lEncontrado)
                            {
                                Horarios.Add(miDetalle[i].Hora.ToString());
                            }
                        }

                        Horarios.RemoveAt(0);

                        Horarios.Sort();

                        string[] Celdas = { "", "", "B", "C", "D", "E", "F", "G", "H" };

                        if (Horarios.Count > 1)
                        {
                            oSheet.Row(41).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(36).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(31).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(26).InsertRowsBelow(Horarios.Count - 1);
                            oSheet.Row(21).InsertRowsBelow(Horarios.Count - 1);

                        }

                        int FilaBase = 21;

                        for (int j = 1; j <= 5; j++) //semanas
                        {

                            for (int k = 2; k <= 8; k++) //columnas
                            {
                                oSheet.Cell(FilaBase + Horarios.Count, k).FormulaA1 = string.Format("=COUNTA({0}{1}:{2}{3})", Celdas[k].ToString(), (FilaBase).ToString(), Celdas[k].ToString(), (FilaBase + Horarios.Count - 1).ToString());
                            }

                            FilaBase += 5 + Horarios.Count - 1;

                        }

                        int anio = Convert.ToInt32(miCabecera.AnoMes.ToString().Substring(0, 4));
                        int mes = Convert.ToInt32(miCabecera.AnoMes.ToString().Substring(4, 2));

                        int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);

                        for (int x = 0; x <= 4; x++)
                        {
                            FilaBase = 20 + (5 * x);

                            FilaBase += ((Horarios.Count - 1) * x) + 1;

                            for (int i = 0; i <= Horarios.Count - 1; i++)
                            {
                                oSheet.Cell(FilaBase + i, 1).Value = Horarios[i].ToString();

                            }

                            for (int i = 2; i <= 8; i++)
                            {
                                string DiaAnio = oSheet.Cell(FilaBase - 1, i).Value.ToString();

                                for (int j = 0; j <= miDetalle.Count - 1; j++)
                                {
                                    if (miDetalle[j].Fecha.ToString() == DiaAnio)
                                    {
                                        for (int k = 0; k <= Horarios.Count - 1; k++)
                                        {
                                            if (miDetalle[j].Hora.ToString() == oSheet.Cell(FilaBase + k, 1).Value.ToString())
                                            {
                                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[j].IdentifAviso));

                                                oSheet.Cell(FilaBase + k, i).Value = aviso.EtiquetaProd;

                                                break;

                                            }
                                        }
                                    }
                                }
                            }

                            decimal CosOp = 0;

                            for (int i = 0; i <= miDetalle.Count - 1; i++)
                            {
                                CosOp += miDetalle[i].CostoOp;
                            }

                            int CeldaCosOp = 48 + ((Horarios.Count - 1) * 5);

                            oSheet.Cells("H" + CeldaCosOp.ToString()).Value = CosOp;


                            ////firma//
                            oSheet.Cells("H" + (CeldaCosOp + 5).ToString()).Value = _Espacio.Responsable;

                            DAO.SetUpDAO STUP = new DAO.SetUpDAO();
                            SetUpDTO setup = STUP.Read(1);
                            oSheet.Cells("H" + (CeldaCosOp + 6).ToString()).Value = setup.Sector;

                            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);
                            oSheet.Cells("H" + (CeldaCosOp + 7).ToString()).Value = empresa.Name;

                            oSheet.Name = "O" + aniomes + miCabecera.IdentifEspacio;


                        }

                        break;
                    }

                #endregion
            }

            oWB.SaveAs(nomArchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect();
        }

        public void Detalle_OP_Calendario_Numerico(string nomArchivo)
        {
            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomArchivo);
            var oSheet = oWB.Worksheet(1);

            int iDia = 0;
            int DiasEnMes = 0;
            string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
            string[] DiasSemana = { "", "L", "M", "M", "J", "V", "S", "D" };
            int[] DiasPautados = new int[10];
            int ocurrencias = 0;
            int MaxLines = 0;
            List<string> IdentifAvisos = new List<string>();

            switch (_Estado.ToUpper())
            {
                case "ORDENADO":
                    {
                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                        List<OrdenadoDetDTO> miDetalle = (List<OrdenadoDetDTO>)_Detalle;
                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();
                        int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;
                        string aniomes = miCabecera.AnoMes.ToString();
                        int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                        int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                        DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                        DiasPautados = new int[DiasEnMes + 1];
                        int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                        int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                        int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                        iDia = PrimerDiaSemana;

                        for (int x = 1; x <= DiasEnMes; x++)
                        {
                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].Fecha.Day == x)
                                {
                                    DiasPautados[x]++;
                                }
                            }
                        }

                        iDia = PrimerDiaSemana;

                        //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                        MaxLines = 0;

                        for (int k = 0; k <= DiasPautados.Length - 1; k++)
                        {
                            if (DiasPautados[k] > MaxLines)
                            {
                                MaxLines = DiasPautados[k];
                            }
                        }



                        IdentifAvisos.Add(miDetalle[0].IdentifAviso);

                        bool lencontrado = false;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            lencontrado = false;

                            for(int j = 0; j<=IdentifAvisos.Count-1;j++)
                            {
                                if (IdentifAvisos[j] == miDetalle[i].IdentifAviso)
                                {
                                    lencontrado = true;

                                    break;

                                }
                            }

                            if (!lencontrado)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                            }
                        }


                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {

                            int cantsal = 0;

                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                {

                                    int ThisDay = (int)miDetalle[j].Dia;
                                    ocurrencias = 0;

                                    for (int k = 0; k <= miDetalle.Count - 1; k++)
                                    {
                                        if ((int)miDetalle[k].Dia == ThisDay && miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                        {
                                            ocurrencias++;
                                        }
                                    }

                                    cantsal++;
                                }
                            }
                        }


                        //LIMPIEZA DE DIAS Y FECHAS
                        for (int i = 1; i <= 31; i++)
                        {
                            oSheet.Cell(18, 2 + i).Value = "";
                            oSheet.Cell(19, 2 + i).Value = "";
                        }

                        for (int i = 1; i <= (DiasEnMes); i++)
                        {
                            oSheet.Cell(18, 2 + i).Value = DiasSemana[iDia].ToUpper();
                            oSheet.Cell(19, 2 + i).Value = i;

                            iDia++;

                            if (iDia == 8)
                            {
                                iDia = 1;
                            }
                        }
                        /// FIN ARMADO CALENDARIO ///
                        /// 

                        /////////////////
                        //EMPIEZO A ARMAR FRAME PARA CADA UNO DE LOS PRODUCTOS

                        var xlSourceRange = oSheet.Range("B17:AK22");


                        for (int i = 1; i <= IdentifAvisos.Count - 1; i++)
                        {
                            var rng = oSheet.Range("B16:AK16");
                            rng.InsertRowsBelow(6);
                        }

                        int Fila = 17;

                        for (int i = 1; i <= IdentifAvisos.Count - 1; i++)
                        {
                            oSheet.Cell(Fila, 2).Value = xlSourceRange;

                            Fila = Fila + 6;
                        }


                        ////// CUENTO OCURRENCIAS X AVISO PARA INSERTAR LAS LINEAS RESEPECIIVAS /////

                        int[] FilasProducto = new int[IdentifAvisos.Count];

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            for (int diaMes = 1; diaMes <= DiasEnMes - 1; diaMes++)
                            {
                                ocurrencias = 0;

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if (miDetalle[k].IdentifAviso == IdentifAvisos[i] && miDetalle[k].Fecha.Day == diaMes && miDetalle[k].Fecha.Month == mes)
                                    {
                                        ocurrencias++;
                                    }

                                }

                                if (ocurrencias > 0)
                                {
                                    if (ocurrencias > FilasProducto[i])
                                    {
                                        FilasProducto[i] = ocurrencias;
                                    }
                                }


                            }
                        }

                        //inserta filas en cada uno de los productos de acuerdo a lo que hay en filasproducto
                        Fila = 20;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            var rng = oSheet.Range("B" + Fila.ToString() + ":AK" + Fila.ToString());

                            rng.InsertRowsBelow(FilasProducto[i]-1);

                            Fila += 6 + Convert.ToInt32(FilasProducto[i]-1);
                        }

                        ////// comienzo a rellenar con valores //

                        Fila = 17;

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[j].ToString()));

                            oSheet.Cell(Fila, 3).Value = aviso.EtiquetaProd;

                            Fila += 6 + FilasProducto[j]-1;
                        }


                        Fila = 20;
                        int FilaPosicional = 20;

                        string[] Celdas = { "", "", "", "C", "D", "E", "F", "G", "H", "I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG" };

                        miDetalle = miDetalle.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            for (int k = 1; k <= DiasEnMes; k++)
                            {

                                oSheet.Column(k + 2).Width = 3;

                                for (int l = 0; l <= miDetalle.Count - 1; l++)
                                {
                                    
                                    if (miDetalle[l].Fecha.Day == k &&
                                        miDetalle[l].Fecha.Month == mes &&
                                        miDetalle[l].IdentifAviso == IdentifAvisos[j])
                                    {
                                        oSheet.Cell(Fila, k + 2).Value = 1;

                                        int sumDesde = FilaPosicional;
                                        int sumHasta = FilaPosicional + FilasProducto[j]-1;

                                        string Formula = string.Format("=+SUM({0}{1}:{2}{3})", Celdas[k + 2].ToString(), sumDesde.ToString(), Celdas[k + 2].ToString(), sumHasta.ToString());

                                        oSheet.Cell(FilaPosicional + FilasProducto[j], k + 2).FormulaA1 = Formula;

                                        Fila++;
                                    }

                                }

                                Fila = FilaPosicional;
                            }

                            FilaPosicional = Fila += 6 + FilasProducto[j]-1;

                        }


                        //// SUBTOTALES DE CADA COLUMNA /////

                        int[] STP = new int[DiasEnMes + 1];
                        int[] STT = new int[DiasEnMes + 1];

                        FilaPosicional = 15;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            FilaPosicional = FilaPosicional + 5 + FilasProducto[i];

                            for (int j = 1; j <= DiasEnMes; j++)
                            {
                                if (oSheet.Cell(FilaPosicional, j + 2).Value.ToString() == "")
                                {
                                    STP[j] = Convert.ToInt32(oSheet.Cell(FilaPosicional, j + 2).Value + "0");
                                }
                                else
                                {
                                    STP[j] = Convert.ToInt32(oSheet.Cell(FilaPosicional, j + 2).Value);
                                }

                                STT[j] += STP[j];
                            }
                        }

                        FilaPosicional += 2;

                        for (int j = 1; j <= DiasEnMes; j++)
                        {
                            oSheet.Cell(FilaPosicional, j + 2).Value = STT[j];
                        }

                        ////// LLENO MARGEN DERECHO CON TOTALES ////////////

                        FilaPosicional = 15;

                        decimal CostoTotal = 0;
                        decimal GrandTotal = 0;
                        int UnidadesTotales = 0;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            FilaPosicional = FilaPosicional + 5;

                            for(int j=0;j<=FilasProducto[i]-1;j++)
                            {
                                oSheet.Cell("AH" + FilaPosicional.ToString()).FormulaA1 = string.Format("=SUM({0}{1}:{2}{3})", "C",FilaPosicional ,"AG", FilaPosicional);

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if (miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                    {
                                        oSheet.Cell("AI" + FilaPosicional.ToString()).Value = miDetalle[k].CostoOp;
                                        oSheet.Cell("AJ" + FilaPosicional.ToString()).Value = miDetalle[k].Duracion;
                                        oSheet.Cell("AK" + FilaPosicional.ToString()).Value = Convert.ToDecimal(oSheet.Cell("AH" + FilaPosicional.ToString()).Value) * Convert.ToDecimal(oSheet.Cell("AI" + FilaPosicional.ToString()).Value);  

                                        CostoTotal += Convert.ToDecimal(oSheet.Cell("AK" + FilaPosicional.ToString()).Value);

                                        break;
                                    }
                                }

                                oSheet.Cell("AK" + FilaPosicional.ToString()).Value = Convert.ToDecimal(oSheet.Cell("AH" + FilaPosicional.ToString()).Value) * Convert.ToDecimal(oSheet.Cell("AI" + FilaPosicional.ToString()).Value);  

                                FilaPosicional++;

                            }

                            oSheet.Cell("AH" + FilaPosicional.ToString()).FormulaA1 = string.Format("=SUM({0}{1}:{2}{3})", "C", FilaPosicional, "AG", FilaPosicional);
                            oSheet.Cell("AI" + FilaPosicional.ToString()).Value = "";
                            oSheet.Cell("AJ" + FilaPosicional.ToString()).Value = "";
                            oSheet.Cell("AK" + FilaPosicional.ToString()).Value = CostoTotal;

                            GrandTotal += CostoTotal;

                            UnidadesTotales += Convert.ToInt32(oSheet.Cell("AH" + FilaPosicional.ToString()).Value);

                            CostoTotal = 0;

                        }

                        FilaPosicional += 2;

                        oSheet.Cell("AH" + FilaPosicional.ToString()).Value = UnidadesTotales;
                        oSheet.Cell("AK" + FilaPosicional.ToString()).Value = GrandTotal;

                        FilaPosicional += 3;

                        oSheet.Cell("AI" + FilaPosicional.ToString()).Value = _Espacio.Responsable;


                        DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                        SetUpDTO setup = STUP.Read(1);

                        oSheet.Cell("AI" + (FilaPosicional + 1).ToString()).Value = setup.Sector;

                        EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

                        oSheet.Cell("AI" + (FilaPosicional + 2).ToString()).Value = empresa.Name;

                        ////////////////////////////////////////////////////
                        
                        break;
                    }

                case "ESTIMADO":
                    {
                        EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                        List<EstimadoDetDTO> miDetalle = (List<EstimadoDetDTO>)_Detalle;
                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();
                        int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;
                        string aniomes = miCabecera.AnoMes.ToString();
                        int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                        int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                        DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                        DiasPautados = new int[DiasEnMes + 1];
                        int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                        int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                        int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                        iDia = PrimerDiaSemana;

                        for (int x = 1; x <= DiasEnMes; x++)
                        {
                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].Fecha.Day == x)
                                {
                                    DiasPautados[x]++;
                                }
                            }
                        }

                        iDia = PrimerDiaSemana;

                        //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                        MaxLines = 0;

                        for (int k = 0; k <= DiasPautados.Length - 1; k++)
                        {
                            if (DiasPautados[k] > MaxLines)
                            {
                                MaxLines = DiasPautados[k];
                            }
                        }



                        IdentifAvisos.Add(miDetalle[0].IdentifAviso);

                        bool lencontrado = false;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            lencontrado = false;

                            for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                            {
                                if (IdentifAvisos[j] == miDetalle[i].IdentifAviso)
                                {
                                    lencontrado = true;

                                    break;

                                }
                            }

                            if (!lencontrado)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                            }
                        }


                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {

                            int cantsal = 0;

                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                {

                                    int ThisDay = (int)miDetalle[j].Dia;
                                    ocurrencias = 0;

                                    for (int k = 0; k <= miDetalle.Count - 1; k++)
                                    {
                                        if ((int)miDetalle[k].Dia == ThisDay && miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                        {
                                            ocurrencias++;
                                        }
                                    }

                                    cantsal++;
                                }
                            }
                        }


                        //LIMPIEZA DE DIAS Y FECHAS
                        for (int i = 1; i <= 31; i++)
                        {
                            oSheet.Cell(18, 2 + i).Value = "";
                            oSheet.Cell(19, 2 + i).Value = "";
                        }

                        for (int i = 1; i <= (DiasEnMes); i++)
                        {
                            oSheet.Cell(18, 2 + i).Value = DiasSemana[iDia].ToUpper();
                            oSheet.Cell(19, 2 + i).Value = i;

                            iDia++;

                            if (iDia == 8)
                            {
                                iDia = 1;
                            }
                        }
                        /// FIN ARMADO CALENDARIO ///
                        /// 

                        /////////////////
                        //EMPIEZO A ARMAR FRAME PARA CADA UNO DE LOS PRODUCTOS

                        var xlSourceRange = oSheet.Range("B17:AK22");


                        for (int i = 1; i <= IdentifAvisos.Count - 1; i++)
                        {
                            var rng = oSheet.Range("B16:AK16");
                            rng.InsertRowsBelow(6);
                        }

                        int Fila = 17;

                        for (int i = 1; i <= IdentifAvisos.Count - 1; i++)
                        {
                            oSheet.Cell(Fila, 2).Value = xlSourceRange;

                            Fila = Fila + 6;
                        }


                        ////// CUENTO OCURRENCIAS X AVISO PARA INSERTAR LAS LINEAS RESEPECIIVAS /////

                        int[] FilasProducto = new int[IdentifAvisos.Count];

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            for (int diaMes = 1; diaMes <= DiasEnMes - 1; diaMes++)
                            {
                                ocurrencias = 0;

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if (miDetalle[k].IdentifAviso == IdentifAvisos[i] && miDetalle[k].Fecha.Day == diaMes && miDetalle[k].Fecha.Month == mes)
                                    {
                                        ocurrencias++;
                                    }

                                }

                                if (ocurrencias > 0)
                                {
                                    if (ocurrencias > FilasProducto[i])
                                    {
                                        FilasProducto[i] = ocurrencias;
                                    }
                                }


                            }
                        }

                        //inserta filas en cada uno de los productos de acuerdo a lo que hay en filasproducto
                        Fila = 20;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            var rng = oSheet.Range("B" + Fila.ToString() + ":AK" + Fila.ToString());

                            rng.InsertRowsBelow(FilasProducto[i] - 1);

                            Fila += 6 + Convert.ToInt32(FilasProducto[i] - 1);
                        }

                        ////// comienzo a rellenar con valores //

                        Fila = 17;

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[j].ToString()));

                            oSheet.Cell(Fila, 3).Value = aviso.EtiquetaProd;

                            Fila += 6 + FilasProducto[j] - 1;
                        }


                        Fila = 20;
                        int FilaPosicional = 20;

                        string[] Celdas = { "", "", "", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG" };

                        miDetalle = miDetalle.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            for (int k = 1; k <= DiasEnMes; k++)
                            {

                                oSheet.Column(k + 2).Width = 3;

                                for (int l = 0; l <= miDetalle.Count - 1; l++)
                                {

                                    if (miDetalle[l].Fecha.Day == k &&
                                        miDetalle[l].Fecha.Month == mes &&
                                        miDetalle[l].IdentifAviso == IdentifAvisos[j])
                                    {
                                        oSheet.Cell(Fila, k + 2).Value = 1;

                                        int sumDesde = FilaPosicional;
                                        int sumHasta = FilaPosicional + FilasProducto[j] - 1;

                                        string Formula = string.Format("=+SUM({0}{1}:{2}{3})", Celdas[k + 2].ToString(), sumDesde.ToString(), Celdas[k + 2].ToString(), sumHasta.ToString());

                                        oSheet.Cell(FilaPosicional + FilasProducto[j], k + 2).FormulaA1 = Formula;

                                        Fila++;
                                    }

                                }

                                Fila = FilaPosicional;
                            }

                            FilaPosicional = Fila += 6 + FilasProducto[j] - 1;

                        }


                        //// SUBTOTALES DE CADA COLUMNA /////

                        int[] STP = new int[DiasEnMes + 1];
                        int[] STT = new int[DiasEnMes + 1];

                        FilaPosicional = 15;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            FilaPosicional = FilaPosicional + 5 + FilasProducto[i];

                            for (int j = 1; j <= DiasEnMes; j++)
                            {
                                if (oSheet.Cell(FilaPosicional, j + 2).Value.ToString() == "")
                                {
                                    STP[j] = Convert.ToInt32(oSheet.Cell(FilaPosicional, j + 2).Value + "0");
                                }
                                else
                                {
                                    STP[j] = Convert.ToInt32(oSheet.Cell(FilaPosicional, j + 2).Value);
                                }

                                STT[j] += STP[j];
                            }
                        }

                        FilaPosicional += 2;

                        for (int j = 1; j <= DiasEnMes; j++)
                        {
                            oSheet.Cell(FilaPosicional, j + 2).Value = STT[j];
                        }

                        ////// LLENO MARGEN DERECHO CON TOTALES ////////////

                        FilaPosicional = 15;

                        decimal CostoTotal = 0;
                        decimal GrandTotal = 0;
                        int UnidadesTotales = 0;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            FilaPosicional = FilaPosicional + 5;

                            for (int j = 0; j <= FilasProducto[i] - 1; j++)
                            {
                                oSheet.Cell("AH" + FilaPosicional.ToString()).FormulaA1 = string.Format("=SUM({0}{1}:{2}{3})", "C", FilaPosicional, "AG", FilaPosicional);

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if (miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                    {
                                        oSheet.Cell("AI" + FilaPosicional.ToString()).Value = miDetalle[k].CostoOp;
                                        oSheet.Cell("AJ" + FilaPosicional.ToString()).Value = miDetalle[k].Duracion;
                                        oSheet.Cell("AK" + FilaPosicional.ToString()).Value = Convert.ToDecimal(oSheet.Cell("AH" + FilaPosicional.ToString()).Value) * Convert.ToDecimal(oSheet.Cell("AI" + FilaPosicional.ToString()).Value);

                                        CostoTotal += Convert.ToDecimal(oSheet.Cell("AK" + FilaPosicional.ToString()).Value);

                                        break;
                                    }
                                }

                                oSheet.Cell("AK" + FilaPosicional.ToString()).Value = Convert.ToDecimal(oSheet.Cell("AH" + FilaPosicional.ToString()).Value) * Convert.ToDecimal(oSheet.Cell("AI" + FilaPosicional.ToString()).Value);

                                FilaPosicional++;

                            }

                            oSheet.Cell("AH" + FilaPosicional.ToString()).FormulaA1 = string.Format("=SUM({0}{1}:{2}{3})", "C", FilaPosicional, "AG", FilaPosicional);
                            oSheet.Cell("AI" + FilaPosicional.ToString()).Value = "";
                            oSheet.Cell("AJ" + FilaPosicional.ToString()).Value = "";
                            oSheet.Cell("AK" + FilaPosicional.ToString()).Value = CostoTotal;

                            GrandTotal += CostoTotal;

                            UnidadesTotales += Convert.ToInt32(oSheet.Cell("AH" + FilaPosicional.ToString()).Value);

                            CostoTotal = 0;

                        }

                        FilaPosicional += 2;

                        oSheet.Cell("AH" + FilaPosicional.ToString()).Value = UnidadesTotales;
                        oSheet.Cell("AK" + FilaPosicional.ToString()).Value = GrandTotal;

                        FilaPosicional += 3;

                        oSheet.Cell("AI" + FilaPosicional.ToString()).Value = _Espacio.Responsable;


                        DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                        SetUpDTO setup = STUP.Read(1);

                        oSheet.Cell("AI" + (FilaPosicional + 1).ToString()).Value = setup.Sector;

                        EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

                        oSheet.Cell("AI" + (FilaPosicional + 2).ToString()).Value = empresa.Name;

                        ////////////////////////////////////////////////////

                        break;

                    }
                case "CERTIFICADO":
                    {
                        CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                        List<CertificadoDetDTO> miDetalle = (List<CertificadoDetDTO>)_Detalle;
                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();
                        int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;
                        string aniomes = miCabecera.AnoMes.ToString();
                        int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                        int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                        DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                        DiasPautados = new int[DiasEnMes + 1];
                        int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                        int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                        int dia = PrimerDiaMes - PrimerDiaSemana + 2;

                        iDia = PrimerDiaSemana;

                        for (int x = 1; x <= DiasEnMes; x++)
                        {
                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].Fecha.Day == x)
                                {
                                    DiasPautados[x]++;
                                }
                            }
                        }

                        iDia = PrimerDiaSemana;

                        //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                        MaxLines = 0;

                        for (int k = 0; k <= DiasPautados.Length - 1; k++)
                        {
                            if (DiasPautados[k] > MaxLines)
                            {
                                MaxLines = DiasPautados[k];
                            }
                        }



                        IdentifAvisos.Add(miDetalle[0].IdentifAviso);

                        bool lencontrado = false;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            lencontrado = false;

                            for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                            {
                                if (IdentifAvisos[j] == miDetalle[i].IdentifAviso)
                                {
                                    lencontrado = true;

                                    break;

                                }
                            }

                            if (!lencontrado)
                            {
                                IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                            }
                        }


                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {

                            int cantsal = 0;

                            for (int j = 0; j <= miDetalle.Count - 1; j++)
                            {
                                if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                {

                                    int ThisDay = (int)miDetalle[j].Dia;
                                    ocurrencias = 0;

                                    for (int k = 0; k <= miDetalle.Count - 1; k++)
                                    {
                                        if ((int)miDetalle[k].Dia == ThisDay && miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                        {
                                            ocurrencias++;
                                        }
                                    }

                                    cantsal++;
                                }
                            }
                        }


                        //LIMPIEZA DE DIAS Y FECHAS
                        for (int i = 1; i <= 31; i++)
                        {
                            oSheet.Cell(18, 2 + i).Value = "";
                            oSheet.Cell(19, 2 + i).Value = "";
                        }

                        for (int i = 1; i <= (DiasEnMes); i++)
                        {
                            oSheet.Cell(18, 2 + i).Value = DiasSemana[iDia].ToUpper();
                            oSheet.Cell(19, 2 + i).Value = i;

                            iDia++;

                            if (iDia == 8)
                            {
                                iDia = 1;
                            }
                        }
                        /// FIN ARMADO CALENDARIO ///
                        /// 

                        /////////////////
                        //EMPIEZO A ARMAR FRAME PARA CADA UNO DE LOS PRODUCTOS

                        var xlSourceRange = oSheet.Range("B17:AK22");


                        for (int i = 1; i <= IdentifAvisos.Count - 1; i++)
                        {
                            var rng = oSheet.Range("B16:AK16");
                            rng.InsertRowsBelow(6);
                        }

                        int Fila = 17;

                        for (int i = 1; i <= IdentifAvisos.Count - 1; i++)
                        {
                            oSheet.Cell(Fila, 2).Value = xlSourceRange;

                            Fila = Fila + 6;
                        }


                        ////// CUENTO OCURRENCIAS X AVISO PARA INSERTAR LAS LINEAS RESEPECIIVAS /////

                        int[] FilasProducto = new int[IdentifAvisos.Count];

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            for (int diaMes = 1; diaMes <= DiasEnMes - 1; diaMes++)
                            {
                                ocurrencias = 0;

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if (miDetalle[k].IdentifAviso == IdentifAvisos[i] && miDetalle[k].Fecha.Day == diaMes && miDetalle[k].Fecha.Month == mes)
                                    {
                                        ocurrencias++;
                                    }

                                }

                                if (ocurrencias > 0)
                                {
                                    if (ocurrencias > FilasProducto[i])
                                    {
                                        FilasProducto[i] = ocurrencias;
                                    }
                                }


                            }
                        }

                        //inserta filas en cada uno de los productos de acuerdo a lo que hay en filasproducto
                        Fila = 20;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            var rng = oSheet.Range("B" + Fila.ToString() + ":AK" + Fila.ToString());

                            rng.InsertRowsBelow(FilasProducto[i] - 1);

                            Fila += 6 + Convert.ToInt32(FilasProducto[i] - 1);
                        }

                        ////// comienzo a rellenar con valores //

                        Fila = 17;

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[j].ToString()));

                            oSheet.Cell(Fila, 3).Value = aviso.EtiquetaProd;

                            Fila += 6 + FilasProducto[j] - 1;
                        }


                        Fila = 20;
                        int FilaPosicional = 20;

                        string[] Celdas = { "", "", "", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG" };

                        miDetalle = miDetalle.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            for (int k = 1; k <= DiasEnMes; k++)
                            {

                                oSheet.Column(k + 2).Width = 3;

                                for (int l = 0; l <= miDetalle.Count - 1; l++)
                                {

                                    if (miDetalle[l].Fecha.Day == k &&
                                        miDetalle[l].Fecha.Month == mes &&
                                        miDetalle[l].IdentifAviso == IdentifAvisos[j])
                                    {
                                        oSheet.Cell(Fila, k + 2).Value = 1;

                                        int sumDesde = FilaPosicional;
                                        int sumHasta = FilaPosicional + FilasProducto[j] - 1;

                                        string Formula = string.Format("=+SUM({0}{1}:{2}{3})", Celdas[k + 2].ToString(), sumDesde.ToString(), Celdas[k + 2].ToString(), sumHasta.ToString());

                                        oSheet.Cell(FilaPosicional + FilasProducto[j], k + 2).FormulaA1 = Formula;

                                        Fila++;
                                    }

                                }

                                Fila = FilaPosicional;
                            }

                            FilaPosicional = Fila += 6 + FilasProducto[j] - 1;

                        }


                        //// SUBTOTALES DE CADA COLUMNA /////

                        int[] STP = new int[DiasEnMes + 1];
                        int[] STT = new int[DiasEnMes + 1];

                        FilaPosicional = 15;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            FilaPosicional = FilaPosicional + 5 + FilasProducto[i];

                            for (int j = 1; j <= DiasEnMes; j++)
                            {
                                if (oSheet.Cell(FilaPosicional, j + 2).Value.ToString() == "")
                                {
                                    STP[j] = Convert.ToInt32(oSheet.Cell(FilaPosicional, j + 2).Value + "0");
                                }
                                else
                                {
                                    STP[j] = Convert.ToInt32(oSheet.Cell(FilaPosicional, j + 2).Value);
                                }

                                STT[j] += STP[j];
                            }
                        }

                        FilaPosicional += 2;

                        for (int j = 1; j <= DiasEnMes; j++)
                        {
                            oSheet.Cell(FilaPosicional, j + 2).Value = STT[j];
                        }

                        ////// LLENO MARGEN DERECHO CON TOTALES ////////////

                        FilaPosicional = 15;

                        decimal CostoTotal = 0;
                        decimal GrandTotal = 0;
                        int UnidadesTotales = 0;

                        for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                        {
                            FilaPosicional = FilaPosicional + 5;

                            for (int j = 0; j <= FilasProducto[i] - 1; j++)
                            {
                                oSheet.Cell("AH" + FilaPosicional.ToString()).FormulaA1 = string.Format("=SUM({0}{1}:{2}{3})", "C", FilaPosicional, "AG", FilaPosicional);

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if (miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                    {
                                        oSheet.Cell("AI" + FilaPosicional.ToString()).Value = miDetalle[k].CostoOp;
                                        oSheet.Cell("AJ" + FilaPosicional.ToString()).Value = miDetalle[k].Duracion;
                                        oSheet.Cell("AK" + FilaPosicional.ToString()).Value = Convert.ToDecimal(oSheet.Cell("AH" + FilaPosicional.ToString()).Value) * Convert.ToDecimal(oSheet.Cell("AI" + FilaPosicional.ToString()).Value);

                                        CostoTotal += Convert.ToDecimal(oSheet.Cell("AK" + FilaPosicional.ToString()).Value);

                                        break;
                                    }
                                }

                                oSheet.Cell("AK" + FilaPosicional.ToString()).Value = Convert.ToDecimal(oSheet.Cell("AH" + FilaPosicional.ToString()).Value) * Convert.ToDecimal(oSheet.Cell("AI" + FilaPosicional.ToString()).Value);

                                FilaPosicional++;

                            }

                            oSheet.Cell("AH" + FilaPosicional.ToString()).FormulaA1 = string.Format("=SUM({0}{1}:{2}{3})", "C", FilaPosicional, "AG", FilaPosicional);
                            oSheet.Cell("AI" + FilaPosicional.ToString()).Value = "";
                            oSheet.Cell("AJ" + FilaPosicional.ToString()).Value = "";
                            oSheet.Cell("AK" + FilaPosicional.ToString()).Value = CostoTotal;

                            GrandTotal += CostoTotal;

                            UnidadesTotales += Convert.ToInt32(oSheet.Cell("AH" + FilaPosicional.ToString()).Value);

                            CostoTotal = 0;

                        }

                        FilaPosicional += 2;

                        oSheet.Cell("AH" + FilaPosicional.ToString()).Value = UnidadesTotales;
                        oSheet.Cell("AK" + FilaPosicional.ToString()).Value = GrandTotal;

                        FilaPosicional += 3;

                        oSheet.Cell("AI" + FilaPosicional.ToString()).Value = _Espacio.Responsable;


                        DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                        SetUpDTO setup = STUP.Read(1);

                        oSheet.Cell("AI" + (FilaPosicional + 1).ToString()).Value = setup.Sector;

                        EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

                        oSheet.Cell("AI" + (FilaPosicional + 2).ToString()).Value = empresa.Name;

                        ////////////////////////////////////////////////////

                        break;
                    }

            }

            /////////////////

            oWB.SaveAs(nomArchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect();
        }

        public void Detalle_OP_Grafica(string nomArchivo)
        {

            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomArchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO": 
                    {
                        OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                        List<OrdenadoDetDTO> miDetalle = (List<OrdenadoDetDTO>)_Detalle;
                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                        int Fila = 18;

                        if (miDetalle.Count > 1)
                        {
                            for (int i = 1; i <= miDetalle.Count - 1; i++)
                            {
                               var rng = oSheet.Range("B" + Fila.ToString() + ":G" + (Fila).ToString());
                               rng.InsertRowsBelow(3);

                            }
                        }

                        Fila = 19;

                        for (int i = 1; i <= miDetalle.Count - 1; i++)
                        {
                            var xlSourceRange = oSheet.Range("B16:G17");

                            oSheet.Cell(Fila,2).Value = xlSourceRange;

                            Fila = Fila + 3;
                        }

                        Fila = 14;

                        decimal SubTotal = 0;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            Fila += 3;

                            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[i].IdentifAviso));
                            DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));
                            oSheet.Cell(Fila, 2).Value = medio.Name;
                            oSheet.Cell(Fila, 3).Value = _Espacio.Name;
                            oSheet.Cell(Fila, 4).Value = aviso.EtiquetaProd;
                            FormAvisoDTO formaviso = Business.CRUDHelper.Read(string.Format("IdentifFormAviso = '{0}'", aviso.IdentifFormAviso), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FormAviso));
                            oSheet.Cell(Fila, 5).Value = formaviso.Name;
                            oSheet.Cell(Fila, 6).Value = miDetalle[i].Dia.ToString("00") +"/" + miCabecera.AnoMes.ToString("00").Substring(4, 2) + "/" + miCabecera.AnoMes.ToString().Substring(0, 4);
                            oSheet.Cell(Fila, 7).Value = miDetalle[i].CostoOp;

                            SubTotal += miDetalle[i].CostoOp;

                        }

                        Fila += 2;
                        oSheet.Cell(Fila, 7).Value = SubTotal;

                        Fila += 6;

                        oSheet.Cell(Fila, 6).Value = _Espacio.Responsable;

                        DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                        SetUpDTO setup = STUP.Read(1);

                        string sector = setup.Sector;

                        oSheet.Cell(Fila + 1, 6).Value = sector;

                        oSheet.Style.Alignment.SetShrinkToFit();

                        var Cols = oSheet.Columns();
                        var Rows = oSheet.Rows();

                        foreach (IXLRow fila in Rows)
                        {
                            fila.AdjustToContents();
                        }

                        foreach( IXLColumn columna in Cols)
                        {
                            columna.AdjustToContents();
                            
                        }

                    break; 

                    }
                case "ESTIMADO": 
                    {

                        EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                        List<EstimadoDetDTO> miDetalle = (List<EstimadoDetDTO>)_Detalle;
                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                        int Fila = 18;

                        if (miDetalle.Count > 1)
                        {
                            for (int i = 1; i <= miDetalle.Count - 1; i++)
                            {
                                var rng = oSheet.Range("B" + Fila.ToString() + ":G" + (Fila).ToString());
                                rng.InsertRowsBelow(3);

                            }
                        }

                        Fila = 19;

                        for (int i = 1; i <= miDetalle.Count - 1; i++)
                        {
                            var xlSourceRange = oSheet.Range("B16:G17");

                            oSheet.Cell(Fila, 2).Value = xlSourceRange;

                            Fila = Fila + 3;
                        }

                        Fila = 14;

                        decimal SubTotal = 0;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            Fila += 3;

                            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[i].IdentifAviso));
                            DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));
                            oSheet.Cell(Fila, 2).Value = medio.Name;
                            oSheet.Cell(Fila, 3).Value = _Espacio.Name;
                            oSheet.Cell(Fila, 4).Value = aviso.EtiquetaProd;
                            FormAvisoDTO formaviso = Business.CRUDHelper.Read(string.Format("IdentifFormAviso = '{0}'", aviso.IdentifFormAviso), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FormAviso));
                            oSheet.Cell(Fila, 5).Value = formaviso.Name;
                            oSheet.Cell(Fila, 6).Value = miDetalle[i].Dia.ToString("00") + "/" + miCabecera.AnoMes.ToString("00").Substring(4, 2) + "/" + miCabecera.AnoMes.ToString().Substring(0, 4);
                            oSheet.Cell(Fila, 7).Value = miDetalle[i].CostoOp;

                            SubTotal += miDetalle[i].CostoOp;

                        }

                        Fila += 2;
                        oSheet.Cell(Fila, 7).Value = SubTotal;

                        Fila += 6;

                        oSheet.Cell(Fila, 6).Value = _Espacio.Responsable;

                        DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                        SetUpDTO setup = STUP.Read(1);

                        string sector = setup.Sector;

                        oSheet.Cell(Fila + 1, 6).Value = sector;

                        oSheet.Style.Alignment.SetShrinkToFit();

                        var Cols = oSheet.Columns();
                        var Rows = oSheet.Rows();

                        foreach (IXLRow fila in Rows)
                        {
                            fila.AdjustToContents();
                        }

                        foreach (IXLColumn columna in Cols)
                        {
                            columna.AdjustToContents();

                        }

                        break;
                    }
                case "CERTIFICADO": 
                    {

                        CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                        List<CertificadoDetDTO> miDetalle = (List<CertificadoDetDTO>)_Detalle;
                        miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                        int Fila = 18;

                        if (miDetalle.Count > 1)
                        {
                            for (int i = 1; i <= miDetalle.Count - 1; i++)
                            {
                                var rng = oSheet.Range("B" + Fila.ToString() + ":G" + (Fila).ToString());
                                rng.InsertRowsBelow(3);

                            }
                        }

                        Fila = 19;

                        for (int i = 1; i <= miDetalle.Count - 1; i++)
                        {
                            var xlSourceRange = oSheet.Range("B16:G17");

                            oSheet.Cell(Fila, 2).Value = xlSourceRange;

                            Fila = Fila + 3;
                        }

                        Fila = 14;

                        decimal SubTotal = 0;

                        for (int i = 0; i <= miDetalle.Count - 1; i++)
                        {
                            Fila += 3;

                            AvisosDTO aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", miDetalle[i].IdentifAviso));
                            DTO.MediosPubDTO medio = CRUDHelper.Read(string.Format("IdentifMedio = '{0}'", _Espacio.IdentifMedio), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.MediosPub));
                            oSheet.Cell(Fila, 2).Value = medio.Name;
                            oSheet.Cell(Fila, 3).Value = _Espacio.Name;
                            oSheet.Cell(Fila, 4).Value = aviso.EtiquetaProd;
                            FormAvisoDTO formaviso = Business.CRUDHelper.Read(string.Format("IdentifFormAviso = '{0}'", aviso.IdentifFormAviso), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.FormAviso));
                            oSheet.Cell(Fila, 5).Value = formaviso.Name;
                            oSheet.Cell(Fila, 6).Value = miDetalle[i].Dia.ToString("00") + "/" + miCabecera.AnoMes.ToString("00").Substring(4, 2) + "/" + miCabecera.AnoMes.ToString().Substring(0, 4);
                            oSheet.Cell(Fila, 7).Value = miDetalle[i].CostoOp;

                            SubTotal += miDetalle[i].CostoOp;

                        }

                        Fila += 2;
                        oSheet.Cell(Fila, 7).Value = SubTotal;

                        Fila += 6;

                        oSheet.Cell(Fila, 6).Value = _Espacio.Responsable;

                        DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                        SetUpDTO setup = STUP.Read(1);

                        string sector = setup.Sector;

                        oSheet.Cell(Fila + 1, 6).Value = sector;

                        oSheet.Style.Alignment.SetShrinkToFit();

                        var Cols = oSheet.Columns();
                        var Rows = oSheet.Rows();

                        foreach (IXLRow fila in Rows)
                        {
                            fila.AdjustToContents();
                        }

                        foreach (IXLColumn columna in Cols)
                        {
                            columna.AdjustToContents();

                        }

                        break;

                    }
            }


            oWB.SaveAs(nomArchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 
        
        }

        public void Detalle_OP_PNT_Salida(string nomarchivo)
        {

            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);

            int Espaciado = 4;

            switch (_Estado.ToUpper())
            {
                #region Ordenado
                case "ORDENADO": {

                    AvisosDTO aviso = null;
                    OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                    List<OrdenadoDetDTO> miDetalle = (List<OrdenadoDetDTO>)_Detalle;
                    miDetalle = miDetalle.OrderBy(p =>p.Dia).ThenBy(q =>q.Hora).ThenBy(r =>r.Salida).ToList();

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

                    //ARMADO DE CALENDARIO - MEJORAR
                    string aniomes = miCabecera.AnoMes.ToString();
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;
                    int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;
                    int[] DiasPautados = new int[DiasEnMes + 1];

                    for (int x = 1; x <= DiasEnMes; x++)
                    {
                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].Fecha.Day == x)
                            {
                                DiasPautados[x]++;
                            }
                        }
                    }

                    //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                    int MaxLines = 0;

                    for (int k = 0; k <= DiasPautados.Length - 1; k++)
                    {
                        if (DiasPautados[k] > MaxLines)
                        {
                            MaxLines = DiasPautados[k];
                        }
                    }

                    string[] Celdas = { "", "", "", "C", "D", "E", "F", "G", "H", "I" };

                    ////// inicio de semanas ///////////

                    int Salida = 0;
                    int _Fila = 21;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila).ToString() + ":N" + (_Fila).ToString());

                        rng.InsertRowsBelow(MaxLines-1);

                        for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                        {
                            Salida++;

                            oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                        }
                    }
                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20, Columnas).Value.ToString();
                        string bb = Convert.ToDateTime(aa).Day.ToString();


                        List<OrdenadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == aa.Substring(3,2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 0; i <= SubLista.Count - 1; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i].IdentifAviso));

                                if (oSheet.Cell(20 + i, 2).Value == null || oSheet.Cell(25 + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(21 + i, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(21 + i, Columnas).Value = aviso.EtiquetaProd;
                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(20 + MaxLines + 1, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (21).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), 20 + SubLista.Count);
                            }
                            else
                            {
                                oSheet.Cell(20 + MaxLines + 1, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (21).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }

                    ////// SEMANA 2 //////

                    Salida = 0;
                    _Fila = 21 + MaxLines -1 + Espaciado;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());

                        rng.InsertRowsBelow(MaxLines - 1);

                        for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                        {
                            Salida++;

                            oSheet.Cell("A" + (i +1)).Value = "SALIDA " + Salida.ToString();
                        }
                    }
                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();

                        string bb = Convert.ToDateTime(aa).Day.ToString();

                        List<OrdenadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == aa.Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i-1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines + 1, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila + 1).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + SubLista.Count + 1, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila + 1).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 3//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();

                        string bb = Convert.ToDateTime(aa).Day.ToString();

                        List<OrdenadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == aa.Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i -1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i -1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i -1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines -1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 4//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado -1;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(_Fila -1, Columnas).Value.ToString();

                        string bb = Convert.ToDateTime(aa).Day.ToString();

                        List<OrdenadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == aa.Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 5//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado -1;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(_Fila -1, Columnas).Value.ToString();

                        string bb = Convert.ToDateTime(aa).Day.ToString();

                        List<OrdenadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == aa.Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }

                    ///////////// fin de semanas //////////////

                    int Columna = 6; //COD. INGESTA
                    int Fila = 5;

                    while (1 == 1)
                    {
                        if (oSheet.Cell(Fila, Columna).Value != null && oSheet.Cell(Fila, Columna).Value.ToString() == "COD. INGESTA")
                        {

                            List<string> IdentifAvisos = new List<string>();

                            for (int i = 0; i <= miDetalle.Count - 1; i++)
                            {

                                if (IdentifAvisos.Count == 0)
                                {
                                    IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                                }

                                bool lencontrado = false;

                                string scadena = miDetalle[i].IdentifAviso;

                                for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                                {
                                    if (IdentifAvisos[j] == scadena)
                                    {
                                        lencontrado = true;

                                        break;
                                    }
                                }

                                if (!lencontrado)
                                {
                                    IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                                    break;
                                }

                            }

                            for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                            {

                                Fila++;

                                Columna = 3;

                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]));
                                oSheet.Cell(Fila, Columna).Value = aviso.EtiquetaProd;

                                DTO.PiezasArteDTO pieza = CRUDHelper.Read(string.Format("IdentifPieza = '{0}'", aviso.IdentifPieza),
                                                          BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

                                DTO.AnunInternosDTO ai = CRUDHelper.Read(string.Format("IdentifAnun = '{0}'", pieza.IdentifAnun),
                                                              BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AnunInternos));

                                oSheet.Cell(Fila, Columna + 1).Value = ai.Name;
                                oSheet.Cell(Fila, Columna + 2).Value = aviso.Zocalo;
                                oSheet.Cell(Fila, Columna + 3).Value = aviso.NroIngesta;

                                int cantsal = 0;

                                for (int j = 0; j <= miDetalle.Count - 1; j++)
                                {
                                    if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                    {
                                        cantsal++;
                                    }
                                }

                                oSheet.Cell(Fila, Columna + 4).Value = cantsal;

                                Columna = 9;

                                decimal CostoPorSalida = miDetalle[0].CostoOp / miDetalle.Count;
                                oSheet.Cell(Fila, Columna).Value = CostoPorSalida;


                            }

                            Fila += 7;

                            oSheet.Cell(Fila, Columna).Value = _Espacio.Responsable;

                            Fila++;

                            DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                            SetUpDTO setup = STUP.Read(1);

                            string sector = setup.Sector;

                            oSheet.Cell(Fila, Columna).Value = sector;

                            Fila++;

                            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

                            oSheet.Cell(Fila, Columna).Value = empresa.Name;

                            // telefonos //

                            for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                            {
                                AvisosIdAtenDTO aia = CRUDHelper.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten));
                                oSheet.Cell("E" + (Fila + i).ToString()).Value = "'" + aia.IdentifIdentAte;


                            }
                            //////////////


                            break;
                        }

                        Fila++;
                    }

                   
                    break;
                }
                #endregion

                #region Estimado
                case "ESTIMADO": {

                    AvisosDTO aviso = null;
                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                    List<EstimadoDetDTO> miDetalle = (List<EstimadoDetDTO>)_Detalle;
                    miDetalle = miDetalle.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

                    //ARMADO DE CALENDARIO - MEJORAR
                    string aniomes = miCabecera.AnoMes.ToString();
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;
                    int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;
                    int[] DiasPautados = new int[DiasEnMes + 1];

                    for (int x = 1; x <= DiasEnMes; x++)
                    {
                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].Fecha.Day == x)
                            {
                                DiasPautados[x]++;
                            }
                        }
                    }

                    //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                    int MaxLines = 0;

                    for (int k = 0; k <= DiasPautados.Length - 1; k++)
                    {
                        if (DiasPautados[k] > MaxLines)
                        {
                            MaxLines = DiasPautados[k];
                        }
                    }

                    string[] Celdas = { "", "", "", "C", "D", "E", "F", "G", "H", "I" };

                    ////// inicio de semanas ///////////

                    int Salida = 0;
                    int _Fila = 21;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila).ToString() + ":N" + (_Fila).ToString());

                        rng.InsertRowsBelow(MaxLines - 1);

                        for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                        {
                            Salida++;

                            oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                        }
                    }
                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20, Columnas).Value.ToString();

                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays( Convert.ToInt32(aa) - 2);

                        //string bb = Convert.ToDateTime(aa).Day.ToString();
                        string bb = theDate.Day.ToString();

                        List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 0; i <= SubLista.Count - 1; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i].IdentifAviso));

                                if (oSheet.Cell(20 + i, 2).Value == null || oSheet.Cell(25 + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(21 + i, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(21 + i, Columnas).Value = aviso.EtiquetaProd;
                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(20 + MaxLines + 1, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (21).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), 20 + SubLista.Count);
                            }
                            else
                            {
                                oSheet.Cell(20 + MaxLines + 1, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (21).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }

                    ////// SEMANA 2 //////

                    Salida = 0;

                    _Fila = 21 + MaxLines - 1 + Espaciado;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());

                        rng.InsertRowsBelow(MaxLines - 1);

                        for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                        {
                            Salida++;

                            oSheet.Cell("A" + (i + 1)).Value = "SALIDA " + Salida.ToString();
                        }
                    }
                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();

                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        //string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();

                        //DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 1);

                        ////string bb = Convert.ToDateTime(aa).Day.ToString();
                        //string bb = theDate.Day.ToString();

                        //List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString() == theDate.ToString().Substring(3, 2)).ToList();


                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines + 1, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila + 1).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + SubLista.Count + 1, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila + 1).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 3//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(_Fila - 1, Columnas).Value.ToString();
                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        //string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();
                        //DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 1);
                        //string bb = theDate.Day.ToString();
                        //List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString() == theDate.ToString().Substring(3, 2)).ToList();


                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 4//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado - 1;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {

                        string aa = oSheet.Cell(_Fila - 1, Columnas).Value.ToString();
                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 5//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado - 1;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {

                        string aa = oSheet.Cell(_Fila - 1, Columnas).Value.ToString();
                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }

                    ///////////// fin de semanas //////////////

                    int Columna = 6; //COD. INGESTA
                    int Fila = 5;

                    while (1 == 1)
                    {
                        if (oSheet.Cell(Fila, Columna).Value != null && oSheet.Cell(Fila, Columna).Value.ToString() == "COD. INGESTA")
                        {

                            List<string> IdentifAvisos = new List<string>();

                            for (int i = 0; i <= miDetalle.Count - 1; i++)
                            {

                                if (IdentifAvisos.Count == 0)
                                {
                                    IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                                }

                                bool lencontrado = false;

                                string scadena = miDetalle[i].IdentifAviso;

                                for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                                {
                                    if (IdentifAvisos[j] == scadena)
                                    {
                                        lencontrado = true;

                                        break;
                                    }
                                }

                                if (!lencontrado)
                                {
                                    IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                                    break;
                                }

                            }

                            for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                            {

                                Fila++;

                                Columna = 3;

                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]));
                                oSheet.Cell(Fila, Columna).Value = aviso.EtiquetaProd;

                                DTO.PiezasArteDTO pieza = CRUDHelper.Read(string.Format("IdentifPieza = '{0}'", aviso.IdentifPieza),
                                                          BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

                                DTO.AnunInternosDTO ai = CRUDHelper.Read(string.Format("IdentifAnun = '{0}'", pieza.IdentifAnun),
                                                              BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AnunInternos));

                                oSheet.Cell(Fila, Columna + 1).Value = ai.Name;
                                oSheet.Cell(Fila, Columna + 2).Value = aviso.Zocalo;
                                oSheet.Cell(Fila, Columna + 3).Value = aviso.NroIngesta;

                                int cantsal = 0;

                                for (int j = 0; j <= miDetalle.Count - 1; j++)
                                {
                                    if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                    {
                                        cantsal++;
                                    }
                                }

                                oSheet.Cell(Fila, Columna + 4).Value = cantsal;

                                Columna = 9;

                                decimal CostoPorSalida = miDetalle[0].CostoOp / miDetalle.Count;
                                oSheet.Cell(Fila, Columna).Value = CostoPorSalida;


                            }

                            Fila += 7;

                            oSheet.Cell(Fila, Columna).Value = _Espacio.Responsable;

                            Fila++;

                            DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                            SetUpDTO setup = STUP.Read(1);

                            string sector = setup.Sector;

                            oSheet.Cell(Fila, Columna).Value = sector;

                            Fila++;

                            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

                            oSheet.Cell(Fila, Columna).Value = empresa.Name;

                            // telefonos //

                            for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                            {
                                AvisosIdAtenDTO aia = CRUDHelper.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten));
                                oSheet.Cell("E" + (Fila + i).ToString()).Value = "'" + aia.IdentifIdentAte;


                            }
                            //////////////


                            break;
                        }

                        Fila++;
                    }


                    break;
                }
#endregion

                #region Certificado

                case "CERTIFICADO": {

                    AvisosDTO aviso = null;
                    CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                    List<CertificadoDetDTO> miDetalle = (List<CertificadoDetDTO>)_Detalle;
                    miDetalle = miDetalle.OrderBy(p => p.Dia).ThenBy(q => q.Hora).ThenBy(r => r.Salida).ToList();

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

                    //ARMADO DE CALENDARIO - MEJORAR
                    string aniomes = miCabecera.AnoMes.ToString();
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                    string[] DiasSemana = { "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;
                    int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;
                    int[] DiasPautados = new int[DiasEnMes + 1];

                    for (int x = 1; x <= DiasEnMes; x++)
                    {
                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].Fecha.Day == x)
                            {
                                DiasPautados[x]++;
                            }
                        }
                    }

                    //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                    int MaxLines = 0;

                    for (int k = 0; k <= DiasPautados.Length - 1; k++)
                    {
                        if (DiasPautados[k] > MaxLines)
                        {
                            MaxLines = DiasPautados[k];
                        }
                    }

                    string[] Celdas = { "", "", "", "C", "D", "E", "F", "G", "H", "I" };

                    ////// inicio de semanas ///////////

                    int Salida = 0;
                    int _Fila = 21;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila).ToString() + ":N" + (_Fila).ToString());

                        rng.InsertRowsBelow(MaxLines - 1);

                        for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                        {
                            Salida++;

                            oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                        }
                    }
                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20, Columnas).Value.ToString();

                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);

                        //string bb = Convert.ToDateTime(aa).Day.ToString();
                        string bb = theDate.Day.ToString();

                        List<CertificadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 0; i <= SubLista.Count - 1; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i].IdentifAviso));

                                if (oSheet.Cell(20 + i, 2).Value == null || oSheet.Cell(25 + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(21 + i, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(21 + i, Columnas).Value = aviso.EtiquetaProd;
                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(20 + MaxLines + 1, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (21).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), 20 + SubLista.Count);
                            }
                            else
                            {
                                oSheet.Cell(20 + MaxLines + 1, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (21).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }

                    ////// SEMANA 2 //////

                    Salida = 0;

                    _Fila = 21 + MaxLines - 1 + Espaciado;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());

                        rng.InsertRowsBelow(MaxLines - 1);

                        for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                        {
                            Salida++;

                            oSheet.Cell("A" + (i + 1)).Value = "SALIDA " + Salida.ToString();
                        }
                    }
                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();

                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<CertificadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        //string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();

                        //DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 1);

                        ////string bb = Convert.ToDateTime(aa).Day.ToString();
                        //string bb = theDate.Day.ToString();

                        //List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString() == theDate.ToString().Substring(3, 2)).ToList();


                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines + 1, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila + 1).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + SubLista.Count + 1, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila + 1).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 3//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {
                        string aa = oSheet.Cell(_Fila - 1, Columnas).Value.ToString();
                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<CertificadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        //string aa = oSheet.Cell(20 + MaxLines + 4, Columnas).Value.ToString();
                        //DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 1);
                        //string bb = theDate.Day.ToString();
                        //List<EstimadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString() == theDate.ToString().Substring(3, 2)).ToList();


                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 4//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado - 1;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {

                        string aa = oSheet.Cell(_Fila - 1, Columnas).Value.ToString();
                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<CertificadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }


                    //////////////////////

                    ////// SEMANA 5//////

                    Salida = 0;
                    _Fila += MaxLines + Espaciado - 1;

                    if (MaxLines > 1)
                    {
                        //INSERCION DE FILAS PARA SALIDAS ADICIONALES
                        var rng = oSheet.Range("A" + (_Fila + 1).ToString() + ":N" + (_Fila + 1).ToString());
                        rng.InsertRowsBelow(MaxLines - 1);
                    }

                    _Fila++;

                    for (int i = _Fila; i <= _Fila + (MaxLines - 1); i++)
                    {
                        Salida++;

                        oSheet.Cell("A" + (i)).Value = "SALIDA " + Salida.ToString();
                    }

                    aviso = null;

                    for (int Columnas = 3; Columnas <= 9; Columnas++)
                    {

                        string aa = oSheet.Cell(_Fila - 1, Columnas).Value.ToString();
                        DateTime theDate = new DateTime(miDetalle[0].Fecha.Year, 1, 1).AddDays(Convert.ToInt32(aa) - 2);
                        string bb = theDate.Day.ToString();
                        List<CertificadoDetDTO> SubLista = miDetalle.FindAll(x => x.Dia.ToString() == bb && x.Fecha.Month.ToString("00") == theDate.ToString().Substring(3, 2)).ToList();

                        if (SubLista.Count > 0)
                        {
                            for (int i = 1; i <= SubLista.Count; i++)
                            {
                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", SubLista[i - 1].IdentifAviso));

                                if (oSheet.Cell(_Fila + i - 1, 2).Value == null || oSheet.Cell(_Fila + i, 2).Value.ToString() == "")
                                {
                                    oSheet.Cell(_Fila + i - 1, 2).Value = aviso.Duracion.ToString() + " SEG.";
                                }

                                oSheet.Cell(_Fila + i - 1, Columnas).Value = aviso.EtiquetaProd;

                            }

                            if (SubLista.Count > 1)
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=COUNTA({0}" + (_Fila).ToString() + ":{1}{2})", Celdas[Columnas].ToString(), Celdas[Columnas].ToString(), _Fila + MaxLines - 1);
                            }
                            else
                            {
                                oSheet.Cell(_Fila + MaxLines, Columnas).FormulaA1 = string.Format("=+COUNTA({0}" + (_Fila).ToString() + ")", Celdas[Columnas].ToString());
                            }
                        }
                    }

                    ///////////// fin de semanas //////////////

                    int Columna = 6; //COD. INGESTA
                    int Fila = 5;

                    while (1 == 1)
                    {
                        if (oSheet.Cell(Fila, Columna).Value != null && oSheet.Cell(Fila, Columna).Value.ToString() == "COD. INGESTA")
                        {

                            List<string> IdentifAvisos = new List<string>();

                            for (int i = 0; i <= miDetalle.Count - 1; i++)
                            {

                                if (IdentifAvisos.Count == 0)
                                {
                                    IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                                }

                                bool lencontrado = false;

                                string scadena = miDetalle[i].IdentifAviso;

                                for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                                {
                                    if (IdentifAvisos[j] == scadena)
                                    {
                                        lencontrado = true;

                                        break;
                                    }
                                }

                                if (!lencontrado)
                                {
                                    IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                                    break;
                                }

                            }

                            for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                            {

                                Fila++;

                                Columna = 3;

                                aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]));
                                oSheet.Cell(Fila, Columna).Value = aviso.EtiquetaProd;

                                DTO.PiezasArteDTO pieza = CRUDHelper.Read(string.Format("IdentifPieza = '{0}'", aviso.IdentifPieza),
                                                          BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

                                DTO.AnunInternosDTO ai = CRUDHelper.Read(string.Format("IdentifAnun = '{0}'", pieza.IdentifAnun),
                                                              BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AnunInternos));

                                oSheet.Cell(Fila, Columna + 1).Value = ai.Name;
                                oSheet.Cell(Fila, Columna + 2).Value = aviso.Zocalo;
                                oSheet.Cell(Fila, Columna + 3).Value = aviso.NroIngesta;

                                int cantsal = 0;

                                for (int j = 0; j <= miDetalle.Count - 1; j++)
                                {
                                    if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                                    {
                                        cantsal++;
                                    }
                                }

                                oSheet.Cell(Fila, Columna + 4).Value = cantsal;

                                Columna = 9;

                                decimal CostoPorSalida = miDetalle[0].CostoOp / miDetalle.Count;
                                oSheet.Cell(Fila, Columna).Value = CostoPorSalida;


                            }

                            Fila += 7;

                            oSheet.Cell(Fila, Columna).Value = _Espacio.Responsable;

                            Fila++;

                            DAO.SetUpDAO STUP = new DAO.SetUpDAO();

                            SetUpDTO setup = STUP.Read(1);

                            string sector = setup.Sector;

                            oSheet.Cell(Fila, Columna).Value = sector;

                            Fila++;

                            EmpresaDTO empresa = new DAO.EmpresaDAO().Read(1);

                            oSheet.Cell(Fila, Columna).Value = empresa.Name;

                            // telefonos //

                            for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                            {
                                AvisosIdAtenDTO aia = CRUDHelper.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten));
                                oSheet.Cell("E" + (Fila + i).ToString()).Value = "'" + aia.IdentifIdentAte;


                            }
                            //////////////


                            break;
                        }

                        Fila++;
                    }


                    break;


                }
#endregion
            }

            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect(); 
        }

        public void Detalle_OP_PNT_Producto(string nomarchivo)
        {

            ////// ClosedXML ////////
            XLWorkbook oWB = new XLWorkbook(nomarchivo);
            var oSheet = oWB.Worksheet(1);

            switch (_Estado.ToUpper())
            {
                case "ORDENADO": { 

                    OrdenadoCabDTO miCabecera = (OrdenadoCabDTO)_Cabecera;
                    List<OrdenadoDetDTO> miDetalle = (List<OrdenadoDetDTO>)_Detalle;
                    miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string[] DiasSemana = { "", "L", "M", "M", "J", "V", "S", "D" };

                    AvisosDTO aviso = null;

                    string aniomes = miCabecera.AnoMes.ToString();
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                    int[] DiasPautados = new int[DiasEnMes + 1];
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;
                    int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;

                    //ARMADO DE CALENDARIO - MEJORAR

                    for (int x = 1; x <= DiasEnMes; x++)
                    {
                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].Fecha.Day == x)
                            {
                                DiasPautados[x]++;
                            }
                        }
                    }

                    int iDia = PrimerDiaSemana;
                    //LIMPIEZA DE DIAS Y FECHAS
                    for (int i = 1; i <= 31; i++)
                    {
                        oSheet.Cell(19, 6 + i).Value = "";
                        oSheet.Cell(20, 6 + i).Value = "";
                    }

                    iDia = PrimerDiaSemana;
            
                    for (int i = 1; i <= DiasEnMes; i++)
                    {

                        oSheet.Cell(19, 6 + i).Value = DiasSemana[iDia].ToUpper();
                        oSheet.Cell(20, 6 + i).Value = i;

                        if (oSheet.Cell(19, 6 + i).Value.ToString() == "S" || oSheet.Cell(19, 6 + i).Value.ToString() == "D")
                        {
                            for (int x = 21; x <= 25; x++)
                            {
                                Color color = Color.FromArgb(66, 66, 66);
                                //oSheet.Cell(x, 6 + i)].interior.color = color;
                            }
                        }
                        else
                        {
                            for (int x = 21; x <= 25; x++)
                            {
                                Color color = Color.FromArgb(128, 128, 128);
                                //oSheet.Cell(x, 6 + i).interior.color = color;
                            }
                        }

                        iDia++;

                        if (iDia == 8)
                        {
                            iDia = 1;
                        }
                    }

                    //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                    int MaxLines = 0;

                    for (int k = 0; k <= DiasPautados.Length - 1; k++)
                    {
                        if (DiasPautados[k] > MaxLines)
                        {
                            MaxLines = DiasPautados[k];
                        }
                    }

                    oSheet.Cell(22, 2).Value = _Espacio.Name;

                    List<string> IdentifAvisos = new List<string>();

                    for (int i = 0; i <= miDetalle.Count - 1; i++)
                    {

                        if (IdentifAvisos.Count == 0)
                        {
                            IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                        }

                        bool lencontrado = false;

                        string scadena = miDetalle[i].IdentifAviso;

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            if (IdentifAvisos[j] == scadena)
                            {
                                lencontrado = true;

                                break;
                            }
                        }

                        if (!lencontrado)
                        {
                            IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                            break;
                        }
                    }

                    int Fila = 21;

                    for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                    {
                        aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]));

                        DTO.PiezasArteDTO pieza = CRUDHelper.Read(string.Format("IdentifPieza = '{0}'", aviso.IdentifPieza),
                                                    BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

                        DTO.AnunInternosDTO ai = CRUDHelper.Read(string.Format("IdentifAnun = '{0}'", pieza.IdentifAnun),
                                                        BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AnunInternos));

                        oSheet.Cell(Fila + i, 3).Value = aviso.EtiquetaProd.ToUpper();
                        oSheet.Cell(Fila + i, 4).Value = ai.Name.ToUpper();
                        oSheet.Cell(Fila + i, 6).Value = aviso.Duracion.ToString() + " SEG.";

                        oSheet.Cell(35 + i, 3).Value = aviso.EtiquetaProd.ToUpper();
                        oSheet.Cell(35 + i, 4).Value = ai.Name;
                        oSheet.Cell(35 + i, 5).Value = aviso.Zocalo;
                        oSheet.Cell(35 + i, 6).Value = aviso.NroIngesta;

                        AvisosIdAtenDTO aia = CRUDHelper.Read(string.Format("IdentifAviso = '{0}'",IdentifAvisos[i]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten));

                        oSheet.Cell("D" + (30 + i).ToString()).Value = aia.IdentifIdentAte;
                        
                        int cantsal = 0;

                        decimal CosOp = 0;

                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                            {

                                int ThisDay = (int)miDetalle[j].Dia;
                                int ocurrencias = 0;

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if ((int)miDetalle[k].Dia == ThisDay && miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                    {
                                        ocurrencias++;
                                    }
                                }

                                CosOp = miDetalle[j].CostoOp;

                                oSheet.Cell(Fila + i, 6 + ThisDay).Value = ocurrencias;

                                cantsal++;

                            }

                        }

                        oSheet.Cell(Fila + i, 39).Value = CosOp;

                    }

                    oSheet.Name = _Estado.Substring(0, 1).ToUpper() + aniomes + miCabecera.IdentifEspacio;

                   break;
                };

                case "ESTIMADO":{

                    EstimadoCabDTO miCabecera = (EstimadoCabDTO)_Cabecera;
                    List<EstimadoDetDTO> miDetalle = (List<EstimadoDetDTO>)_Detalle;
                    miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string[] DiasSemana = { "", "L", "M", "M", "J", "V", "S", "D" };

                    AvisosDTO aviso = null;

                    string aniomes = miCabecera.AnoMes.ToString();
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                    int[] DiasPautados = new int[DiasEnMes + 1];
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;
                    int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;

                    //ARMADO DE CALENDARIO - MEJORAR

                    for (int x = 1; x <= DiasEnMes; x++)
                    {
                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].Fecha.Day == x)
                            {
                                DiasPautados[x]++;
                            }
                        }
                    }

                    int iDia = PrimerDiaSemana;
                    //LIMPIEZA DE DIAS Y FECHAS
                    for (int i = 1; i <= 31; i++)
                    {
                        oSheet.Cell(19, 6 + i).Value = "";
                        oSheet.Cell(20, 6 + i).Value = "";
                    }

                    iDia = PrimerDiaSemana;
            
                    for (int i = 1; i <= DiasEnMes; i++)
                    {

                        oSheet.Cell(19, 6 + i).Value = DiasSemana[iDia].ToUpper();
                        oSheet.Cell(20, 6 + i).Value = i;

                        if (oSheet.Cell(19, 6 + i).Value.ToString() == "S" || oSheet.Cell(19, 6 + i).Value.ToString() == "D")
                        {
                            for (int x = 21; x <= 25; x++)
                            {
                                Color color = Color.FromArgb(66, 66, 66);
                                //oSheet.Cell(x, 6 + i)].interior.color = color;
                            }
                        }
                        else
                        {
                            for (int x = 21; x <= 25; x++)
                            {
                                Color color = Color.FromArgb(128, 128, 128);
                                //oSheet.Cell(x, 6 + i).interior.color = color;
                            }
                        }

                        iDia++;

                        if (iDia == 8)
                        {
                            iDia = 1;
                        }
                    }

                    //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                    int MaxLines = 0;

                    for (int k = 0; k <= DiasPautados.Length - 1; k++)
                    {
                        if (DiasPautados[k] > MaxLines)
                        {
                            MaxLines = DiasPautados[k];
                        }
                    }

                    oSheet.Cell(22, 2).Value = _Espacio.Name;

                    List<string> IdentifAvisos = new List<string>();

                    for (int i = 0; i <= miDetalle.Count - 1; i++)
                    {

                        if (IdentifAvisos.Count == 0)
                        {
                            IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                        }

                        bool lencontrado = false;

                        string scadena = miDetalle[i].IdentifAviso;

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            if (IdentifAvisos[j] == scadena)
                            {
                                lencontrado = true;

                                break;
                            }
                        }

                        if (!lencontrado)
                        {
                            IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                            break;
                        }
                    }

                    int Fila = 21;

                    for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                    {
                        aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]));

                        DTO.PiezasArteDTO pieza = CRUDHelper.Read(string.Format("IdentifPieza = '{0}'", aviso.IdentifPieza),
                                                    BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

                        DTO.AnunInternosDTO ai = CRUDHelper.Read(string.Format("IdentifAnun = '{0}'", pieza.IdentifAnun),
                                                        BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AnunInternos));

                        oSheet.Cell(Fila + i, 3).Value = aviso.EtiquetaProd.ToUpper();
                        oSheet.Cell(Fila + i, 4).Value = ai.Name.ToUpper();
                        oSheet.Cell(Fila + i, 6).Value = aviso.Duracion.ToString() + " SEG.";

                        oSheet.Cell(35 + i, 3).Value = aviso.EtiquetaProd.ToUpper();
                        oSheet.Cell(35 + i, 4).Value = ai.Name;
                        oSheet.Cell(35 + i, 5).Value = aviso.Zocalo;
                        oSheet.Cell(35 + i, 6).Value = aviso.NroIngesta;

                        AvisosIdAtenDTO aia = CRUDHelper.Read(string.Format("IdentifAviso = '{0}'",IdentifAvisos[i]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten));

                        oSheet.Cell("D" + (30 + i).ToString()).Value = aia.IdentifIdentAte;
                        
                        int cantsal = 0;

                        decimal CosOp = 0;

                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                            {

                                int ThisDay = (int)miDetalle[j].Dia;
                                int ocurrencias = 0;

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if ((int)miDetalle[k].Dia == ThisDay && miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                    {
                                        ocurrencias++;
                                    }
                                }

                                CosOp = miDetalle[j].CostoOp;

                                oSheet.Cell(Fila + i, 6 + ThisDay).Value = ocurrencias;

                                cantsal++;

                            }

                        }

                        oSheet.Cell(Fila + i, 39).Value = CosOp;

                    }

                    oSheet.Name = _Estado.Substring(0, 1).ToUpper() + aniomes + miCabecera.IdentifEspacio;

                   break;
                };
                #endregion
                #region Certificado
                case "CERTIFICADO":{

                    CertificadoCabDTO miCabecera = (CertificadoCabDTO)_Cabecera;
                    List<CertificadoDetDTO> miDetalle = (List<CertificadoDetDTO>)_Detalle;
                    miDetalle = miDetalle.OrderBy(P => P.Dia).ThenBy(Q => Q.Hora).ThenBy(R => R.Salida).ToList();

                    string[] Meses = { "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                    string[] DiasSemana = { "", "L", "M", "M", "J", "V", "S", "D" };

                    AvisosDTO aviso = null;

                    string aniomes = miCabecera.AnoMes.ToString();
                    int mes = Convert.ToInt32(aniomes.Substring(4, 2));
                    int anio = Convert.ToInt32(aniomes.Substring(0, 4));
                    int DiasEnMes = System.DateTime.DaysInMonth(anio, mes);
                    int[] DiasPautados = new int[DiasEnMes + 1];
                    int PrimerDiaSemana = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfWeek);
                    int PrimerDiaMes = Convert.ToInt32(Convert.ToDateTime(aniomes.Substring(0, 4) + "-" + aniomes.Substring(4, 2) + "-" + "01").DayOfYear);
                    int dia = PrimerDiaMes - PrimerDiaSemana + 2;
                    int DiaDelAnio = miDetalle[0].Fecha.DayOfYear;

                    //ARMADO DE CALENDARIO - MEJORAR

                    for (int x = 1; x <= DiasEnMes; x++)
                    {
                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].Fecha.Day == x)
                            {
                                DiasPautados[x]++;
                            }
                        }
                    }

                    int iDia = PrimerDiaSemana;
                    //LIMPIEZA DE DIAS Y FECHAS
                    for (int i = 1; i <= 31; i++)
                    {
                        oSheet.Cell(19, 6 + i).Value = "";
                        oSheet.Cell(20, 6 + i).Value = "";
                    }

                    iDia = PrimerDiaSemana;

                    for (int i = 1; i <= DiasEnMes; i++)
                    {

                        oSheet.Cell(19, 6 + i).Value = DiasSemana[iDia].ToUpper();
                        oSheet.Cell(20, 6 + i).Value = i;

                        if (oSheet.Cell(19, 6 + i).Value.ToString() == "S" || oSheet.Cell(19, 6 + i).Value.ToString() == "D")
                        {
                            for (int x = 21; x <= 25; x++)
                            {
                                Color color = Color.FromArgb(66, 66, 66);
                                //oSheet.Cell(x, 6 + i)].interior.color = color;
                            }
                        }
                        else
                        {
                            for (int x = 21; x <= 25; x++)
                            {
                                Color color = Color.FromArgb(128, 128, 128);
                                //oSheet.Cell(x, 6 + i).interior.color = color;
                            }
                        }

                        iDia++;

                        if (iDia == 8)
                        {
                            iDia = 1;
                        }
                    }

                    //CALCULA LA MAXIMA CANTIDAD DE SALIDAS DE AVISOS
                    int MaxLines = 0;

                    for (int k = 0; k <= DiasPautados.Length - 1; k++)
                    {
                        if (DiasPautados[k] > MaxLines)
                        {
                            MaxLines = DiasPautados[k];
                        }
                    }

                    oSheet.Cell(22, 2).Value = _Espacio.Name;

                    List<string> IdentifAvisos = new List<string>();

                    for (int i = 0; i <= miDetalle.Count - 1; i++)
                    {

                        if (IdentifAvisos.Count == 0)
                        {
                            IdentifAvisos.Add(miDetalle[i].IdentifAviso);
                        }

                        bool lencontrado = false;

                        string scadena = miDetalle[i].IdentifAviso;

                        for (int j = 0; j <= IdentifAvisos.Count - 1; j++)
                        {
                            if (IdentifAvisos[j] == scadena)
                            {
                                lencontrado = true;

                                break;
                            }
                        }

                        if (!lencontrado)
                        {
                            IdentifAvisos.Add(miDetalle[i].IdentifAviso);

                            break;
                        }
                    }

                    int Fila = 21;

                    for (int i = 0; i <= IdentifAvisos.Count - 1; i++)
                    {
                        aviso = Business.Avisos.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]));

                        DTO.PiezasArteDTO pieza = CRUDHelper.Read(string.Format("IdentifPieza = '{0}'", aviso.IdentifPieza),
                                                    BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.PiezasArte));

                        DTO.AnunInternosDTO ai = CRUDHelper.Read(string.Format("IdentifAnun = '{0}'", pieza.IdentifAnun),
                                                        BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AnunInternos));

                        oSheet.Cell(Fila + i, 3).Value = aviso.EtiquetaProd.ToUpper();
                        oSheet.Cell(Fila + i, 4).Value = ai.Name.ToUpper();
                        oSheet.Cell(Fila + i, 6).Value = aviso.Duracion.ToString() + " SEG.";

                        oSheet.Cell(35 + i, 3).Value = aviso.EtiquetaProd.ToUpper();
                        oSheet.Cell(35 + i, 4).Value = ai.Name;
                        oSheet.Cell(35 + i, 5).Value = aviso.Zocalo;
                        oSheet.Cell(35 + i, 6).Value = aviso.NroIngesta;

                        AvisosIdAtenDTO aia = CRUDHelper.Read(string.Format("IdentifAviso = '{0}'", IdentifAvisos[i]), BusinessMapper.GetDaoByEntity(BusinessMapper.eEntities.AvisosIdAten));

                        oSheet.Cell("D" + (30 + i).ToString()).Value = aia.IdentifIdentAte;

                        int cantsal = 0;

                        decimal CosOp = 0;

                        for (int j = 0; j <= miDetalle.Count - 1; j++)
                        {
                            if (miDetalle[j].IdentifAviso == IdentifAvisos[i])
                            {

                                int ThisDay = (int)miDetalle[j].Dia;
                                int ocurrencias = 0;

                                for (int k = 0; k <= miDetalle.Count - 1; k++)
                                {
                                    if ((int)miDetalle[k].Dia == ThisDay && miDetalle[k].IdentifAviso == IdentifAvisos[i])
                                    {
                                        ocurrencias++;
                                    }
                                }

                                CosOp = miDetalle[j].CostoOp;

                                oSheet.Cell(Fila + i, 6 + ThisDay).Value = ocurrencias;

                                cantsal++;

                            }

                        }

                        oSheet.Cell(Fila + i, 39).Value = CosOp;

                    }

                    oSheet.Name = _Estado.Substring(0, 1).ToUpper() + aniomes + miCabecera.IdentifEspacio;

                    break;
                }
                #endregion
            }


            oWB.SaveAs(nomarchivo);
            oWB.Dispose();
            oSheet = null;
            oWB = null;
            GC.Collect();
        }

#region Pies

        public void Pie_OP_Calendario_Descriptivo(){ }

        public void Pie_OP_Calendario_Numerico(){ }

        public void Pie_OP_Grafica(){ }

        public void Pie_OP_PNT_Producto(){ }

        public void Pie_OP_PNT_SALIDA(){ }

#endregion

    }
}