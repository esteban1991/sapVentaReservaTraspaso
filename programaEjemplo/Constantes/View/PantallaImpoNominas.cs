using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ventaRT.Constantes.Views
{
    public static class PantallaImpoNominas
    {
        public const string DataTableLog = "dt_logNomi";
        public const string EditTxtFichero = "txt_Ruta";
        public const string GridLog = "gr_noDpto";
        public const string ButtonBuscarFichero = "btn_fit";
        public const string ButtonImportacion = "btn_imp";
        public const string TxtFechaContab = "txt_Fecha";
        public const string CmbSerie = "CMB_Serie";
    }

    public static class PantallaImpoPuente
    {
        public const string DataTableLog = "dt_logNomi";
        public const string EditTxtFichero = "txt_Ruta";
        public const string GridLog = "gr_Puente";
        public const string ButtonBuscarFichero = "btn_fit";
        public const string ButtonImportacion = "btn_imp2";

    }

    public static class oitm
    {
        public const string OITM = "OITM";
        public const string ItemCode = "\"ItemCode\"";
        public const string ItemName = "\"ItemName\"";
    }


    public static class ColGridLog
    {
        public const string PrCode = "\"PrcCode\"";
        public const string Nombre = "\"PrcName\"";
        public const string Dpto = "\"U_SSI_DPTOS\"";
        public const string ColMensaje = "Mensaje";
        public const string Tabla = "OPRC";
        public const string DimCode = "\"DimCode\"";
        public const string ColLinea = "Línea";
        public const string ColLineaSize = "15";
        public const string ColMensajeSize = "250";
        public const string ColObjectType = "ObjectType";
        public const string ColObjectTypeSize = "10";
        public const string ColTipoError = "Error";
        public const string ColTipoErrorSize = "100";
        public const string ColDatoCreado = "Dato Creado";
        public const string ColDatoCreadoSize = "50";
        public const string ColAsiento = "Asiento";
        public const string ColFecha = "Fecha";


    }

    public static class SerieConsulta {

        public const string SeriesName = "\"SeriesName\"";
        public const string Indicator = "\"Indicator\"";
        public const string GroupCode = "\"GroupCode\"";
        public const string F_RefDate = "\"F_RefDate\"";
        public const string NNM1 = "NNM1";
        public const string OFPR = "OFPR";
        public const string ObjectCode = "\"ObjectCode\"";
        public const string Series = "\"Series\"";
        

    }

    public static class puenteConsulta
    {

        public const string ctaPteQry = "\"U_SSI_CtaPu\"";
        public const string SSICtaDesc = "\"U_SSICtaDesc\"";
        public const string UOADM = "\"@OADM\"";



    }

    public static class CAB_REC_IMPT
    {
        public const string CAB_REC_IMP = "\"@CAB_REC_IMP\"";
        public const string Code = "\"Code\"";
        public const string U_Num_OC = "\"U_Num_OC\"";
        public const string U_Nom_Creador = "\"U_Nom_Creador\"";
        public const string U_Nom_Autorizador = "\"U_Nom_Autorizador\"";
        public const string U_Total_Carton = "\"U_Total_Carton\"";
        public const string U_Comentarios = "\"U_Comentarios\"";
        public const string U_Nom_Proveedor = "\"U_Nom_Proveedor\"";
        public const string U_Fecha = "\"U_Fecha\"";
        public const string U_Estado = "\"U_Estado\"";
    }

    public static class DET_REC_IMPT
    {
        public const string DET_REC_IMP = "\"@DET_REC_IMP\"";
        public const string U_Cod_Articulo = "\"U_Cod_Articulo\"";
        public const string U_Nom_Articulo = "\"U_Nom_Articulo\"";
        public const string U_Cant_Carton = "\"U_Cant_Carton\"";
        public const string U_Cant_OC = "\"U_Cant_OC\"";
        public const string U_Case_Pack = "\"U_Case_Pack\"";
        public const string U_Cant_Recibida = "\"U_Cant_Recibida\"";
        public const string U_EanBar = "\"U_EanBar\"";
        public const string U_Peso = "\"U_Peso\"";
        public const string U_Largo = "\"U_Largo\"";
        public const string U_Ancho = "\"U_Ancho\"";
        public const string U_Alto = "\"U_Alto\"";
        public const string U_Inner_Pack = "\"U_Inner_Pack\"";
        public const string U_DunBar = "\"U_DunBar\"";
        public const string U_C_Peso = "\"U_C_Peso\"";
        public const string U_C_Largo = "\"U_C_Largo\"";
        public const string U_C_Ancho = "\"U_C_Ancho\"";
        public const string U_C_Alto = "\"U_C_Alto\"";
        public const string U_Base = "\"U_Base\"";
        public const string U_Altura = "\"U_Altura\"";
        public const string U_Saldo = "\"U_Saldo\"";
        public const string U_Carton_Pallet = "\"U_Carton_Pallet\"";
        public const string U_Unidad_Pallet = "\"U_Unidad_Pallet\"";
        public const string U_Num_OC = "\"U_Num_OC\"";
        public const string Code = "\"Code\"";
    
       
    
    }

    public static class CAB_REC_IMPTTabla
    {
        public const string CAB_REC_IMP = "@CAB_REC_IMP";
        public const string Code = "Code";
        public const string U_Num_OC = "U_Num_OC";
        public const string U_Nom_Creador = "U_Nom_Creador";
        public const string U_Nom_Autorizador = "U_Nom_Autorizador";
        public const string U_Total_Carton = "U_Total_Carton";
        public const string U_Comentarios = "U_Comentarios";
        public const string U_Nom_Proveedor = "U_Nom_Proveedor";
        public const string U_Fecha = "U_Fecha";
        public const string U_Estado = "U_Estado";
    }

    public static class DET_REC_IMPTTabla
    {
        public const string DET_REC_IMP = "@DET_REC_IMP";
        public const string U_Cod_Articulo = "U_Cod_Articulo";
        public const string U_Nom_Articulo = "U_Nom_Articulo";
        public const string U_Cant_OC = "U_Cant_OC";
        public const string U_Cant_Carton = "U_Cant_Carton";
        public const string U_Case_Pack = "U_Case_Pack";
        public const string U_Cant_Recibida = "U_Cant_Recibida";
        public const string U_EanBar = "U_EanBar";
        public const string U_Peso = "U_Peso";
        public const string U_Largo = "U_Largo";
        public const string U_Ancho = "U_Ancho";
        public const string U_Alto = "U_Alto";
        public const string U_Inner_Pack = "U_Inner_Pack";
        public const string U_DunBar = "U_DunBar";
        public const string U_C_Peso = "U_C_Peso";
        public const string U_C_Largo = "U_C_Largo";
        public const string U_C_Ancho = "U_C_Ancho";
        public const string U_C_Alto = "U_C_Alto";
        public const string U_Base = "U_Base";
        public const string U_Altura = "U_Altura";
        public const string U_Saldo = "U_Saldo";
        public const string U_Carton_Pallet = "U_Carton_Pallet";
        public const string U_Unidad_Pallet = "U_Unidad_Pallet";
        public const string U_Num_OC = "U_Num_OC";
        public const string Code = "Code";

    }

    public static class Oc_cabecera
    {
        public const string OPOR = "OPOR";
        public const string DocNum = "\"DocNum\"";
        public const string DocEntry = "\"DocEntry\"";
        public const string CardName = "\"CardName\"";

    }

    public static class Oc_Detalle
    {
        public const string POR1 = "POR1";
        public const string LineNum = "\"LineNum\"";
        public const string ItemCode = "\"ItemCode\"";
        public const string Dscription = "\"Dscription\"";
        public const string Quantity = "\"Quantity\"";
    }

    public static class formatos{

        public const string formatoFechaQrie = "yyyy-MM-dd";

        public const string formatoFechaAsiento = "dd-MM-yyyy";


    }

    public static class Cuentas
    {
        // barcelona  y becarios
        public const String BarEmbSalarial = "465000";
        public const String BarDescPresta = "242500";
        public const String BarDescPagExtra = "465050";
        public const String BarCSaldo = "649050";
        public const String BarCotiComunes = "476000";
        public const String BarIrpf = "475100";
        public const String BarTotalLiquido = "465050";
        public const String BarIrpfDineraria = "640050";
        public const String BarCvalores = "649050";
        public const String CtaPuente = "465999";

        // Madrid
        public const String MadEmbSalarial = "465000";
        public const String MadDescPresta = "242500";
        public const String MadDescPagExtra = "465060";
        public const String MadCSaldo = "649050";
        public const String ImpCtaVlrEspecie = "649050";
        public const String MadCotiComunes = "476000";
        public const String MadrIrpf = "475100";
        public const String MadTotalLiquido = "465060";
        public const String MadIrpfDineraria = "640060";
        public const String MadCvalores = "649060";


        // Canarias
        public const String CanEmbSalarial = "465000";
        public const String CanDescPresta = "242500";
        public const String CanDescPagExtra = "465070";
        public const String CanCSaldo = "649070";
        public const String CanCotiComunes = "476000";
        public const String CanIrpf = "475100";
        public const String CanTotalLiquido = "465070";
        public const String CanIrpfDineraria = "640070";
        public const String CanCvalores = "649070";


        // Andalucia
        public const String AndEmbSalarial = "465000";
        public const String AndDescPresta = "242500"; 
        public const String AndDescPagExtra = "465080";
        public const String AndCSaldo = "649080";
        public const String AndCotiComunes = "476000";
        public const String AndIrpf = "475100";
        public const String AndTotalLiquido = "465080";
        public const String AndIrpfDineraria = "640080";
        public const String AndCvalores = "649080";
    }
        public static class ASIENTOS
        {
            public const String Barcelona = "00001";
            public const String Madrid = "00002";
            public const String Andalucia = "00003";
            public const String Canarias = "00004";
            public const String Becarios = "00005";
        }
}