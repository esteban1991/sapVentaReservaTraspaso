﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ventaRT.Constantes.View
{
    public static class registro
    {
        public const string txt_numoc = "txt_numoc";
        public const string txt_fechac = "txt_fechac";
        public const string txt_fechav = "txt_fechav";
        public const string txt_estado = "txt_estado";
        public const string txt_idtv = "txt_idtv";
        public const string txt_idtr = "txt_idtr";
        public const string txt_idvend = "txt_idvend";
        public const string txt_vend = "txt_vend";
        public const string txt_idcli = "txt_idcli";
        public const string txt_cli = "txt_cli";
        public const string txt_idaut = "txt_idaut";
        public const string txt_aut = "txt_aut";
        public const string cbnd = "cbnd";
        public const string grid = "grid";
        public const string mtx = "mtx";
        public const string txt_com = "txt_com";
        public const string txt_log = "txt_log";
        public const string btn_crear = "1";
        public const string btn_cancel = "2";
        public const string btn_autorizar = "btnAut";
        public const string btn_TR = "btnTR";
        public const string btn_cancelar = "btnCan";
        public const string btn_TV = "btnTV";
    }


    public static class ousr
    {
        public const string uId = "\"USERID\"";
        public const string uCode = "\"USER_CODE\"";
        public const string uName = "\"U_NAME\"";
        public const string uLocked = "\"Locked\"";
        public const string OUSR = "OUSR";
    }    // Maestro de Usuarios

    public static class oitm
    {
        public const string OITM = "OITM";
        public const string ItemCode = "\"ItemCode\"";
        public const string ItemName = "\"ItemName\"";
        public const string AvgPrice = "\"AvgPrice\"";


    }    // Maestro de Articulos

    public static class ocrd
    {
        public const string CardCode = "\"CardCode\"";
        public const string CardName = "\"CardName\"";
        public const string CardType = "\"CardType\"";
        public const string validFor = "\"validFor\"";
        public const string OCRD = "OCRD";
    }    // Maestro de Socios de Negocios

    public static class oitw
    {
        public const string OITW = "OITW";
        public const string ItemCode = "\"ItemCode\"";
        public const string WhsCode = "\"WhsCode\"";
        public const string OnHand = "\"OnHand\"";
    }   // Existencias de articulos pr almacenes

    public static class owhs  // Maestro de Almacenes o Bodegas
    {
        public const string OWHS = "OWHS";
        public const string WhsCode = "\"WhsCode\"";
        public const string WhsName = "\"WhsName\"";
    }

    public static class owtr  // Documentos de Transferencias de Inventarios
    {
        public const string OWTR = "OWTR";
        public const string DocEntry = "\"DocEntry\"";
        public const string DocNum = "\"DocNum\"";
    }

    public static class wtr1  // Lineas de Documentos de Transferencias de Inventarios
    {
        public const string WTR1 = "WTR1";
        public const string DocEntry = "\"DocEntry\"";
        public const string ItemCode = "\"ItemCode\"";
        public const string Quantity = "\"Quantity\"";
        public const string ItemDescription = "\"Dscription\"";
    }


    public static class CAB_RVT
    {
        public const string CAB_RV = "\"@CAB_RSTV\"";
        public const string Code = "\"Code\"";
        public const string Name = "\"Name\"";
        public const string U_numDoc = "\"U_numDoc\"";
        public const string U_fechaC = "\"U_fechaC\"";
        public const string U_fechaV = "\"U_fechaV\"";
        public const string U_estado = "\"U_estado\"";
        public const string U_codCli = "\"U_codCli\"";
        public const string U_cliente = "\"U_cliente\"";
        public const string U_amount = "\"U_amount\"";
        public const string U_idTR = "\"U_idTR\"";
        public const string U_idTV = "\"U_idTV\"";
        public const string U_comment = "\"U_comment\"";
        public const string U_idVend = "\"U_idVend\"";
        public const string U_idAut = "\"U_idAut\"";
        public const string U_vend = "\"U_vend\"";
        public const string U_aut = "\"U_aut\"";
        public const string U_logs = "\"U_logs\"";
    }

    public static class DET_RVT
    {
        public const string DET_RV = "\"@DET_RSTV\"";
        public const string U_numOC = "\"U_numOC\"";
        public const string U_codArt = "\"U_codArt\"";
        public const string U_cant = "\"U_cant\"";
        public const string U_price = "\"U_price\"";
        public const string U_amount = "\"U_amount\"";
        public const string U_onHand = "\"U_onHand\"";
        public const string U_estado = "\"U_estado\"";
        public const string U_idTV = "\"U_idTV\"";
        public const string U_articulo = "\"U_articulo\"";
        public const string Code = "\"Code\"";
    }

    public static class CAB_RVTabla
    {
        public const string CAB_RV = "@CAB_RSTV";
        public const string Code = "Code";
        public const string U_numDoc = "U_numDoc";
        public const string U_fechaC = "U_fechaC";
        public const string U_fechaV = "U_fechaV";
        public const string U_estado = "U_estado";
        public const string U_codCli = "U_codCli";
        public const string U_cliente = "U_cliente";
        public const string U_amount = "U_amount";
        public const string U_idTR = "U_idTR";
        public const string U_idTV = "U_idTV";
        public const string U_comment = "U_comment";
        public const string U_idVend = "U_idVend";
        public const string U_idAut = "U_idAut";
        public const string U_logs = "U_logs";

    }

    public static class DET_RVTabla
    {
        public const string DET_RV = "@DET_RSTV";
        public const string U_numOC = "U_numOC";
        public const string U_codArt = "U_codArt";
        public const string U_price = "U_price";
        public const string U_amount = "U_amount";
        public const string U_onHand = "U_onHand";
        public const string U_cant = "U_cant";
        public const string U_estado = "U_estado";
        public const string U_idTV = "U_idTV";
        public const string Code = "Code";
        public const string articulo = "U_articulo";
    }

}
