using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
using System.Collections;

using SSIFramework;
using System.Reflection;
using SSIFramework.Utilidades;
using System.ComponentModel;
using System.Data;
using SSIFramework.UI.UIApi;
using ventaRT.Model;
using SAPbobsCOM;

namespace ventaRT.Importacion
{
   
    class Imp_Nominas 
    {
       
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        int  ExtraerSerie;

        private string  cuentaPT = "";
        private string cuentaRestos = "";
        public class Trabajador
        {
            public string SeguroConvenio;
            public string CodigoTrabajador;
            public string NombreTrabajador;
            public string CodigoCentro;
            public string cEmb;
            public string cPrest;
            public string cDifNom;
            public string DesPagoEX;
            public string cValor;
            public string cValor2;
            public string ccCocotiComunesti;
            public string ccotiFormaciones;
            public string ccotiDmpleo;
            public string cIrpf;
            public string cTLiquido;
            public string cDinerada;


        }
    

            //SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ContabilizacionDeNominas.Constantes.Views.PantallaImpoNominas.TxtFechaContab).Specific;

        //DateTime dt2 = DateTime.Parse(txt.String);
        ////DateTime dt3 = Convert.ToDateTime(dt2);
        //string s3 = dt2.ToString("dd-MM-yyyy");
        //DateTime dtnew2 = DateTime.Parse(s3);
        //cargarCombo(dtnew2);







        VIEW.PantallaImportacion oPantallaImp;
        public bool GestionarImportacionNominas(string sFichero, VIEW.PantallaImportacion oPantalla, ref List<String> ResultadoDocumentos)
        {

            oPantallaImp = oPantalla;

            oPantallaImp.CrearGrid();
            B1.Application.SetStatusBarMessage("Obteniendo Excel", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            bool EncontradoError = false;
            //Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(
            //sFichero, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);

            //Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;




            //Excel.Application ExcelApp = new Excel.Application();
            //ExcelApp.Visible = true;   




            int resultado = 0;
       

                     try{
                      
                    }
                    catch (Exception ex)
                    {
                        //B1.Company.GetLastError(out Error, out errMsg);
                        //String Mensaje = String.Format(errMsg);
                        //oPantallaImp.AnotarEnGridError(errMsg, Error.ToString(), sCentro);
                        ////oPantallaImp.AnotarEnGridError(Trabajadores.ToString().PadLeft(4, '0'), Mensaje, "", "");
                        //continue;

                        throw ex;



                    }
                    finally
                    {
                        GC.Collect();


                    
                



                /*----------------------------------------------asiento resumen-----------------------------*/



               
                       

                   

             
                //  oProgressBar.Stop();
            }
          
            return EncontradoError;
        }

        private void actualizarFechaCentroC(string codTraba, IProfitCenter oProfitCenter, IProfitCentersService oProfitCentersService)
        {

            oProfitCenter.CenterCode =codTraba;
            oProfitCenter.Effectivefrom = new DateTime(2000, 1, 1); ;

            oProfitCentersService.UpdateProfitCenter((SAPbobsCOM.ProfitCenter)oProfitCenter);

        }

       

        private string obtenerCuentaRestos()
        {
            try
            {
                SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string consultaPte = string.Format("select {0} from {1}",
                ventaRT.Constantes.Views.puenteConsulta.SSICtaDesc,
                ventaRT.Constantes.Views.puenteConsulta.UOADM);
                oRecordset.DoQuery(consultaPte);
                cuentaRestos = oRecordset.Fields.Item("U_SSICtaDesc").Value;
                if (cuentaRestos != null)
                {
                    return cuentaRestos;
                }
                else
                {

                    return "0";
                }
            }
            catch (Exception ex) { throw ex; }

        }
   

        private string buscarCtaPte()
        {
            //rrevisar que tome el dato de la cuenta puente

            try
            {
                SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string consultaPte = string.Format("select {0} from {1}",
                ventaRT.Constantes.Views.puenteConsulta.ctaPteQry,
                ventaRT.Constantes.Views.puenteConsulta.UOADM);
                oRecordset.DoQuery(consultaPte);
                cuentaPT = oRecordset.Fields.Item("U_SSI_CtaPu").Value;
                if (cuentaPT != null)
                {
                    return cuentaPT;
                }
                else
                {

                    return "0";
                }
            }
            catch (Exception ex) { throw ex; }

             }
           
        public void cargarCombo(DateTime FechaConta)
        {
            try { 



            DateTime PrimeDia = new DateTime(FechaConta.Year, FechaConta.Month, 1);

                string FechaContaForm = PrimeDia.ToString(ventaRT.Constantes.Views.formatos.formatoFechaQrie);
               
                string buscarPeriodo = string.Format("select t1.{0}, t2.{3} from {1} as t1 left join {2} as t2 on t1.{3}=t2.{3} where t1.{4} = '30' and t1.{5}='10' and t2.{6}='{7}'",
                    ventaRT.Constantes.Views.SerieConsulta.SeriesName,
                    ventaRT.Constantes.Views.SerieConsulta.NNM1,
                    ventaRT.Constantes.Views.SerieConsulta.OFPR,
                    ventaRT.Constantes.Views.SerieConsulta.Indicator,
                    ventaRT.Constantes.Views.SerieConsulta.ObjectCode,
                    ventaRT.Constantes.Views.SerieConsulta.GroupCode,
                    ventaRT.Constantes.Views.SerieConsulta.F_RefDate,
                    FechaContaForm);

               SAPbouiCOM.ComboBox oComboSerie = B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.PantallaImpoNominas.CmbSerie).Specific;
             
                SSIFramework.Utilidades.GenericFunctions.fillComboBySQL(ref oComboSerie, buscarPeriodo, "SeriesName", "Indicator", true);
                //   string ObtenerIndicador = oComboSerie.Value.ToString();
                oComboSerie.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription;

                /*   if (ObtenerIndicador!="") {

                       ObtenerIndicarAsiento(ObtenerIndicador,FechaContaForm);



                   }
               */

            }
            catch (Exception ex){ throw ex; }
        }

        public int  ObtenerIndicarAsiento(string obtenerIndicador, DateTime FechaContaForm)
        {

            DateTime PrimeDia = new DateTime(FechaContaForm.Year, FechaContaForm.Month, 1);

            string s4 = PrimeDia.ToString(ventaRT.Constantes.Views.formatos.formatoFechaQrie);
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            string BuscarGrupo = string.Format("select  t1.{0} from {1} as t1 left join {2} as t2  on t1.{3}=t2.{3} where t1.{7}='{4}' and t2.{5}='{6}' ",
              ventaRT.Constantes.Views.SerieConsulta.Series,
              ventaRT.Constantes.Views.SerieConsulta.NNM1,
              ventaRT.Constantes.Views.SerieConsulta.OFPR,
              ventaRT.Constantes.Views.SerieConsulta.Indicator,
              obtenerIndicador,
              ventaRT.Constantes.Views.SerieConsulta.F_RefDate,
                s4,
                 ventaRT.Constantes.Views.SerieConsulta.SeriesName

                       );

            oRecordset.DoQuery(BuscarGrupo);
            ExtraerSerie = oRecordset.Fields.Item("Series").Value;
            return ExtraerSerie;

        }




        //VIEW.PantallaImportacionpuente oPantallaImp2;
        public bool GestionarImportacionPuente(string sFichero, VIEW.PantallaImportacionpuente oPantalla, ref List<String> ResultadoDocumentos)
        {

        //    oPantallaImp2 = oPantalla;

        //    oPantallaImp.CrearGrid();
        //    B1.Application.SetStatusBarMessage("Obteniendo Excel", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

        //    bool EncontradoError = false;
        //    Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(
        //    sFichero, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);

        //    Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;


        //    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);
        //    Microsoft.Office.Interop.Excel.Range xlRange = worksheet.UsedRange;

        //    SAPbouiCOM.ProgressBar oProgressBar = B1.Application.StatusBar.CreateProgressBar("Cargando datos del Excel ...", xlRange.Rows.Count + 1, false);
        //    for (int x = 1; x <= sheets.Count; x = x + 2)
        //    {
        //        for (int i = 9; i <= xlRange.Rows.Count; i++)
        //        {



        //            //siempre comenzara EN EL 9 DE LA COLUMNA 2?
        //            #region Cuenta
        //            String Cuenta = "";
        //            string otraCelda;
        //            try
        //            {


        //                Cuenta = xlRange.Cells[i, 2].Value2.ToString();
        //                if (Cuenta == "TOTAL")
        //                {

        //                    break;
        //                }
        //                else
        //                {
        //                    otraCelda = xlRange.Cells[i, 3].value2.ToString();
        //                    continue;
        //                }

        //            }
        //            catch (Exception)
        //            {

        //                EncontradoError = true;
        //                String Mensaje = String.Format("Pasando a la siguiente pagina");
        //                oPantallaImp.AnotarEnGridError(Mensaje, "");
        //                //oPantallaImp.AnotarEnGridError(i.ToString().PadLeft(4, '0'), Mensaje,  "", "");

        //            }
        //            #endregion
        //            try
        //            {
        //                oProgressBar.Text = "Cargando datos del Excel ...";
        //            }
        //            catch (Exception)
        //            {
        //                oProgressBar = B1.Application.StatusBar.CreateProgressBar("Cargando datos del Excel ...", xlRange.Rows.Count + 1, false);
        //            }
        //            oProgressBar.Value = i;
        //        }
        //    }
        //    oProgressBar.Stop();
        //    if (!EncontradoError)
        //    {
        //        //Imp.Add();
        //        B1.Application.SetStatusBarMessage("Creando Asientos en SAP ...", SAPbouiCOM.BoMessageTime.bmt_Long, false);
        //        ResultadoDocumentos = Imp.AddToSAPFactura();

        //    }
        //    theWorkbook.Close(false);
        //    ExcelObj.Quit();
          return true;
        }

       

    }

  







}
