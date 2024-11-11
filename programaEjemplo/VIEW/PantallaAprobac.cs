using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.Globalization;
using SAPbouiCOM;
using SAPbobsCOM;
using SSIFramework;
using SSIFramework.DI.Attributes;
using SSIFramework.Utilidades;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;



namespace ventaRT.VIEW
{
    class PantallaAprobac
        : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        private List<string> lineasnodisp = new List<string>();

        private string formActual = "";
        SAPbouiCOM.Form AForm = null;
        SAPbouiCOM.Matrix AMatrix = null;
        SAPbouiCOM.Button btnRev = null;
        SAPbouiCOM.Button btnExcel = null;
        String CompanyName = "";
        private int row = 1;
        String ExcelPath = "";
        private bool validPathExcel = false;
        private Excel._Application excelApp = null;
        private Excel._Workbook workBook = null;
        private Excel._Worksheet workSheet = null;
        private Excel._Worksheet workSheet2 = null;

        public PantallaAprobac()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Aprobac.srf"], "AprRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString())
        {

            formActual = "AprRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();

            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
            string errorMessage = "";
            //errorMessage = cancelar_vencidaspormas10D();
            //if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

            errorMessage = GetConfigSociety();
            if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
            errorMessage = cargar_datos_iniciales();
            if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
            errorMessage = cargar_datos_matriz();
            if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }


        }

        private void HandleError(Exception ex, bool checkExcel = false)
        {
            if (checkExcel)
            {
                if (workSheet != null)
                {
                    Marshal.ReleaseComObject(workSheet);
                }
                if (workSheet2 != null)
                {
                    Marshal.ReleaseComObject(workSheet2);
                }
                   if (workBook != null)
                {
                    workBook.Close();
                    Marshal.ReleaseComObject(workBook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                workSheet = null;
                workSheet2 = null;
                workBook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            if (B1.Company.InTransaction) {B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);}
            string msgError = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
            B1.Application.SetStatusBarMessage("Error: " + msgError, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
        }

        private string cargar_datos_iniciales()
        {
            string errorMessage = "";
            try
            {
                SAPbouiCOM.CheckBox cboxNue = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxNue).Specific;
                cboxNue.Checked = true;

                SAPbouiCOM.CheckBox cboxApr = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxApr).Specific;
                cboxApr.Checked = true;

                //SAPbouiCOM.CheckBox cboxTra = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxTra).Specific;
                //cboxTra.Checked = true;

                SAPbouiCOM.CheckBox cboxCan = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCan).Specific;
                cboxCan.Checked = true;

                SAPbouiCOM.CheckBox cboxDev = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxDev).Specific;
                cboxDev.Checked = true;

                AForm = B1.Application.Forms.ActiveForm;
                AForm.EnableMenu("1290", false); AForm.EnableMenu("1289", false);
                AForm.EnableMenu("1288", false); AForm.EnableMenu("1291", false);
                AForm.EnableMenu("1282", false);   // crear
                AForm.EnableMenu("1281", false);  //buscar
                AForm.EnableMenu("1283", false);  //eliminar
                AForm.EnableMenu("1292", false);  //buscar
                AForm.EnableMenu("1293", false);  //eliminar

                btnExcel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.aprobac.btn_exp).Specific;
                btnExcel.Item.Enabled = validPathExcel;
                btnExcel.Image = System.Environment.CurrentDirectory + "\\img\\excelbtn.jpg";

                btnRev = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.aprobac.btn_rev).Specific;
                btnRev.Image = System.Environment.CurrentDirectory + "\\img\\atrasbtn.jpg";

                if (!btnExcel.Item.Enabled)
                {
                   errorMessage = "Configurar Sociedad: No se puede ejecutar este Addon en esta Sociedad porque no tiene configurado correctamente el Directorio para Excel ..";
                   return errorMessage;
                }

            }
            catch (Exception ex)
            {
                errorMessage = "Error cargando datos iniciales: " + 
                    ((B1.Company.GetLastErrorCode() != 0) 
                    ? B1.Company.GetLastErrorDescription() 
                    : ex.Message);
             }
            return errorMessage;
        }

        public string cargar_datos_matriz()
        {
            string errorMessage = "";
            try
            {   
                //B1.Application.SetStatusBarMessage("Cargando datos de Solicitudes de Reservas de Stock para su Autorización...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                formActual = B1.Application.Forms.ActiveForm.UniqueID;
                AForm = B1.Application.Forms.ActiveForm;
                AMatrix = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item("mtxaprob").Specific;
                string filtrado = "";
                B1.Application.Forms.ActiveForm.Freeze(true);

                SAPbouiCOM.CheckBox cboxNue = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxNue).Specific;
                SAPbouiCOM.CheckBox cboxApr = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxApr).Specific;
                //SAPbouiCOM.CheckBox cboxTra = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxTra).Specific;
                SAPbouiCOM.CheckBox cboxCan = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCan).Specific;
                SAPbouiCOM.CheckBox cboxDev = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxDev).Specific;

                string SQLQuery = string.Empty;

                SAPbouiCOM.CheckBox cboxPer = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxPer).Specific;
                string condPer = String.Empty;
                if (cboxPer.Checked == true)
                {
                    SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                    SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                    condPer = desde.Value!="" && hasta.Value != "" ?Constantes.View.CAB_RVT.U_fechaC + " between '" + desde.Value + "' AND '" + hasta.Value + "'":"" ;
                    filtrado = filtrado + (condPer != "" ? "\n>Período Seleccionado: " + 
                        DateTime.ParseExact(desde.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy")+
                        " a "+
                        DateTime.ParseExact(hasta.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy") : "");
                }   

                SAPbouiCOM.CheckBox cboxCli = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCli).Specific;
                string condCli = String.Empty;
                if (cboxCli.Checked == true)
                {
                    SAPbouiCOM.ComboBox cli = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbCli).Specific;
                    string selCli = (cli.Value != "") ? cli.Value : "";
                    condCli = selCli != "" ? Constantes.View.CAB_RVT.U_codCli + " = '" + selCli + "'" : condCli;
                    filtrado = filtrado + (condCli != ""? "\n>Cliente Seleccionado: " + selCli:"");
                }

                SAPbouiCOM.CheckBox cboxArt = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxArt).Specific;
                string condArt = String.Empty;
                if (cboxArt.Checked == true)
                {
                    SAPbouiCOM.ComboBox art = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbArt).Specific;
                    string selArt = (art.Value != "") ? art.Value : "";
                    condArt = selArt != "" ? Constantes.View.DET_RVT.U_codArt + " = '" + selArt + "'": condArt;
                    filtrado = filtrado + (condArt != ""? "\n>Artículo Seleccionado: "+ selArt: "");
                }

                SAPbouiCOM.CheckBox cboxVend = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxVend).Specific;
                string condVend = String.Empty;
                if (cboxVend.Checked == true)
                {
                    SAPbouiCOM.ComboBox vend = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbVend).Specific;
                    string selVend = (vend.Value != "" )? vend.Value : "";
                    condVend = selVend != "" ? Constantes.View.CAB_RVT.U_idVend + " = '" + selVend + "'" : condVend;
                    filtrado = filtrado + (condVend != "" ? "\n>Vendedor Seleccionado: " + selVend: "");

                }

                string condNue = String.Empty;
                condNue = (cboxNue.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'R' " : condNue;

                string condApr = String.Empty;
                condApr = (cboxApr.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'A' " : condApr;

                //string condTra = String.Empty;
                //condTra = (cboxTra.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'T' " : condTra;

                string condCan = String.Empty;
                condCan = (cboxCan.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'C' " : condCan;

                string condDev = String.Empty;
                condDev = (cboxDev.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'D' " : condDev;

                filtrado = filtrado + "\n>Estados Seleccionados: " +
                        (condNue == String.Empty ? "Reservada" : "") + (condApr == String.Empty ? "-Aprobada" : "") + 
                        //(condTra == String.Empty ? "-Transferida" : "")+
                        (condCan == String.Empty ? "-Cancelada " : "") + (condDev == String.Empty ? "-Devuelta" : "") ;

                //if (filtrado != "")
                //{
                //    int respuesta = B1.Application.MessageBox("Filtros aplicados: " + filtrado, 1, "OK");
                //}

                B1.Application.SetStatusBarMessage("Cargando datos de Solicitudes...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                string cadw = "";
                cadw = condPer != String.Empty || condCli != String.Empty || condArt != String.Empty || condVend != String.Empty ||
                       condNue != String.Empty || condApr != String.Empty || condCan != String.Empty  || condDev != String.Empty
                       ? " WHERE " : "";
                cadw = cadw + (condPer  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condPer  : "");
                cadw = cadw + (condCli  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCli  : "");
                cadw = cadw + (condArt  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condArt  : "");
                cadw = cadw + (condVend != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condVend : "");
                cadw = cadw + (condNue  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condNue  : "");
                cadw = cadw + (condApr  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condApr : "");
                //cadw = cadw + (condTra  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condTra : "");
                cadw = cadw + (condCan  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCan : "");
                cadw = cadw + (condDev  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condDev : "");

                string adicjoin = (condArt != String.Empty) ? " INNER JOIN " +
                Constantes.View.DET_RVT.DET_RV + " T3 ON T0." + Constantes.View.CAB_RVT.U_numDoc +
                " = T3."  + Constantes.View.DET_RVT.U_numOC : "";

                string adicgroup = (condCli != String.Empty || condArt != String.Empty) ? " GROUP BY " +
                 " T0." + Constantes.View.CAB_RVT.U_numDoc +", T0." + Constantes.View.CAB_RVT.U_idVend +" , T1." +
                Constantes.View.ousr.uName + ", T0."  +
                Constantes.View.CAB_RVT.U_codCli + ", T0." +
                Constantes.View.CAB_RVT.U_cliente + ", T0." + 
                Constantes.View.CAB_RVT.U_fechaC + ", T0." +
                Constantes.View.CAB_RVT.U_fechaV + ", DAYS_BETWEEN( CURRENT_DATE,T0."  + 
                Constantes.View.CAB_RVT.U_fechaV +" ), T0." + Constantes.View.CAB_RVT.U_idAut +", T2." +
                Constantes.View.ousr.uName +", T0." + Constantes.View.CAB_RVT.U_estado + ", T0." +
                Constantes.View.CAB_RVT.U_idTR +" , T0."  +Constantes.View.CAB_RVT.U_idTV +
                ", T0." + Constantes.View.CAB_RVT.U_amount
                //+",  T0." +   Constantes.View.CAB_RVT.U_comment
                : "" ;              

                SQLQuery = String.Format("SELECT T0.{1} , T0.{4}, T1.{3} U_vend, T0.{6}, T0.{7}, DAYS_BETWEEN( CURRENT_DATE,T0.{7}) U_diasv, " +
                      " T0.{8}, T2.{3} U_aut, T0.{9}, T0.{10}, T0.{11}, T0.{16}, T0.{17}, T0.{26}, CAST(T0.{1} AS INT) AS ND" +
                      //, T0.{22} " +
                      " FROM {0} T0 INNER JOIN {2} T1 ON T0.{4} = T1.{5} " +
                      " LEFT JOIN {2} T2 ON T0.{8} = T2.{5} {24}  {23}  {25}  ORDER BY CAST(T0.{1} AS INT) ",
                                              Constantes.View.CAB_RVT.CAB_RV, //0
                                              Constantes.View.CAB_RVT.U_numDoc,//1
                                              Constantes.View.ousr.OUSR, //2
                                              Constantes.View.ousr.uName, //3
                                              Constantes.View.CAB_RVT.U_idVend,//4
                                              Constantes.View.ousr.uCode, //5
                                              Constantes.View.CAB_RVT.U_fechaC, //6
                                              Constantes.View.CAB_RVT.U_fechaV, //7
                                              Constantes.View.CAB_RVT.U_idAut, //8
                                              Constantes.View.CAB_RVT.U_estado, //9
                                              Constantes.View.CAB_RVT.U_idTR, //10
                                              Constantes.View.CAB_RVT.U_idTV, //11
                                              Constantes.View.DET_RVT.DET_RV, //12
                                              Constantes.View.DET_RVT.U_numOC, //13
                                              Constantes.View.DET_RVT.U_codArt, //14
                                              Constantes.View.DET_RVT.U_articulo, //15
                                              Constantes.View.CAB_RVT.U_codCli, //16
                                              Constantes.View.CAB_RVT.U_cliente, //17
                                              Constantes.View.DET_RVT.U_cant, //18
                                              Constantes.View.DET_RVT.U_onHand, //19
                                              Constantes.View.DET_RVT.U_estado, //20
                                              Constantes.View.DET_RVT.U_idTV, //21                                                
                                              Constantes.View.CAB_RVT.U_comment,//22,
                                              cadw, //23,
                                              adicjoin, //24
                                              adicgroup, //25
                                              Constantes.View.CAB_RVT.U_amount //26
                                             
                                              );

                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(SQLQuery);
                AMatrix.Clear();
                SAPbobsCOM.Fields fields = rsCards.Fields;
                rsCards.MoveFirst();
                for (int i = 1; !rsCards.EoF; i++)
                {
                    AMatrix.AddRow(1);
                    AMatrix.Columns.Item(1).Cells.Item(i).Specific.Value = fields.Item("U_numDoc").Value.ToString();
                    AMatrix.Columns.Item(2).Cells.Item(i).Specific.Value = fields.Item("U_IdVend").Value.ToString();
                    AMatrix.Columns.Item(3).Cells.Item(i).Specific.Value = fields.Item("U_vend").Value.ToString();
                    AMatrix.Columns.Item(4).Cells.Item(i).Specific.Value = fields.Item("U_fechaC").Value.ToString("yyyyMMdd");
                    AMatrix.Columns.Item(5).Cells.Item(i).Specific.Value = fields.Item("U_fechaV").Value.ToString("yyyyMMdd");
                    AMatrix.Columns.Item(6).Cells.Item(i).Specific.Value = fields.Item("U_diasv").Value.ToString();
                    if (Int32.Parse(fields.Item("U_diasv").Value.ToString()) < 0)
                    {
                        AMatrix.Columns.Item(6).Cells.Item(i).Specific.Value = 0;
                        AMatrix.CommonSetting.SetCellFontColor(i, 6, 255);
                    }
                    else
                    {
                        if (Int32.Parse(fields.Item("U_diasv").Value.ToString()) <= 5)
                        {
                            AMatrix.CommonSetting.SetCellFontColor(i, 6, 255);
                        }
                        else
                        {
                            AMatrix.CommonSetting.SetCellFontColor(i, 6, 0);
                        }
                    }

                    AMatrix.Columns.Item(7).Cells.Item(i).Specific.Value = fields.Item("U_IdAut").Value.ToString();
                    AMatrix.Columns.Item(8).Cells.Item(i).Specific.Value = fields.Item("U_aut").Value.ToString();
                    AMatrix.Columns.Item(9).Cells.Item(i).Specific.Value = fields.Item("U_codCli").Value.ToString();
                    AMatrix.Columns.Item(10).Cells.Item(i).Specific.Value = fields.Item("U_cliente").Value.ToString();


                    string txtestado = fields.Item("U_estado").Value.ToString();
                    txtestado = txtestado.Substring(0, 1);
                    SAPbouiCOM.ComboBox mc = (SAPbouiCOM.ComboBox)AMatrix.Columns.Item(11).Cells.Item(i).Specific;

                    mc.Select(txtestado,BoSearchKey.psk_ByValue);

                    if (txtestado=="C" )
                    {
                        AMatrix.CommonSetting.SetCellFontColor(i, 11, 255); 
                    }
                    else
                        if (txtestado == "A" || txtestado == "D")
                        {
                            AMatrix.CommonSetting.SetCellFontColor(i, 11, 000102000);
                        }
                        else { AMatrix.CommonSetting.SetCellFontColor(i, 11, 0); }


                    AMatrix.Columns.Item(12).Cells.Item(i).Specific.Value = fields.Item("U_amount").Value.ToString();
                    AMatrix.Columns.Item(13).Cells.Item(i).Specific.Value = obtener_DocNum(fields.Item("U_IdTR").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                    AMatrix.Columns.Item(14).Cells.Item(i).Specific.Value = obtener_DocNum(fields.Item("U_idTV").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                    AMatrix.Columns.Item(15).Cells.Item(i).Specific.Value = obtener_Comentario(fields.Item("U_numDoc").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                    rsCards.MoveNext();
                    if (!rsCards.EoF)
                    { B1.Application.SetStatusBarMessage("ESPERE......Cargando datos de Solicitud...No." + fields.Item("U_numDoc").Value.ToString() + "    (" + i.ToString() + "/" + rsCards.RecordCount.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, false); }
                    else
                    { B1.Application.SetStatusBarMessage(" ", SAPbouiCOM.BoMessageTime.bmt_Short, false); }
                    
                }
                AMatrix.AutoResizeColumns();
                B1.Application.Forms.ActiveForm.Freeze(false);
            }

            catch (Exception ex)
            {
                errorMessage = "Cargar Solicitudes: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return errorMessage;
        }

        private void ThisSapApiForm_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorMessage = "";
            try
            {
                if (FormUID == formActual)
                {
                    if (!pVal.BeforeAction)
                    {
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                {
                                    // Aquí puedes desasociar los eventos si es necesario
                                    this.B1.Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
                                }
                                break;

                            case BoEventTypes.et_VALIDATE:
                                {
                                    if (pVal.ItemUID == "txtDesde" || pVal.ItemUID == "txtHasta")
                                    {
                                        SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                                        SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                                        if (desde.Value.ToString() != "" && hasta.Value.ToString() != "")
                                        {
                                            DateTime tempdate;
                                            DateTime tempdate2;
                                            if (DateTime.TryParseExact(desde.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out tempdate) &&
                                                DateTime.TryParseExact(hasta.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out tempdate2))
                                            {

                                                if ((DateTime.ParseExact(hasta.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture) -
                                                        DateTime.ParseExact(desde.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)).Days < 0)
                                                {
                                                    B1.Application.SetStatusBarMessage("Error: Desde <= Hasta", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                    BubbleEvent = true;
                                                }
                                                else {
                                                    errorMessage = cargar_datos_matriz();
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                }
                                            }
                                            else
                                            {
                                                B1.Application.SetStatusBarMessage("Error: Desde o Hasta no poseen formato adecuado..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                BubbleEvent = true;
                                            }
                                        }
                                        else
                                        {
                                            B1.Application.SetStatusBarMessage("Error: Desde o Hasta no poseen formato adecuado..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                            BubbleEvent = true;
                                        }
                                    }

                                }
                                break;

                            case BoEventTypes.et_COMBO_SELECT:
                                {
                                    switch (pVal.ItemUID)
                                    {
                                        case "cbCli":
                                            {
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxCli").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbCli").Specific;

                                                oCbox.Checked = oCombo.Value != "";
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                //if (oCombo.Value != "") 
                                                //{
                                                //    if (oCbox.Checked) { 
                                                //        errorMessage = cargar_datos_matriz();
                                                //        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                //     }
                                                //    else { oCbox.Checked = true; }
                                                //}
                                                //else { oCbox.Checked = false; }
                                            }
                                            break;
                                        case "cbArt":
                                            {
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxArt").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbArt").Specific;

                                                oCbox.Checked = oCombo.Value != "";
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                //if (oCombo.Value != "")
                                                //{
                                                //    if (oCbox.Checked)
                                                //    {
                                                //        errorMessage = cargar_datos_matriz();
                                                //        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                //    }
                                                //    else { oCbox.Checked = true; }
                                                //}
                                                //else { oCbox.Checked = false; }
                                            }
                                            break;
                                        case "cbVend":
                                            {
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxVend").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbVend").Specific;

                                                oCbox.Checked = oCombo.Value != ""; 
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

    
                                                //if (oCombo.Value != "")
                                                //{
                                                //    if (oCbox.Checked)
                                                //    {
                                                //        errorMessage = cargar_datos_matriz();
                                                //        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                //    }
                                                //    else { oCbox.Checked = true; }
                                                //}
                                                //else { oCbox.Checked = false; }
                                            }
                                            break;
                                    }
                                    break;
                                }

                            case BoEventTypes.et_ITEM_PRESSED:
                                {
                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.aprobac.btn_exit:
                                            {
                                                SAPbouiCOM.Form oForm = B1.Application.Forms.ActiveForm;
                                                oForm.Close();
                                            }
                                            BubbleEvent = true;
                                            break;

                                        case "cboxPer":
                                            {
                                                // Activar busqueda por periodo  
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxPer").Specific;
                                                SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                                                SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                                                if (desde.Value.ToString() != "" && hasta.Value.ToString() != "")
                                                {
                                                    errorMessage = cargar_datos_matriz();
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                }
                                            }
                                            break;

                                        case "cboxArt":
                                            {
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxArt").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbArt").Specific;
                                                if (oCombo.Value != "")
                                                {
                                                    errorMessage = cargar_datos_matriz();
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
 
                                                }                                                
                                            }
                                            break;

                                        case "cboxVend":
                                            {
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxVend").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbVend").Specific;
                                                if (oCombo.Value != "")
                                                {
                                                    errorMessage = cargar_datos_matriz();
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                 }
                                            }
                                            break;

                                        case "cboxCli":
                                            {
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxCli").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbCli").Specific;
                                                if (oCombo.Value != "")
                                                {
                                                    errorMessage = cargar_datos_matriz();
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                 }
                                            }
                                            break;

                                        case "cboxNue":
                                            {
                                                // Desactivar el estado Nueva  
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                             }
                                            break;

                                        case "cboxApr":
                                            {
                                                // Desactivar el estado Autorizada
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                             }
                                            break;


                                        case "cboxCan":
                                            {
                                                // Desactivar el estado Cancelada  
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                             }
                                            break;

                                        case "cboxDev":
                                            {
                                                // Desactivar el estado Devuelta  
                                                errorMessage = cargar_datos_matriz();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                            }
                                            break;
                                    }
                                    break;

                                }
                        }

                    }
                    else
                    {
                        // Antes de Accion

                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                {
                                    // Mensaje de error para los combobox cdo se editan los campos e fecha q SAP no los activa
                                    if (pVal.ItemUID == "cboxDev" || pVal.ItemUID == "cboxCan" ||
                                        pVal.ItemUID == "cboxApr" || pVal.ItemUID == "cboxNue" || pVal.ItemUID == "cboxCli" ||
                                        pVal.ItemUID == "cboxArt" || pVal.ItemUID == "cboxVend")
                                    {
                                        SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtDesde").Specific;
                                        SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtHasta").Specific;
                                        SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxPer").Specific;

                                        if ((oCbox.Checked && (desde.Value != "" || hasta.Value != "")) &&
                                            (AForm.ActiveItem == "txtHasta" || AForm.ActiveItem == "txtDesde"))
                                        {
                                            B1.Application.SetStatusBarMessage("Error: Finalice la edición primero de " + (AForm.ActiveItem == "txtDesde" ? "Desde" : "Hasta") + 
                                                ". Si finalizó, salga con " + (AForm.ActiveItem == "txtDesde" ? "<SHIFT+TAB>" : "<TAB>")  , SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                            BubbleEvent = false;
                                        }
                                    }
                                    else
                                    {
                                        switch (pVal.ItemUID)
                                        {
                                            case "cbCli":
                                                {
                                                    // Rellenando combo de busqueda
                                                    SAPbouiCOM.ComboBox oCombo = null;
                                                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbCli").Specific;
                                                    string SQLQuery = string.Empty;
                                                    SQLQuery = String.Format("SELECT {1}, {2} FROM {0} GROUP BY {1}, {2} ORDER BY {1}",
                                                                                        Constantes.View.CAB_RVT.CAB_RV,
                                                                                        Constantes.View.CAB_RVT.U_codCli,
                                                                                        Constantes.View.CAB_RVT.U_cliente);

                                                    errorMessage = llenar_combo_busq(oCombo, SQLQuery);
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                }
                                                break;
                                            case "cbArt":
                                                {
                                                    // Rellenando combo de busqueda
                                                    SAPbouiCOM.ComboBox oCombo = null;
                                                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbArt").Specific;
                                                    string SQLQuery = string.Empty;
                                                    SQLQuery = String.Format("SELECT {1}, {2} FROM {0} GROUP BY {1}, {2} ORDER BY {1}",
                                                                                        Constantes.View.DET_RVT.DET_RV,
                                                                                        Constantes.View.DET_RVT.U_codArt,
                                                                                        Constantes.View.DET_RVT.U_articulo);
                                                    errorMessage = llenar_combo_busq(oCombo, SQLQuery);
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                }
                                                break;

                                            case "cbVend":
                                                {
                                                    // Rellenando combo de busqueda
                                                    SAPbouiCOM.ComboBox oCombo = null;
                                                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbVend").Specific;
                                                    string SQLQuery = string.Empty;
                                                    SQLQuery = String.Format("SELECT {1}, {2} FROM {0} GROUP BY {1}, {2} ORDER BY {1}",
                                                                                        Constantes.View.CAB_RVT.CAB_RV,
                                                                                        Constantes.View.CAB_RVT.U_idVend,
                                                                                        Constantes.View.CAB_RVT.U_vend);
                                                    errorMessage = llenar_combo_busq(oCombo, SQLQuery);
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                }
                                                break;

                                            case "mtxaprob":
                                                {
                                                    if (pVal.Row > 0)
                                                    {
                                                        string nodoc = AMatrix.Columns.Item(1).Cells.Item(pVal.Row).Specific.Value;
                                                        //new VIEW.PantallaRegistro(this, false, nodoc);
                                                        // Verificamos cuántos formularios Registro están abiertos
                                                        if (addonGeneral.contadorRegistrosAbiertos < addonGeneral.maxRegistrosAbiertos)
                                                        {
                                                            // Crear nueva instancia de PantallaRegistro
                                                            PantallaRegistro nuevoRegistro = new PantallaRegistro(this, false, nodoc);
                                                            addonGeneral.contadorRegistrosAbiertos++; // Incrementar contador
                                                        }
                                                        else
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error: No se puede abrir más de 3 formularios de Registro.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        }
                                                    }

                                                }
                                                break;
                                        }
                                    }
                                }
                                break;

                            case BoEventTypes.et_ITEM_PRESSED:
                                {

                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.aprobac.btn_rev:
                                            {
                                                errorMessage = cancelar_vencidaspormas10D();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                break;
                                            }

                                        case Constantes.View.aprobac.btn_exp:
                                            {
                                                errorMessage = exportar();
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage), true); }
                                                break;
                                            }


                                        case Constantes.View.aprobac.cboxPer:
                                            {
                                                if (pVal.InnerEvent == false  )
                                                {

                                                    SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtDesde").Specific;
                                                    SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtHasta").Specific;
                                                    SAPbouiCOM.CheckBox oCbox = null;
                                                    oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxPer").Specific;

                                                    if (oCbox.Checked && (desde.Value == "" || hasta.Value == ""))
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error: No se puede filtrar por Período si Desde o Hasta están vacíos", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        oCbox.Checked = false;
                                                        BubbleEvent = false;
                                                    }
                                                }
                                            }
                                            break;
                                        case Constantes.View.aprobac.cboxVend:
                                            {
                                                if (pVal.InnerEvent == false)
                                                {
                                                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbVend").Specific;
                                                    SAPbouiCOM.CheckBox oCbox =  (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxVend").Specific;

                                                    if (oCbox.Checked && oCombo.Value == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error No se puede filtrar por Vendedor si no tiene Vendedor Seleccionado", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        oCbox.Checked = false;
                                                        BubbleEvent = false;
                                                    }
                                                }
                                            }
                                            break;

                                        case Constantes.View.aprobac.cboxArt:
                                            {
                                                if (pVal.InnerEvent == false)
                                                {
                                                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbArt").Specific;
                                                    SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxArt").Specific;

                                                    if (oCbox.Checked && oCombo.Value == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error No se puede filtrar por Artículo si no tiene Artículo Seleccionado", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        oCbox.Checked = false;
                                                        BubbleEvent = false;
                                                    }
                                                }
                                            }
                                            break;
                                        case Constantes.View.aprobac.cboxCli:
                                            {
                                                if (pVal.InnerEvent == false)
                                                {
                                                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbCli").Specific;
                                                    SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxCli").Specific;

                                                    if (oCbox.Checked && oCombo.Value == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error No se puede filtrar por Cliente si no tiene Cliente Seleccionado", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        oCbox.Checked = false;
                                                        BubbleEvent = false;
                                                    }
                                                }
                                            }
                                            break;
                                    
                                    }
                                }
                                break;

                            case BoEventTypes.et_VALIDATE:
                                {
                                    if (pVal.InnerEvent == false && (pVal.ItemUID == "txtDesde" || pVal.ItemUID == "txtHasta"))
                                    {

                                        SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtDesde").Specific;
                                        SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtHasta").Specific;

                                        if (desde.Value != "" && hasta.Value != "")
                                            if ((DateTime.ParseExact(hasta.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture) -
                                                  DateTime.ParseExact(desde.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)).Days < 0)
                                            {
                                                B1.Application.SetStatusBarMessage("Error Desde <= Hasta", SAPbouiCOM.BoMessageTime.bmt_Long, true);
                                                if (pVal.ItemUID == "txtDesde") { desde.Value = hasta.Value; }
                                                else {hasta.Value = desde.Value;}
                                                BubbleEvent = false;
                                            }
                                    }


                                    if (pVal.InnerEvent == false && (pVal.ItemUID == "cboxPer" || pVal.ItemUID == "txtDesde" || pVal.ItemUID == "txtHasta"))
                                    {

                                        SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtDesde").Specific;
                                        SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtHasta").Specific;
                                        SAPbouiCOM.CheckBox oCbox = null;
                                        oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxPer").Specific;

                                        if (oCbox.Checked && (desde.Value == "" || hasta.Value == ""))
                                        {
                                            B1.Application.SetStatusBarMessage("Error No se puede filtrar por Período si el Desde o Hasta están vacíos, se desmarca Periodo..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                            BubbleEvent = false;
                                        }
                                    }
                                    break;
                                }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
                B1.Application.SetStatusBarMessage("Error: " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                //throw;
            }

        }

        public string llenar_combo_busq(SAPbouiCOM.ComboBox oCombo, string SqlQuery)
        {
            string errorMessage = "";
            try
            {
                SAPbobsCOM.Recordset oRecordSet = null;

                while (oCombo.ValidValues.Count > 0)
                {
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(SqlQuery);
                oCombo.ValidValues.Add("", "");
                for (int i = 1; !oRecordSet.EoF; i++)
                {
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString(), oRecordSet.Fields.Item(1).Value.ToString());
                    oRecordSet.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            catch (Exception ex)
            {
                errorMessage = "Cargar Lista de Solicitudes: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return errorMessage;

        }

        private string obtener_DocNum(string dentry, out string errorMessage)
        {
            errorMessage = "";
            string dnum = "";
            if (dentry != "")
            {
                try
                {
                    String strSQL = String.Format("SELECT {2} FROM {0} Where {1}='{3}'",
                              Constantes.View.owtr.OWTR,
                              Constantes.View.owtr.DocEntry,
                              Constantes.View.owtr.DocNum,
                              dentry);
                    Recordset rsDoc = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsDoc.DoQuery(strSQL);
                    SAPbobsCOM.Fields fields = rsDoc.Fields;
                    rsDoc.MoveFirst();
                    if (!rsDoc.EoF)
                    {
                        dnum = rsDoc.Fields.Item("DocNum").Value.ToString();
                    }
                }

                catch (Exception ex)
                {
                    errorMessage = "Obtener DocNum de la Transferencia: : " +
                        ((B1.Company.GetLastErrorCode() != 0)
                        ? B1.Company.GetLastErrorDescription()
                        : ex.Message);
                }
            }
            return dnum;
        }

        private string obtener_Comentario(string solnum, out string errorMessage)
        {
            errorMessage = "";
            string dcom = "";
            if (solnum != "")
            {
                try
                {

                    String strSQL = String.Format("SELECT {2} FROM {0} Where {1}='{3}'",
                              Constantes.View.CAB_RVT.CAB_RV,
                              Constantes.View.CAB_RVT.U_numDoc,
                              Constantes.View.CAB_RVT.U_comment,
                              solnum);
                    Recordset rsDoc = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsDoc.DoQuery(strSQL);
                    SAPbobsCOM.Fields fields = rsDoc.Fields;
                    rsDoc.MoveFirst();
                    if (!rsDoc.EoF)
                    {
                        dcom = rsDoc.Fields.Item("U_comment").Value.ToString();
                    }
                }
                catch (Exception ex)
                {
                    errorMessage = "Obtener Comentarios de la Solicitud: " +
                        ((B1.Company.GetLastErrorCode() != 0)
                        ? B1.Company.GetLastErrorDescription()
                        : ex.Message);
                }
            }
            return dcom;

        }

        private double obtener_exist_articulo(string codart, string codwhs, out string errorMessage)
        {
            errorMessage = "";
            double exist = 0.00;
            try
            {
                String strSQL = String.Format("SELECT {0} FROM {3} " +
                    " WHERE contains({1},'%{4}%') AND {2}='{5}'  ",
                          Constantes.View.oitw.OnHand,
                          Constantes.View.oitw.ItemCode,
                          Constantes.View.oitw.WhsCode,
                          Constantes.View.oitw.OITW,
                          codart,
                          codwhs);
                Recordset rsResult = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsResult.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsResult.Fields;
                rsResult.MoveFirst();
                if (!rsResult.EoF)
                {
                    exist = Double.Parse(rsResult.Fields.Item("OnHand").Value.ToString());
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Obtener Stock Disponible: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return exist;
        }

        private string revertir(string sCode, string docentry, string sCli, string snCli)
        {
            string errorMessage = "";
            int result = 0;
            try
            {
                GC.Collect();
                B1.Company.StartTransaction();
                SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                doctransf.DocDate = DateTime.Today;
                doctransf.TaxDate = DateTime.Today;
                doctransf.CardCode = sCli;
                doctransf.CardName = snCli;

                // Serie Primaria
                doctransf.Series = 27;

                doctransf.FromWarehouse = "CD_RSV";
                doctransf.ToWarehouse = "CD";
                doctransf.JournalMemo = "Addons VentasRT Canc.Aut. Solic:" + sCode;
                if (docentry != "")
                {
                    String strSQL = String.Format("SELECT {2}, {3}, {4} FROM {0} Where {1}='{5}'",
                                Constantes.View.wtr1.WTR1,
                                Constantes.View.wtr1.DocEntry,
                                Constantes.View.wtr1.ItemCode,
                                Constantes.View.wtr1.ItemDescription,
                                Constantes.View.wtr1.Quantity,
                                docentry);

                    Recordset rsDoc = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsDoc.DoQuery(strSQL);
                    SAPbobsCOM.Fields fields = rsDoc.Fields;
                    rsDoc.MoveFirst();
                    string artcurrent = "";
                    string art = "";
                    double totalart = 0.00;
                    int cantlines = 1;
                    int linestransf = 0;
                    double disponible = 0;
                    for (int i = 1; i <= rsDoc.RecordCount; i++)
                    {
                        art = rsDoc.Fields.Item("ItemCode").Value.ToString();
                        if (artcurrent != art)
                        {
                            if (artcurrent != "")
                            {
                                disponible = obtener_exist_articulo(artcurrent, "CD_RSV", out errorMessage);
                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                if (disponible >= totalart)
                                {
                                    if (cantlines > 1)
                                    {
                                        result = doctransf.Lines.Count;
                                        doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                        doctransf.Lines.Add();
                                        doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                    }
                                    cantlines++;
                                    linestransf++;
                                    doctransf.Lines.ItemCode = artcurrent;
                                    doctransf.Lines.ItemDescription = rsDoc.Fields.Item("Dscription").Value.ToString();
                                    doctransf.Lines.Quantity = totalart;
                                    doctransf.Lines.FromWarehouseCode = "CD_RSV";
                                    doctransf.Lines.WarehouseCode = "CD";
                                }
                                else
                                {
                                    // Procesar los articulos no disponibles a la hora de transferir
                                    lineasnodisp.Add(artcurrent);
                                }
                            }
                            artcurrent = art;
                            totalart = Double.Parse(rsDoc.Fields.Item("Quantity").Value.ToString());
                        }
                        else
                        {
                            totalart += Double.Parse(rsDoc.Fields.Item("Quantity").Value.ToString());
                        }
                        if (i < rsDoc.RecordCount) { rsDoc.MoveNext(); }
                    }
                    // Adicionar ultima fila
                    if (artcurrent != "")
                    {
                        disponible = obtener_exist_articulo(artcurrent, "CD_RSV", out errorMessage);
                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                        if (disponible >= totalart)
                        {
                            if (cantlines > 1)
                            {
                                result = doctransf.Lines.Count;
                                doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                doctransf.Lines.Add();
                                doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                            }
                            linestransf++;
                            doctransf.Lines.ItemCode = artcurrent;
                            doctransf.Lines.ItemDescription = rsDoc.Fields.Item("Dscription").Value.ToString();
                            doctransf.Lines.Quantity = totalart;
                            doctransf.Lines.FromWarehouseCode = "CD_RSV";
                            doctransf.Lines.WarehouseCode = "CD";
                        }
                        else
                        {
                            // Procesar los articulos no disponibles a la hora de transferir
                            lineasnodisp.Add(artcurrent);
                        }
                    }

                    if (linestransf > 0)
                    {
                        result = doctransf.Add();
                        if (result != 0)
                        {
                            if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                            errorMessage = "Revertir Transferencia: " +
                                ((B1.Company.GetLastErrorCode() != 0)
                                ? B1.Company.GetLastErrorDescription()
                                : "");
                            GC.Collect();
                            return errorMessage;
                        }
                    }
                    else
                    {
                        string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                        if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        errorMessage = "Revertir Transferencia: No existen artículos disponibles. " + infonodisp;
                        GC.Collect();
                        return errorMessage;
                    }
                    GC.Collect();
                }

                if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); }
                string newkey = B1.Company.GetNewObjectKey();
                if (newkey != "")
                {
                    //Actualizar datos de Transferencia en Solicitud
                    string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                    string slog = "Cancelada Automáticamente: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + " DocNum:" + obtener_DocNum(newkey, out errorMessage) + infonodisp;
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                    string scom = "Solicitud Cancelada por vencer su período de revisión: " + DateTime.Now.Date.ToString("dd/MM/yyyy");
                    Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string sestado = "C";
                    string SQLQuery = String.Format("UPDATE {1} SET {2} = '{8}', {5} = '{4}', {6} = '{7}', {9} = '{10}' FROM {1} WHERE {0} = '{3}' ",
                                                Constantes.View.CAB_RVT.U_numDoc,   //0
                                                Constantes.View.CAB_RVT.CAB_RV,    //1
                                                Constantes.View.CAB_RVT.U_logs,    //2
                                                sCode,                             //3
                                                scom,                              //4
                                                Constantes.View.CAB_RVT.U_comment,    //5
                                                Constantes.View.CAB_RVT.U_estado,    //6
                                                sestado,   //7
                                                slog,//8
                                                Constantes.View.CAB_RVT.U_idTV,  //9
                                                newkey); ////10
                    oRecordSet.DoQuery(SQLQuery);
                    SQLQuery = String.Format("UPDATE {1} SET {3} = '{4}' FROM {1} WHERE {0} = '{2}'  ",
                                                Constantes.View.DET_RVT.U_numOC,   //0
                                                Constantes.View.DET_RVT.DET_RV,    //1
                                                sCode,                             //2
                                                Constantes.View.DET_RVT.U_idTV,   //3
                                                newkey);                           //4
                    oRecordSet.DoQuery(SQLQuery);
                    cancelar_filas_nodisp(newkey, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                    B1.Application.SetStatusBarMessage("Solicitud Cancelada Automática Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                else
                {
                    if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                    //Actualizar logs en Solicitud
                    string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                    string slog = "Error:No pudo ser Cancelada Automáticamente por no tener disponibilidad: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + infonodisp;
                    string scom = "Solicitud sin disponibilidad al intentar Cancelada por vencer su período de revisión: " + DateTime.Now.Date.ToString("dd/MM/yyyy");
                    Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string SQLQuery = String.Format("UPDATE {1} SET {2} = '{6}', {5} = '{4}'  FROM {1} WHERE {0} = '{3}' ",
                                             Constantes.View.CAB_RVT.U_numDoc,   //0
                                             Constantes.View.CAB_RVT.CAB_RV,    //1
                                             Constantes.View.CAB_RVT.U_logs,    //2
                                             sCode,                             //3
                                             scom,                              //4
                                             Constantes.View.CAB_RVT.U_comment,    //5
                                             slog);//6
                    oRecordSet.DoQuery(SQLQuery);
                    errorMessage = "Transferir Solicitud Cancelada Automática: " + sCode  
                    + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : "");
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Transferir Solicitud Cancelada Automática:  " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return errorMessage;
        }

        private bool cancelar_filas_nodisp(string newkey, out string errorMessage)
        {
            errorMessage = "";
            bool todoOk = true;
            string SQLQuery = String.Empty;
            try
            {
                string filasnodisp = string.Join("-", lineasnodisp);
                if (lineasnodisp != null && lineasnodisp.Count > 0)
                {
                    for (int i = 0; i < lineasnodisp.Count; i++)
                    {
                        string sestado = "N";
                        string nuevaidtv = "";
                        SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3} = '{7}' FROM {1} WHERE {0} = '{4}' AND {3} = '{6}' ",
                                        Constantes.View.DET_RVT.U_codArt,      //0
                                        Constantes.View.DET_RVT.DET_RV,    //1 
                                        Constantes.View.DET_RVT.U_estado,  //2
                                        Constantes.View.DET_RVT.U_idTV,    //3
                                        lineasnodisp[i],                   //4
                                        sestado,                           //5
                                        newkey,                          //6
                                        nuevaidtv);                    //7
                        Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsCards.DoQuery(SQLQuery);
                    }
                    int respuesta = B1.Application.MessageBox("Los artículos " + filasnodisp + " no están disponibles, por tanto, se cancela su transferencia", 1, "OK", " Cancelar");
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Sincronizando artículos no disponibles: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            finally
            {
                lineasnodisp.Clear();
                System.GC.Collect();
            }
            return todoOk;
        }

        private string cancelar_vencidaspormas10D()
        {
            string errorMessage = "";
            try
            {
                //Actualizar estado y comentario
                B1.Application.SetStatusBarMessage("Realizando Cancelación Automática por Fecha de Vencimiento", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                B1.Application.Forms.ActiveForm.Freeze(true);
                string nodoc = "";
                string dentry = "";
                string sCli = "";
                string snCli = "";
                string SQLQuery = String.Format("SELECT {0}, {4}, {5}, {6} FROM {1} WHERE {2} = 'R' AND DAYS_BETWEEN(CURRENT_DATE,{3}) < 0 ",
                                             Constantes.View.CAB_RVT.U_numDoc,
                                             Constantes.View.CAB_RVT.CAB_RV,
                                             Constantes.View.CAB_RVT.U_estado,
                                             Constantes.View.CAB_RVT.U_fechaV,
                                             Constantes.View.CAB_RVT.U_idTR,
                                             Constantes.View.CAB_RVT.U_codCli,
                                             Constantes.View.CAB_RVT.U_cliente
                                             );
                Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(SQLQuery);
                oRecordSet.MoveFirst();   
                if (oRecordSet.RecordCount == 0)
                {
                    B1.Application.SetStatusBarMessage("!No Existen Solicitudes Vencidas! ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                else
                {
                    for (int i = 0; !oRecordSet.EoF; i++)
                    {
                        nodoc = oRecordSet.Fields.Item("U_numDoc").Value.ToString();
                        dentry = oRecordSet.Fields.Item("U_idTR").Value.ToString();
                        sCli = oRecordSet.Fields.Item("U_codCli").Value.ToString();
                        snCli = oRecordSet.Fields.Item("U_cliente").Value.ToString();
                        revertir(nodoc, dentry, sCli, snCli);
                        oRecordSet.MoveNext();
                    }
                    B1.Application.SetStatusBarMessage("Cancelación Automática realizada con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                
            }
            catch (Exception ex)
            {
                errorMessage = "Cancelación automática: " + 
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            finally
            {
                B1.Application.Forms.ActiveForm.Freeze(false);
            }
            return errorMessage;

        }

        private string obtener_Estado(string abrev)
        {
            string resultado = "";
            switch (abrev)
            {
                case "R":
                    {
                        resultado = "Reservada";
                    }
                    break;

                case "A":
                    {
                        resultado = "Aprobada";
                    }
                    break;

                case "T":
                    {
                        resultado = "Transferida";
                    }
                    break;

                case "C":
                    {
                        resultado = "Cancelada";
                    }
                    break;

                case "D":
                    {
                        resultado = "Devuelta";
                    }
                    break;

            }
            return resultado;
        }


        private string GetConfigSociety()
        {
            string errorMessage = "";
            try
            {
                // Determinar Config Sociedad
                String strSQL = String.Format("SELECT {1}, {2}, {3}, {4} FROM {0} ",
                              Constantes.View.oadm.OADM,
                              Constantes.View.oadm.MainCurncy,
                              Constantes.View.oadm.RevOffice,
                              Constantes.View.oadm.CompnyName,
                              Constantes.View.oadm.ExcelPath
                              );
                Recordset rsResult = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsResult.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsResult.Fields;
                rsResult.MoveFirst();
                ExcelPath = !rsResult.EoF ? rsResult.Fields.Item("ExcelPath").Value.ToString() : "";
                CompanyName = !rsResult.EoF ? rsResult.Fields.Item("CompnyName").Value.ToString() : "";
                validPathExcel = Directory.Exists(ExcelPath);
            }
            catch (Exception ex)
            {
                errorMessage = "Configurar Sociedad: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return errorMessage;
        }

        public static bool CanWritePath(string path)
        {
            var writeAllow = false;
            var writeDeny = false;
            var accessControlList = Directory.GetAccessControl(path);
            if (accessControlList == null)
                return false;
            var accessRules = accessControlList.GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
            if (accessRules == null)
                return false;
            foreach (System.Security.AccessControl.FileSystemAccessRule rule in accessRules)
            {
                if ((System.Security.AccessControl.FileSystemRights.Write & rule.FileSystemRights) != System.Security.AccessControl.FileSystemRights.Write) continue;

                if (rule.AccessControlType == System.Security.AccessControl.AccessControlType.Allow)
                    writeAllow = true;
                else if (rule.AccessControlType == System.Security.AccessControl.AccessControlType.Deny)
                    writeDeny = true;
            }
            return writeAllow && !writeDeny;
        }


        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        private string exportar_inicio()
        {
            string errorMessage = "";
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.Workbooks.Add();
                workBook = (Excel.Workbook)excelApp.ActiveWorkbook;
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Inicio " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage; 
        }

        private string exportar_fin()
        {
            string errorMessage = "";
            try
            {
                try
                {
                    workBook.SaveAs(ExcelPath + "RESERVAS_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ".xlsx");
                }
                catch 
                    //(Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error al guardar el fichero Excel en el directorio configurado...Seleccione otro..", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    string fname = ExcelPath + "RESERVAS_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ".xlsx";
                    FileInfo fileInfo = new FileInfo(fname);
                    object fileName = excelApp.GetSaveAsFilename(fileInfo.Name, string.Format("Excel files (*{0}), *{0}", fileInfo.Extension), 1);
                    if (fileName.ToString() != "False")
                    {
                        try
                        {
                            workBook.SaveAs(fileName);
                        }
                        catch
                        {
                            B1.Application.SetStatusBarMessage("Error: Fichero Excel No Creado..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                    }
                }
                if (workSheet != null)
                {
                    Marshal.ReleaseComObject(workSheet);
                }
                if (workSheet2 != null)
                {
                    Marshal.ReleaseComObject(workSheet2);
                }

                if (workBook != null)
                {
                    workBook.Close();
                    Marshal.ReleaseComObject(workBook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                workSheet = null;
                workSheet2 = null;
                workBook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                B1.Application.SetStatusBarMessage("Exportación EXCEL realizada con éxito...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Fin " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage; 
        }

        public string cargar_items_exportar(out Recordset rs)
        {
            string errorMessage = "";
            string filtrado = "";
            rs = null;
            int lines_filter = 0;
            try
            {
                B1.Application.SetStatusBarMessage("Cargando datos de Solicitudes de Reservas de Stock para exportar a Excel...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                SAPbouiCOM.CheckBox cboxNue = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxNue).Specific;
                SAPbouiCOM.CheckBox cboxApr = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxApr).Specific;
                SAPbouiCOM.CheckBox cboxCan = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCan).Specific;
                SAPbouiCOM.CheckBox cboxDev = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxDev).Specific;
                string SQLQuery = string.Empty;
                SAPbouiCOM.CheckBox cboxPer = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxPer).Specific;
                string condPer = String.Empty;
                if (cboxPer.Checked == true)
                {
                    lines_filter++;

                    SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                    SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                    condPer = desde.Value != "" && hasta.Value != "" ? Constantes.View.CAB_RVT.U_fechaC + " between '" + desde.Value + "' AND '" + hasta.Value + "'" : "";
                    filtrado = filtrado + (condPer != "" ? ">Período Seleccionado: " +
                        DateTime.ParseExact(desde.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy") +
                        " a " +
                        DateTime.ParseExact(hasta.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy") : "");
                }

                SAPbouiCOM.CheckBox cboxCli = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCli).Specific;
                string condCli = String.Empty;
                if (cboxCli.Checked == true)
                {
                    SAPbouiCOM.ComboBox cli = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbCli).Specific;
                    string selCli = (cli.Value != "") ? cli.Value : "";
                    condCli = selCli != "" ? Constantes.View.CAB_RVT.U_codCli + " = '" + selCli + "'" : condCli;
                    lines_filter++;
                    filtrado = filtrado + (condCli != "" ? (lines_filter > 1 ? "\n" : "") + ">Cliente Seleccionado: " + selCli : "");
                }

                SAPbouiCOM.CheckBox cboxArt = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxArt).Specific;
                string condArt = String.Empty;
                if (cboxArt.Checked == true)
                {
                    SAPbouiCOM.ComboBox art = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbArt).Specific;
                    string selArt = (art.Value != "") ? art.Value : "";
                    condArt = selArt != "" ? Constantes.View.DET_RVT.U_codArt + " = '" + selArt + "'" : condArt;
                    lines_filter++;
                    filtrado = filtrado + (condArt != "" ? (lines_filter > 1 ? "\n" : "") + ">Artículo Seleccionado: " + selArt : "");
                }

                SAPbouiCOM.CheckBox cboxVend = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxVend).Specific;
                string condVend = String.Empty;
                if (cboxVend.Checked == true)
                {
                    SAPbouiCOM.ComboBox vend = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbVend).Specific;
                    string selVend = (vend.Value != "") ? vend.Value : "";
                    condVend = selVend != "" ? Constantes.View.CAB_RVT.U_idVend + " = '" + selVend + "'" : condVend;
                    lines_filter++;
                    filtrado = filtrado + (condVend != "" ? (lines_filter > 1 ? "\n" : "") + ">Vendedor Seleccionado: " + selVend : "");

                }

                string condNue = String.Empty;
                condNue = (cboxNue.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'R' " : condNue;

                string condApr = String.Empty;
                condApr = (cboxApr.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'A' " : condApr;

                string condCan = String.Empty;
                condCan = (cboxCan.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'C' " : condCan;

                string condDev = String.Empty;
                condDev = (cboxDev.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'D' " : condDev;

                lines_filter++;
                filtrado = filtrado + (lines_filter > 1 ? "\n" : "") + ">Estados Seleccionados: " +
                        (condNue == String.Empty ? "Reservada" : "") + (condApr == String.Empty ? "-Aprobada" : "") +
                        (condCan == String.Empty ? "-Cancelada " : "") + (condDev == String.Empty ? "-Devuelta" : "");


                workSheet2.Cells[3, 2] = filtrado;
                workSheet2.Rows[3].RowHeight = lines_filter * 15;


                string cadw = "";
                cadw = condPer != String.Empty || condCli != String.Empty || condArt != String.Empty || condVend != String.Empty ||
                       condNue != String.Empty || condApr != String.Empty || condCan != String.Empty || condDev != String.Empty
                       ? " WHERE " : "";
                cadw = cadw + (condPer != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condPer : "");
                cadw = cadw + (condCli != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCli : "");
                cadw = cadw + (condArt != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condArt : "");
                cadw = cadw + (condVend != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condVend : "");
                cadw = cadw + (condNue != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condNue : "");
                cadw = cadw + (condApr != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condApr : "");
                //cadw = cadw + (condTra  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condTra : "");
                cadw = cadw + (condCan != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCan : "");
                cadw = cadw + (condDev != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condDev : "");

                string adicjoin =" INNER JOIN " +
                Constantes.View.DET_RVT.DET_RV + " T3 ON T0." + Constantes.View.CAB_RVT.U_numDoc +
                " = T3." + Constantes.View.DET_RVT.U_numOC ;

                string adicgroup = (condCli != String.Empty || condArt != String.Empty) ? " GROUP BY " +
                 " T0." + Constantes.View.CAB_RVT.U_numDoc + ", T0." + Constantes.View.CAB_RVT.U_idVend + " , T1." +
                Constantes.View.ousr.uName + ", T0." +
                Constantes.View.CAB_RVT.U_codCli + ", T0." +
                Constantes.View.CAB_RVT.U_cliente + ", T0." +
                Constantes.View.CAB_RVT.U_fechaC + ", T0." +
                Constantes.View.CAB_RVT.U_fechaV + ", DAYS_BETWEEN( CURRENT_DATE,T0." +
                Constantes.View.CAB_RVT.U_fechaV + " ), T0." + Constantes.View.CAB_RVT.U_idAut + ", T2." +
                Constantes.View.ousr.uName + ", T0." + Constantes.View.CAB_RVT.U_estado + ", T0." +
                Constantes.View.CAB_RVT.U_idTR + " , T0." + Constantes.View.CAB_RVT.U_idTV +
                ", T0." + Constantes.View.CAB_RVT.U_amount
                    //+",  T0." +   Constantes.View.CAB_RVT.U_comment
                : "";

                SQLQuery = String.Format("SELECT T0.{1} , T0.{4}, T1.{3} U_vend, T0.{6}, T0.{7}, DAYS_BETWEEN( CURRENT_DATE,T0.{7}) U_diasv, " +
                      " T0.{8}, T2.{3} U_aut, T0.{9}, T0.{10}, T0.{11}, T0.{16}, T0.{17}, T0.{26}, CAST(T0.{1} AS INT) AS ND," +
                       " T3.{14}, T3.{15}, T3.{18}, T3.{27}, T3.{26} U_amountitem" +   
                    //, T0.{22} " +
                      " FROM {0} T0 INNER JOIN {2} T1 ON T0.{4} = T1.{5} " +
                      " LEFT JOIN {2} T2 ON T0.{8} = T2.{5} {24}  {23}   ORDER BY CAST(T0.{1} AS INT) ",
                                              Constantes.View.CAB_RVT.CAB_RV, //0
                                              Constantes.View.CAB_RVT.U_numDoc,//1
                                              Constantes.View.ousr.OUSR, //2
                                              Constantes.View.ousr.uName, //3
                                              Constantes.View.CAB_RVT.U_idVend,//4
                                              Constantes.View.ousr.uCode, //5
                                              Constantes.View.CAB_RVT.U_fechaC, //6
                                              Constantes.View.CAB_RVT.U_fechaV, //7
                                              Constantes.View.CAB_RVT.U_idAut, //8
                                              Constantes.View.CAB_RVT.U_estado, //9
                                              Constantes.View.CAB_RVT.U_idTR, //10
                                              Constantes.View.CAB_RVT.U_idTV, //11
                                              Constantes.View.DET_RVT.DET_RV, //12
                                              Constantes.View.DET_RVT.U_numOC, //13
                                              Constantes.View.DET_RVT.U_codArt, //14
                                              Constantes.View.DET_RVT.U_articulo, //15
                                              Constantes.View.CAB_RVT.U_codCli, //16
                                              Constantes.View.CAB_RVT.U_cliente, //17
                                              Constantes.View.DET_RVT.U_cant, //18
                                              Constantes.View.DET_RVT.U_onHand, //19
                                              Constantes.View.DET_RVT.U_estado, //20
                                              Constantes.View.DET_RVT.U_idTV, //21                                                
                                              Constantes.View.CAB_RVT.U_comment,//22,
                                              cadw, //23,
                                              adicjoin, //24
                                              adicgroup, //25
                                              Constantes.View.CAB_RVT.U_amount, //26
                                              Constantes.View.DET_RVT.U_price //27
                                              );

                rs = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(SQLQuery);

            }
            catch (Exception ex)
            {
                errorMessage = "Exportar Excel: Cargar Solicitudes: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return errorMessage;
        }


        public string cargar_datos_exportar(out Recordset rs)
        {
            string errorMessage = "";
            string filtrado = "";
            rs = null;
            int lines_filter = 0;
            try
            {
                B1.Application.SetStatusBarMessage("Cargando datos de Solicitudes de Reservas de Stock para exportar a Excel...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                SAPbouiCOM.CheckBox cboxNue = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxNue).Specific;
                SAPbouiCOM.CheckBox cboxApr = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxApr).Specific;
                SAPbouiCOM.CheckBox cboxCan = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCan).Specific;
                SAPbouiCOM.CheckBox cboxDev = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxDev).Specific;
                string SQLQuery = string.Empty;
                SAPbouiCOM.CheckBox cboxPer = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxPer).Specific;
                string condPer = String.Empty;
                if (cboxPer.Checked == true)
                {
                    lines_filter++;

                    SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                    SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                    condPer = desde.Value != "" && hasta.Value != "" ? Constantes.View.CAB_RVT.U_fechaC + " between '" + desde.Value + "' AND '" + hasta.Value + "'" : "";
                    filtrado = filtrado + (condPer != "" ? ">Período Seleccionado: " +
                        DateTime.ParseExact(desde.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy") +
                        " a " +
                        DateTime.ParseExact(hasta.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy") : "");
                }

                SAPbouiCOM.CheckBox cboxCli = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCli).Specific;
                string condCli = String.Empty;
                if (cboxCli.Checked == true)
                {
                    SAPbouiCOM.ComboBox cli = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbCli).Specific;
                    string selCli = (cli.Value != "") ? cli.Value : "";
                    condCli = selCli != "" ? Constantes.View.CAB_RVT.U_codCli + " = '" + selCli + "'" : condCli;
                    lines_filter++;
                    filtrado = filtrado + (condCli != "" ? (lines_filter > 1 ? "\n" : "") + ">Cliente Seleccionado: " + selCli : "");
                }

                SAPbouiCOM.CheckBox cboxArt = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxArt).Specific;
                string condArt = String.Empty;
                if (cboxArt.Checked == true)
                {
                    SAPbouiCOM.ComboBox art = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbArt).Specific;
                    string selArt = (art.Value != "") ? art.Value : "";
                    condArt = selArt != "" ? Constantes.View.DET_RVT.U_codArt + " = '" + selArt + "'" : condArt;
                    lines_filter++;
                    filtrado = filtrado + (condArt != "" ? (lines_filter > 1 ? "\n" : "") + ">Artículo Seleccionado: " + selArt : "");
                }

                SAPbouiCOM.CheckBox cboxVend = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxVend).Specific;
                string condVend = String.Empty;
                if (cboxVend.Checked == true)
                {
                    SAPbouiCOM.ComboBox vend = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbVend).Specific;
                    string selVend = (vend.Value != "") ? vend.Value : "";
                    condVend = selVend != "" ? Constantes.View.CAB_RVT.U_idVend + " = '" + selVend + "'" : condVend;
                    lines_filter++;
                    filtrado = filtrado + (condVend != "" ? (lines_filter > 1 ? "\n" : "") + ">Vendedor Seleccionado: " + selVend : "");

                }

                string condNue = String.Empty;
                condNue = (cboxNue.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'R' " : condNue;

                string condApr = String.Empty;
                condApr = (cboxApr.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'A' " : condApr;

                string condCan = String.Empty;
                condCan = (cboxCan.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'C' " : condCan;

                string condDev = String.Empty;
                condDev = (cboxDev.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'D' " : condDev;

                lines_filter++;
                filtrado = filtrado + (lines_filter > 1 ? "\n" : "") + ">Estados Seleccionados: " +
                        (condNue == String.Empty ? "Reservada" : "") + (condApr == String.Empty ? "-Aprobada" : "") +
                        (condCan == String.Empty ? "-Cancelada " : "") + (condDev == String.Empty ? "-Devuelta" : "");
                

                workSheet.Cells[3, 2] = filtrado;
                workSheet.Rows[3].RowHeight = lines_filter * 15 ;


                string cadw = "";
                cadw = condPer != String.Empty || condCli != String.Empty || condArt != String.Empty || condVend != String.Empty ||
                       condNue != String.Empty || condApr != String.Empty || condCan != String.Empty || condDev != String.Empty
                       ? " WHERE " : "";
                cadw = cadw + (condPer != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condPer : "");
                cadw = cadw + (condCli != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCli : "");
                cadw = cadw + (condArt != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condArt : "");
                cadw = cadw + (condVend != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condVend : "");
                cadw = cadw + (condNue != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condNue : "");
                cadw = cadw + (condApr != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condApr : "");
                //cadw = cadw + (condTra  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condTra : "");
                cadw = cadw + (condCan != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCan : "");
                cadw = cadw + (condDev != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condDev : "");

                string adicjoin = (condArt != String.Empty) ? " INNER JOIN " +
                Constantes.View.DET_RVT.DET_RV + " T3 ON T0." + Constantes.View.CAB_RVT.U_numDoc +
                " = T3." + Constantes.View.DET_RVT.U_numOC : "";

                string adicgroup = (condCli != String.Empty || condArt != String.Empty) ? " GROUP BY " +
                 " T0." + Constantes.View.CAB_RVT.U_numDoc + ", T0." + Constantes.View.CAB_RVT.U_idVend + " , T1." +
                Constantes.View.ousr.uName + ", T0." +
                Constantes.View.CAB_RVT.U_codCli + ", T0." +
                Constantes.View.CAB_RVT.U_cliente + ", T0." +
                Constantes.View.CAB_RVT.U_fechaC + ", T0." +
                Constantes.View.CAB_RVT.U_fechaV + ", DAYS_BETWEEN( CURRENT_DATE,T0." +
                Constantes.View.CAB_RVT.U_fechaV + " ), T0." + Constantes.View.CAB_RVT.U_idAut + ", T2." +
                Constantes.View.ousr.uName + ", T0." + Constantes.View.CAB_RVT.U_estado + ", T0." +
                Constantes.View.CAB_RVT.U_idTR + " , T0." + Constantes.View.CAB_RVT.U_idTV +
                ", T0." + Constantes.View.CAB_RVT.U_amount
                    //+",  T0." +   Constantes.View.CAB_RVT.U_comment
                : "";

                SQLQuery = String.Format("SELECT T0.{1} , T0.{4}, T1.{3} U_vend, T0.{6}, T0.{7}, DAYS_BETWEEN( CURRENT_DATE,T0.{7}) U_diasv, " +
                      " T0.{8}, T2.{3} U_aut, T0.{9}, T0.{10}, T0.{11}, T0.{16}, T0.{17}, T0.{26}, CAST(T0.{1} AS INT) AS ND" +
                    //, T0.{22} " +
                      " FROM {0} T0 INNER JOIN {2} T1 ON T0.{4} = T1.{5} " +
                      " LEFT JOIN {2} T2 ON T0.{8} = T2.{5} {24}  {23}  {25}  ORDER BY CAST(T0.{1} AS INT) ",
                                              Constantes.View.CAB_RVT.CAB_RV, //0
                                              Constantes.View.CAB_RVT.U_numDoc,//1
                                              Constantes.View.ousr.OUSR, //2
                                              Constantes.View.ousr.uName, //3
                                              Constantes.View.CAB_RVT.U_idVend,//4
                                              Constantes.View.ousr.uCode, //5
                                              Constantes.View.CAB_RVT.U_fechaC, //6
                                              Constantes.View.CAB_RVT.U_fechaV, //7
                                              Constantes.View.CAB_RVT.U_idAut, //8
                                              Constantes.View.CAB_RVT.U_estado, //9
                                              Constantes.View.CAB_RVT.U_idTR, //10
                                              Constantes.View.CAB_RVT.U_idTV, //11
                                              Constantes.View.DET_RVT.DET_RV, //12
                                              Constantes.View.DET_RVT.U_numOC, //13
                                              Constantes.View.DET_RVT.U_codArt, //14
                                              Constantes.View.DET_RVT.U_articulo, //15
                                              Constantes.View.CAB_RVT.U_codCli, //16
                                              Constantes.View.CAB_RVT.U_cliente, //17
                                              Constantes.View.DET_RVT.U_cant, //18
                                              Constantes.View.DET_RVT.U_onHand, //19
                                              Constantes.View.DET_RVT.U_estado, //20
                                              Constantes.View.DET_RVT.U_idTV, //21                                                
                                              Constantes.View.CAB_RVT.U_comment,//22,
                                              cadw, //23,
                                              adicjoin, //24
                                              adicgroup, //25
                                              Constantes.View.CAB_RVT.U_amount //26

                                              );

                rs = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(SQLQuery);
                
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar Excel: Cargar Solicitudes: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
            }
            return errorMessage;
        }

        private string exportar_Resumen_cabecera()
        {
            string errorMessage = "";
            B1.Application.SetStatusBarMessage("Exportando Datos de Solicitudes de Reserva...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            workSheet = (Excel.Worksheet)excelApp.Worksheets.Add();
            ((Excel.Worksheet)excelApp.ActiveWorkbook.Sheets[1]).Select();
            workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "Solicitudes";

            try
            {
                workSheet.Cells[1, 2] = "SOLICITUDES DE RESERVA DE INVENTARIOS";
                workSheet.Cells[1, 2].Font.Bold = true;
                workSheet.Cells[1, 2].Font.Size = 16;
                workSheet.Cells[1, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[2, 2] = CompanyName;
                workSheet.Cells[2, 2].Font.Bold = true;
                workSheet.Cells[2, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                DateTimeFormatInfo dtinfo = new CultureInfo("es-ES", false).DateTimeFormat;
                TextInfo myTI = new CultureInfo("es-ES", false).TextInfo;
  
                //workSheet.Cells[3, 2] = filtrado;
                workSheet.Cells[3, 2].Font.Bold = true;
                workSheet.Cells[3, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[5, "A"] = "NO.";
                workSheet.Cells[5, "B"] = "DOC.";
                workSheet.Cells[5, "C"] = "COD.VEND.";
                workSheet.Cells[5, "D"] = "VENDEDOR";
                workSheet.Cells[5, "E"] = "FECHA";
                workSheet.Cells[5, "F"] = "VENCIMIENTO";
                workSheet.Cells[5, "G"] = "DÍAS";
                workSheet.Cells[5, "H"] = "COD.AUTOR.";
                workSheet.Cells[5, "I"] = "AUTORIZADOR";
                workSheet.Cells[5, "J"] = "COD.CLI";
                workSheet.Cells[5, "K"] = "CLIENTE";
                workSheet.Cells[5, "L"] = "ESTADO";
                workSheet.Cells[5, "M"] = "COSTO";
                workSheet.Cells[5, "N"] = "TR. RESERVA";
                workSheet.Cells[5, "O"] = "TR. DEVOLUC.";
                workSheet.Cells[5, "P"] = "OBSERVACIONES";

                workSheet.Range["A5", "P5"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatList2);
                row = 5;
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Hoja Datos Cabecera " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage; 
        }

        private string exportar_Resumen_cerrar()
        {
            string errorMessage = "";
            try
            {
                if (row>6)
                {
                    row++;
                    workSheet.Cells[row, "B"] = "TOTAL";
                    workSheet.Cells[row, "M"] = "=SUM(M6:M" + (row - 1).ToString() + ")";

                    workSheet.Range["A" + row.ToString(), "P" + row.ToString()].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatAccounting3);
                    workSheet.Range["G" + row.ToString(), "G" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0";
                    workSheet.Range["M" + row.ToString(), "M" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0.00";
                }


                int ultcol = 16;
                var rowRngRptTitle = workSheet.Range[workSheet.Cells[1, 2], workSheet.Cells[1, 15]];
                rowRngRptTitle.Merge(Type.Missing);
                rowRngRptTitle.Font.Bold = true;
                rowRngRptTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                var rowRngRptTitle2 = workSheet.Range[workSheet.Cells[2, 2], workSheet.Cells[2, 15]];
                rowRngRptTitle2.Merge(Type.Missing);
                rowRngRptTitle2.Font.Bold = true;
                rowRngRptTitle2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                var rowRngRptTitle3 = workSheet.Range[workSheet.Cells[3, 2], workSheet.Cells[3, 15]];
                rowRngRptTitle3.Merge(Type.Missing);
                rowRngRptTitle3.Font.Bold = true;
                rowRngRptTitle3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                for (int i = 1; i < ultcol; i++)
                {
                    workSheet.Columns[i].AutoFit();
                }

                // freeze panel superior
                workSheet.Application.ActiveWindow.SplitRow = 5;
                workSheet.Application.ActiveWindow.FreezePanes = true;
                workSheet.PageSetup.Orientation =
                            Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                workSheet.PageSetup.Zoom = false;
                workSheet.PageSetup.FitToPagesTall = row > 25 ? (row / 25) : 1;
                workSheet.PageSetup.FitToPagesWide = 1;
                workSheet.PageSetup.PrintTitleRows = "$A1:$" + "P5";
                workSheet.PageSetup.RightHeader = "Página &P de &N";
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Hoja Datos Totales " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage; 
        }

        private string exportar_Resumen_detalle()
        {
            string errorMessage = "";
            try
            {
                Recordset rs = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                errorMessage = cargar_datos_exportar(out rs);
                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage), true); }
                rs.MoveFirst();
                while (!rs.EoF)
                {
                    row++;
                    workSheet.Cells[row, "A"] = (row - 5).ToString();
                    workSheet.Cells[row, "B"] = rs.Fields.Item("U_numDoc").Value;
                    workSheet.Cells[row, "C"] = rs.Fields.Item("U_IdVend").Value;
                    workSheet.Cells[row, "D"] = rs.Fields.Item("U_vend").Value;
                    workSheet.Cells[row, "E"] = rs.Fields.Item("U_fechaC").Value;
                    workSheet.Cells[row, "F"] = rs.Fields.Item("U_fechaV").Value;
                    workSheet.Cells[row, "G"] = rs.Fields.Item("U_diasv").Value;
                    workSheet.Cells[row, "H"] = rs.Fields.Item("U_IdAut").Value;
                    workSheet.Cells[row, "I"] = rs.Fields.Item("U_aut").Value;
                    workSheet.Cells[row, "J"] = rs.Fields.Item("U_codCli").Value;
                    workSheet.Cells[row, "K"] = rs.Fields.Item("U_cliente").Value;
                    workSheet.Cells[row, "L"] = obtener_Estado(rs.Fields.Item("U_estado").Value);
                    workSheet.Cells[row, "M"] = rs.Fields.Item("U_amount").Value.ToString().Replace(".", "");
                    workSheet.Cells[row, "N"] =  obtener_DocNum(rs.Fields.Item("U_IdTR").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    workSheet.Cells[row, "O"] = obtener_DocNum(rs.Fields.Item("U_idTV").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    workSheet.Cells[row, "P"] = obtener_Comentario(rs.Fields.Item("U_numDoc").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }

                    workSheet.Range["G" + row.ToString(), "G" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0";
                    workSheet.Range["M" + row.ToString(), "M" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0.00";

                    rs.MoveNext();
                    if (!rs.EoF)
                    { B1.Application.SetStatusBarMessage("ESPERE......Exportando Solicitud: " + 
                        rs.Fields.Item("U_numDoc").Value.ToString() + "    (" + (row - 5).ToString() + "/" + 
                        rs.RecordCount.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, false); }
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Hoja Datos Detalle " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage; 
        }


        private string exportar_Items_cabecera()
        {
            string errorMessage = "";
            B1.Application.SetStatusBarMessage("Exportando Datos de Items de Reserva...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            workSheet2 = (Excel.Worksheet)excelApp.Worksheets.Add();
            ((Excel.Worksheet)excelApp.ActiveWorkbook.Sheets[1]).Select();
            workSheet2 = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet2.Name = "Items";

            try
            {
                workSheet2.Cells[1, 2] = "SOLICITUDES DE RESERVA DE INVENTARIOS";
                workSheet2.Cells[1, 2].Font.Bold = true;
                workSheet2.Cells[1, 2].Font.Size = 16;
                workSheet2.Cells[1, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet2.Cells[2, 2] = CompanyName;
                workSheet2.Cells[2, 2].Font.Bold = true;
                workSheet2.Cells[2, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                DateTimeFormatInfo dtinfo = new CultureInfo("es-ES", false).DateTimeFormat;
                TextInfo myTI = new CultureInfo("es-ES", false).TextInfo;

                //workSheet2.Cells[3, 2] = filtrado;
                workSheet2.Cells[3, 2].Font.Bold = true;
                workSheet2.Cells[3, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet2.Cells[5, "A"] = "NO.";
                workSheet2.Cells[5, "B"] = "DOC.";
                workSheet2.Cells[5, "C"] = "COD.VEND.";
                workSheet2.Cells[5, "D"] = "VENDEDOR";
                workSheet2.Cells[5, "E"] = "FECHA";
                workSheet2.Cells[5, "F"] = "VENCIMIENTO";
                workSheet2.Cells[5, "G"] = "DÍAS";
                workSheet2.Cells[5, "H"] = "COD.AUTOR.";
                workSheet2.Cells[5, "I"] = "AUTORIZADOR";
                workSheet2.Cells[5, "J"] = "COD.CLI";
                workSheet2.Cells[5, "K"] = "CLIENTE";
                workSheet2.Cells[5, "L"] = "ESTADO";
                workSheet2.Cells[5, "M"] = "COSTO";
                workSheet2.Cells[5, "N"] = "TR. RESERVA";
                workSheet2.Cells[5, "O"] = "TR. DEVOLUC.";
                workSheet2.Cells[5, "P"] = "OBSERVACIONES";

                workSheet2.Cells[5, "Q"] = "COD.ITEM";
                workSheet2.Cells[5, "R"] = "DESCRIPCIÓN";
                workSheet2.Cells[5, "S"] = "CANT.";
                workSheet2.Cells[5, "T"] = "COSTO";
                workSheet2.Cells[5, "U"] = "COSTO TOTAL";

                workSheet2.Range["A5", "U5"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatList2);
                row = 5;
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Hoja Datos Cabecera Items " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string exportar_Items_cerrar()
        {
            string errorMessage = "";
            try
            {
                if (row > 6)
                {
                    row++;
                    workSheet2.Cells[row, "B"] = "TOTAL";
                    workSheet2.Cells[row, "M"] = "=SUM(M6:M" + (row - 1).ToString() + ")";
                    workSheet2.Cells[row, "U"] = "=SUM(U6:U" + (row - 1).ToString() + ")";

                    workSheet2.Range["A" + row.ToString(), "U" + row.ToString()].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatAccounting3);
                    workSheet2.Range["G" + row.ToString(), "G" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0";
                    workSheet2.Range["M" + row.ToString(), "M" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0.00";
                    workSheet2.Range["S" + row.ToString(), "S" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0";
                    workSheet2.Range["T" + row.ToString(), "U" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0.00";
                }


                int ultcol = 22;
                var rowRngRptTitle = workSheet2.Range[workSheet2.Cells[1, 2], workSheet2.Cells[1, 20]];
                rowRngRptTitle.Merge(Type.Missing);
                rowRngRptTitle.Font.Bold = true;
                rowRngRptTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                var rowRngRptTitle2 = workSheet2.Range[workSheet2.Cells[2, 2], workSheet2.Cells[2, 20]];
                rowRngRptTitle2.Merge(Type.Missing);
                rowRngRptTitle2.Font.Bold = true;
                rowRngRptTitle2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                var rowRngRptTitle3 = workSheet2.Range[workSheet2.Cells[3, 2], workSheet2.Cells[3, 20]];
                rowRngRptTitle3.Merge(Type.Missing);
                rowRngRptTitle3.Font.Bold = true;
                rowRngRptTitle3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                for (int i = 1; i < ultcol; i++)
                {
                    workSheet2.Columns[i].AutoFit();
                }

                // freeze panel superior
                workSheet2.Application.ActiveWindow.SplitRow = 5;
                workSheet2.Application.ActiveWindow.FreezePanes = true;
                workSheet2.PageSetup.Orientation =
                            Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                workSheet2.PageSetup.Zoom = false;
                workSheet2.PageSetup.FitToPagesTall = row > 25 ? (row / 25) : 1;
                workSheet2.PageSetup.FitToPagesWide = 1;
                workSheet2.PageSetup.PrintTitleRows = "$A1:$" + "U5";
                workSheet2.PageSetup.RightHeader = "Página &P de &N";
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Hoja Items Totales " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string exportar_Items_detalle()
        {
            string errorMessage = "";
            try
            {
                Recordset rs = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                errorMessage = cargar_items_exportar(out rs);
                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage), true); }
                rs.MoveFirst();
                while (!rs.EoF)
                {
                    row++;
                    workSheet2.Cells[row, "A"] = (row - 5).ToString();
                    workSheet2.Cells[row, "B"] = rs.Fields.Item("U_numDoc").Value;
                    workSheet2.Cells[row, "C"] = rs.Fields.Item("U_IdVend").Value;
                    workSheet2.Cells[row, "D"] = rs.Fields.Item("U_vend").Value;
                    workSheet2.Cells[row, "E"] = rs.Fields.Item("U_fechaC").Value;
                    workSheet2.Cells[row, "F"] = rs.Fields.Item("U_fechaV").Value;
                    workSheet2.Cells[row, "G"] = rs.Fields.Item("U_diasv").Value;
                    workSheet2.Cells[row, "H"] = rs.Fields.Item("U_IdAut").Value;
                    workSheet2.Cells[row, "I"] = rs.Fields.Item("U_aut").Value;
                    workSheet2.Cells[row, "J"] = rs.Fields.Item("U_codCli").Value;
                    workSheet2.Cells[row, "K"] = rs.Fields.Item("U_cliente").Value;
                    workSheet2.Cells[row, "L"] = obtener_Estado(rs.Fields.Item("U_estado").Value);
                    workSheet2.Cells[row, "M"] = rs.Fields.Item("U_amount").Value.ToString().Replace(".", "");
                    workSheet2.Cells[row, "N"] = obtener_DocNum(rs.Fields.Item("U_IdTR").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    workSheet2.Cells[row, "O"] = obtener_DocNum(rs.Fields.Item("U_idTV").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    workSheet2.Cells[row, "P"] = obtener_Comentario(rs.Fields.Item("U_numDoc").Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    workSheet2.Cells[row, "Q"] = rs.Fields.Item("U_codArt").Value;
                    workSheet2.Cells[row, "R"] = rs.Fields.Item("U_articulo").Value;
                    workSheet2.Cells[row, "S"] = rs.Fields.Item("U_cant").Value;
                    workSheet2.Cells[row, "T"] = rs.Fields.Item("U_price").Value.ToString().Replace(".", "");
                    workSheet2.Cells[row, "U"] = rs.Fields.Item("U_amountitem").Value.ToString().Replace(".", "");

                    workSheet2.Range["G" + row.ToString(), "G" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0";
                    workSheet2.Range["S" + row.ToString(), "S" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0";
                    workSheet2.Range["M" + row.ToString(), "M" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0.00";
                    workSheet2.Range["T" + row.ToString(), "U" + row.ToString()].NumberFormat = "#,##0;[Red]-#,##0.00";

                    rs.MoveNext();
                    if (!rs.EoF)
                    {
                        B1.Application.SetStatusBarMessage("ESPERE......Exportando Solicitud: " +
                          rs.Fields.Item("U_numDoc").Value.ToString() + "  (" + (row - 5).ToString() + "/" +
                          rs.RecordCount.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Exportar a Excel: Hoja Items Detalle " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }


        
        private string exportar()
        {
            string errorMessage = "";

            try {

                if (!CanWritePath(ExcelPath))
                {
                    errorMessage = "Exportar a Excel: No es posible guardar el fichero Excel en el directorio configurado...Seleccione otro..";
                    return errorMessage;
                }
                B1.Application.Forms.ActiveForm.Freeze(true);
                errorMessage = exportar_inicio();
                if (!string.IsNullOrEmpty(errorMessage)) {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage; 
                }

                errorMessage = exportar_Items_cabecera();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }
                errorMessage = exportar_Items_detalle();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }
                errorMessage = exportar_Items_cerrar();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }


                errorMessage = exportar_Resumen_cabecera();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }
                errorMessage = exportar_Resumen_detalle();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }
                errorMessage = exportar_Resumen_cerrar();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }

                // Borrar las hojas innecesarias
                foreach (Excel.Worksheet displayWorksheet in workBook.Worksheets)
                {
                    if (displayWorksheet.Name.Substring(0,1) == "H")
                    {
                        displayWorksheet.Delete();
                    }
                }

                errorMessage = exportar_fin();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    B1.Application.Forms.ActiveForm.Freeze(false);
                    return errorMessage;
                }

                
                B1.Application.SetStatusBarMessage("Exportación a Excel realizada con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            catch (Exception ex)
            {
                
                errorMessage =  "Exportar a Excel: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                B1.Application.Forms.ActiveForm.Freeze(false);
            }

            return errorMessage;            
         }
    }

}

