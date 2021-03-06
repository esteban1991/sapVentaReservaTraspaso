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

        public PantallaAprobac()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Aprobac.srf"], "AprRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString())
        {
            formActual = "AprRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();

            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);

            cancelar_vencidaspormas10D();
            cargar_datos_iniciales();
            cargar_datos_matriz();
        }

        private void cargar_datos_iniciales()
        {
            SAPbouiCOM.CheckBox cboxNue = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxNue).Specific;
            cboxNue.Checked = true;

            SAPbouiCOM.CheckBox cboxApr = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxApr).Specific;
            cboxApr.Checked = true;

            SAPbouiCOM.CheckBox cboxTra = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxTra).Specific;
            cboxTra.Checked = false;

            SAPbouiCOM.CheckBox cboxCan = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCan).Specific;
            cboxCan.Checked = false;

            SAPbouiCOM.CheckBox cboxDev = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxDev).Specific;
            cboxDev.Checked = false;

            AForm = B1.Application.Forms.ActiveForm;
            AForm.EnableMenu("1290", false); AForm.EnableMenu("1289", false);
            AForm.EnableMenu("1288", false); AForm.EnableMenu("1291", false);
            AForm.EnableMenu("1282", false);   // crear
            AForm.EnableMenu("1281", false);  //buscar
            AForm.EnableMenu("1283", false);  //eliminar
            AForm.EnableMenu("1292", false);  //buscar
            AForm.EnableMenu("1293", false);  //eliminar
        }

        public void cargar_datos_matriz()
        {
            B1.Application.SetStatusBarMessage("Cargando datos de Solicitudes de Reservas de Stock para su Autorización...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            bool todoOk = true;
            string serror = "";
            formActual = B1.Application.Forms.ActiveForm.UniqueID;
            AForm = B1.Application.Forms.ActiveForm;
            AMatrix = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item("mtxaprob").Specific;
  
            string filtrado = "";
    
            try
            {
                B1.Application.Forms.ActiveForm.Freeze(true);

                SAPbouiCOM.CheckBox cboxNue = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxNue).Specific;
                SAPbouiCOM.CheckBox cboxApr = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxApr).Specific;
                SAPbouiCOM.CheckBox cboxTra = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxTra).Specific;
                SAPbouiCOM.CheckBox cboxCan = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCan).Specific;
                SAPbouiCOM.CheckBox cboxDev = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxDev).Specific;


                string SQLQuery = string.Empty;

                SAPbouiCOM.CheckBox cboxPer = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxPer).Specific;
                string condPer = String.Empty;
                if (cboxPer.Checked == true)
                {
                    SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                    SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                    condPer = desde.Value!="" && hasta.Value != "" ?Constantes.View.CAB_RVT.U_fechaC + " between '" + desde.Value + "' AND ' " + hasta.Value + "'":"" ;
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
                    condCli = selCli != "" ? Constantes.View.DET_RVT.U_codCli + " = '" + selCli + "'": condCli;
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

                string condTra = String.Empty;
                condTra = (cboxTra.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'T' " : condTra;

                string condCan = String.Empty;
                condCan = (cboxCan.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'C' " : condCan;

                string condDev = String.Empty;
                condDev = (cboxDev.Checked == false) ? "T0." + Constantes.View.CAB_RVT.U_estado + " <> 'D' " : condDev;

                filtrado = filtrado + "\n>Estados Seleccionados: " +
                        (condNue == String.Empty ? "Reservada" : "") + (condApr == String.Empty ? "-Aprobada" : "") + 
                        (condTra == String.Empty ? "-Transferida" : "")+
                        (condCan == String.Empty ? "-Cancelada " : "") + (condDev == String.Empty ? "-Devuelta" : "") ;

                //if (filtrado != "")
                //{
                //    int respuesta = B1.Application.MessageBox("Filtros aplicados: " + filtrado, 1, "OK");
                //}

                B1.Application.SetStatusBarMessage("Cargando datos de Solicitudes...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                string cadw = "";
                cadw = condPer != String.Empty || condCli != String.Empty || condArt != String.Empty || condVend != String.Empty ||
                       condNue != String.Empty || condApr != String.Empty || condTra != String.Empty || condCan != String.Empty  || condDev != String.Empty
                       ? " WHERE " : "";
                cadw = cadw + (condPer  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condPer  : "");
                cadw = cadw + (condCli  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCli  : "");
                cadw = cadw + (condArt  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condArt  : "");
                cadw = cadw + (condVend != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condVend : "");
                cadw = cadw + (condNue  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condNue  : "");
                cadw = cadw + (condApr  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condApr : "");
                cadw = cadw + (condTra  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condTra : "");
                cadw = cadw + (condCan  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCan : "");
                cadw = cadw + (condDev  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condDev : "");

                string adicjoin = (condCli != String.Empty || condArt != String.Empty) ? " INNER JOIN " +
                Constantes.View.DET_RVT.DET_RV + " T3 ON T0." + Constantes.View.CAB_RVT.U_numOC +
                " = T3."  + Constantes.View.DET_RVT.U_numOC : "";

                string adicgroup = (condCli != String.Empty || condArt != String.Empty) ? " GROUP BY " +
                 " T0." + Constantes.View.CAB_RVT.U_numOC +", T0." + Constantes.View.CAB_RVT.U_idVend +" , T1." +
                Constantes.View.ousr.uName + ", T0."  + Constantes.View.CAB_RVT.U_fechaC + ", T0." +
                Constantes.View.CAB_RVT.U_fechaV + ", DAYS_BETWEEN( CURRENT_DATE,T0."  + 
                Constantes.View.CAB_RVT.U_fechaV +" ), T0." + Constantes.View.CAB_RVT.U_idAut +", T2." +
                Constantes.View.ousr.uName +", T0." + Constantes.View.CAB_RVT.U_estado + ", T0." +
                Constantes.View.CAB_RVT.U_idTR +" , T0."  +Constantes.View.CAB_RVT.U_idTV 
                //+",  T0." +   Constantes.View.CAB_RVT.U_comment
                : "" ;              

                SQLQuery = String.Format("SELECT T0.{1} , T0.{4}, T1.{3} U_vend, T0.{6}, T0.{7}, DAYS_BETWEEN( CURRENT_DATE,T0.{7}) U_diasv, " +
                      " T0.{8}, T2.{3} U_aut, T0.{9}, T0.{10}, T0.{11}, CAST(T0.{1} AS INT) AS ND" +
                      //, T0.{22} " +
                      " FROM {0} T0 INNER JOIN {2} T1 ON T0.{4} = T1.{5} " +
                      " LEFT JOIN {2} T2 ON T0.{8} = T2.{5} {24}  {23}  {25}  ORDER BY CAST(T0.{1} AS INT) ",
                                              Constantes.View.CAB_RVT.CAB_RV, //0
                                              Constantes.View.CAB_RVT.U_numOC,//1
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
                                              Constantes.View.DET_RVT.U_codCli, //16
                                              Constantes.View.DET_RVT.U_cliente, //17
                                              Constantes.View.DET_RVT.U_cant, //18
                                              Constantes.View.DET_RVT.U_onHand, //19
                                              Constantes.View.DET_RVT.U_estado, //20
                                              Constantes.View.DET_RVT.U_idTV, //21                                                
                                              Constantes.View.CAB_RVT.U_comment,//22,
                                              cadw, //23,
                                              adicjoin, //24
                                              adicgroup //25
                                              );

                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(SQLQuery);

                AMatrix.Clear();
                SAPbobsCOM.Fields fields = rsCards.Fields;
                rsCards.MoveFirst();
                //B1.Application.Forms.ActiveForm.Freeze(false);
                //SAPbouiCOM.ProgressBar oProgressBar = B1.Application.StatusBar.CreateProgressBar("Cargando datos de Solicitudes...", rsCards.RecordCount, false);

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

                    string txtestado = fields.Item("U_estado").Value.ToString();
                    txtestado = txtestado.Substring(0, 1);
                    SAPbouiCOM.ComboBox mc = (SAPbouiCOM.ComboBox)AMatrix.Columns.Item(9).Cells.Item(i).Specific;

                    mc.Select(txtestado,BoSearchKey.psk_ByValue);

                    if (txtestado=="C" || txtestado =="D")
                    {
                        AMatrix.CommonSetting.SetCellFontColor(i, 9, 255); 
                    }
                    else
                        if (txtestado == "A" || txtestado == "T")
                        {
                            AMatrix.CommonSetting.SetCellFontColor(i, 9, 000102000);
                        }
                        else { AMatrix.CommonSetting.SetCellFontColor(i, 9, 0); }

                    AMatrix.Columns.Item(10).Cells.Item(i).Specific.Value = obtener_DocNum(fields.Item("U_IdTR").Value.ToString());
                    AMatrix.Columns.Item(11).Cells.Item(i).Specific.Value = obtener_DocNum(fields.Item("U_idTV").Value.ToString());
                    AMatrix.Columns.Item(12).Cells.Item(i).Specific.Value = obtener_Comentario(fields.Item("U_numDoc").Value.ToString());


                    rsCards.MoveNext();
                    if (!rsCards.EoF)
                    { B1.Application.SetStatusBarMessage("ESPERE......Cargando datos de Solicitud...No." + fields.Item("U_numDoc").Value.ToString() + "    (" + i.ToString() + "/" + rsCards.RecordCount.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, false); }
                    //try
                    //{
                    //     oProgressBar.Text = "Cargando datos de Solicitudes ...";
                    //}
                    //catch (Exception)
                    //{
                    //    oProgressBar = B1.Application.StatusBar.CreateProgressBar("Cargando datos de Solicitudes...", rsCards.RecordCount, false);
                    //}
                    // oProgressBar.Value = i;

                }
                //oProgressBar.Stop();
                AMatrix.AutoResizeColumns();
                B1.Application.Forms.ActiveForm.Freeze(false);

               

            }
            catch (Exception ex)
            {
                   todoOk = false;
                   serror = ex.Message;
                   throw ex;
            }
            if (todoOk)
            {
                B1.Application.SetStatusBarMessage("Solicitudes cargadas con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            else
            {
                B1.Application.SetStatusBarMessage("Error cargando datos: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }

        }

        private void ThisSapApiForm_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (FormUID == formActual)
                {
                    if (!pVal.BeforeAction)
                    {
                        switch (pVal.EventType)
                        {

                            case BoEventTypes.et_VALIDATE:
                                {
                                    if (pVal.ItemUID == "txtDesde" || pVal.ItemUID == "txtHasta")
                                    {

                                        SAPbouiCOM.CheckBox cboxPer = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxPer).Specific;

                                            SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                                            SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                                            if (cboxPer.Checked && desde.Value.ToString() != "" && hasta.Value.ToString() != "")
                                            {
                                                cargar_datos_matriz();

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
                                                if (oCombo.Value != "") 
                                                {
                                                    if (oCbox.Checked) { cargar_datos_matriz(); }
                                                    else { oCbox.Checked = true; }
                                                }
                                                else { oCbox.Checked = false; }

                                            }
                                            break;
                                        case "cbArt":
                                            {
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxArt").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbArt").Specific;
                                                if (oCombo.Value != "")
                                                {
                                                    if (oCbox.Checked) { cargar_datos_matriz(); }
                                                    else { oCbox.Checked = true; }
                                                }
                                                else { oCbox.Checked = false; }
                                            }
                                            break;
                                        case "cbVend":
                                            {
                                                SAPbouiCOM.CheckBox oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxVend").Specific;
                                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbVend").Specific;
                                                if (oCombo.Value != "")
                                                {
                                                    if (oCbox.Checked) { cargar_datos_matriz(); }
                                                    else { oCbox.Checked = true; }
                                                }
                                                else { oCbox.Checked = false; }
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
                                                    cargar_datos_matriz();
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
                                                    cargar_datos_matriz();
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
                                                    cargar_datos_matriz();
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
                                                    cargar_datos_matriz();
                                                }
                                            }
                                            break;

                                        case "cboxNue":
                                            {
                                                // Desactivar el estado Nueva  
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxApr":
                                            {
                                                // Desactivar el estado Autorizada
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxTra":
                                            {
                                                // Desactivar el estado Transferida  
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxCan":
                                            {
                                                // Desactivar el estado Cancelada  
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxDev":
                                            {
                                                // Desactivar el estado Devuelta  
                                                cargar_datos_matriz();
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
                                    if (pVal.ItemUID == "cboxDev" || pVal.ItemUID == "cboxCan" || pVal.ItemUID == "cboxTra" ||
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
                                                                                        Constantes.View.DET_RVT.DET_RV,
                                                                                        Constantes.View.DET_RVT.U_codCli,
                                                                                        Constantes.View.DET_RVT.U_cliente);

                                                    llenar_combo_busq(oCombo, SQLQuery);
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

                                                    llenar_combo_busq(oCombo, SQLQuery);
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
                                                    llenar_combo_busq(oCombo, SQLQuery);
                                                }
                                                break;

                                            case "mtxaprob":
                                                {
                                                    if (pVal.Row > 0)
                                                    {
                                                        string nodoc = AMatrix.Columns.Item(1).Cells.Item(pVal.Row).Specific.Value;
                                                        new VIEW.PantallaRegistro(this, false, nodoc);
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
                                                        B1.Application.SetStatusBarMessage("Error No se puede filtrar por Período si Desde o Hasta están vacíos", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                B1.Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                BubbleEvent = false;
                throw ex;
            }

        }

        public void llenar_combo_busq(SAPbouiCOM.ComboBox oCombo, string SqlQuery)
        {
            SAPbobsCOM.Recordset oRecordSet = null;

            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(SqlQuery);

            oCombo.ValidValues.Add("", "");

            for (int i = 1; !oRecordSet.EoF;i++ )
            {
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString(), oRecordSet.Fields.Item(1).Value.ToString());
                oRecordSet.MoveNext();

            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);

        }

        private string obtener_DocNum(string dentry)
        {
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
                    B1.Application.SetStatusBarMessage("Error obteniendo DocNum de la Transferencia", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    return dnum;
                    throw ex;

                }
            }
            return dnum;

        }
 
        private string obtener_Comentario(string solnum)
        {
            string dcom = "";
            if (solnum != "")
            {
                try
                {

                    String strSQL = String.Format("SELECT {2} FROM {0} Where {1}='{3}'",
                              Constantes.View.CAB_RVT.CAB_RV,
                              Constantes.View.CAB_RVT.U_numOC,
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
                    B1.Application.SetStatusBarMessage("Error obteniendo Comentarios de la Solicitud", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    return dcom;
                    throw ex;

                }
            }
            return dcom;

        }

        private double obtener_exist_articulo(string codart, string codwhs)
        {
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
                B1.Application.SetStatusBarMessage("Error obteniendo Stock Disponible" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
            return exist;
        }

        private void revertir(string sCode, string docentry)
        {
            bool todoOk = true;
            int result = 0;
            string terror = "";
            try
            {
                GC.Collect();
                B1.Company.StartTransaction();
                SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                doctransf.DocDate = DateTime.Today;
                doctransf.TaxDate = DateTime.Today;
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
                                disponible = obtener_exist_articulo(artcurrent, "CD_RSV");
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
                        disponible = obtener_exist_articulo(artcurrent, "CD_RSV");
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
                        todoOk = (result == 0) && (linestransf > 0);
                    }
                    else
                    {
                        string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                        terror = "No existen artículos disponibles. " + infonodisp;
                        todoOk = false;
                    }
                    GC.Collect();


                }

                if (todoOk)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    string newkey = B1.Company.GetNewObjectKey();
                    if (newkey != "")
                    {
                        //Actualizar datos de Transferencia en Solicitud
                        string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                        string slog = "Cancelada Automáticamente: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + " DocNum:" + obtener_DocNum(newkey) + infonodisp;
                        string scom = "Solicitud Cancelada por vencer su período de revisión: " + DateTime.Now.Date.ToString("dd/MM/yyyy");

                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string sestado = "D";

                        string SQLQuery = String.Format("UPDATE {1} SET {2} = '{8}', {5} = '{4}', {6} = '{7}', {9} = '{10}' FROM {1} WHERE {0} = '{3}' ",
                                                 Constantes.View.CAB_RVT.U_numOC,   //0
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
                        cancelar_filas_nodisp(newkey);
                        B1.Application.SetStatusBarMessage("Solicitud Cancelada Automática Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud Cancelada Automática", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    }
                }
                else
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    B1.Application.SetStatusBarMessage("Error Transferiendo Solicitud Cancelada Automática " +sCode +" "+terror, SAPbouiCOM.BoMessageTime.bmt_Long, true);

                    //Actualizar logs en Solicitud
                    string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                    string slog = "Error:No pudo ser Cancelada Automáticamente por no tener disponibilidad: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + infonodisp;
                    string scom = "Solicitud sin disponibilidad al intentar Cancelada por vencer su período de revisión: " + DateTime.Now.Date.ToString("dd/MM/yyyy");

                    Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                    string SQLQuery = String.Format("UPDATE {1} SET {2} = '{6}', {5} = '{4}'  FROM {1} WHERE {0} = '{3}' ",
                                             Constantes.View.CAB_RVT.U_numOC,   //0
                                             Constantes.View.CAB_RVT.CAB_RV,    //1
                                             Constantes.View.CAB_RVT.U_logs,    //2
                                             sCode,                             //3
                                             scom,                              //4
                                             Constantes.View.CAB_RVT.U_comment,    //5
                                             slog);//6

                    oRecordSet.DoQuery(SQLQuery);


                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Transferiendo Solicitud Cancelada Automática" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, true);
                throw ex;
            }
        }

        private bool cancelar_filas_nodisp(string newkey)
        {
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
                B1.Application.SetStatusBarMessage("Error sincronizando artículos no disponibles " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                todoOk = false;
                throw;
            }

            finally
            {
                lineasnodisp.Clear();
                System.GC.Collect();
            }

            return todoOk;
        }

        private void cancelar_vencidaspormas10D()
        {
            
            try
            {

                //Actualizar estado y comentario
                B1.Application.SetStatusBarMessage("Realizando Cancelación Automática por Fecha de Vencimiento", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                B1.Application.Forms.ActiveForm.Freeze(true);
                string nodoc = "";
                string dentry = "";
                //DAYS_BETWEEN(CURRENT_DATE,{3}) < 0
                string SQLQuery = String.Format("SELECT {0}, {4} FROM {1} WHERE {2} = 'R' AND  DAYS_BETWEEN(CURRENT_DATE,{3}) < 0 ",
                                             Constantes.View.CAB_RVT.U_numOC,
                                             Constantes.View.CAB_RVT.CAB_RV,
                                             Constantes.View.CAB_RVT.U_estado,
                                             Constantes.View.CAB_RVT.U_fechaV,
                                             Constantes.View.CAB_RVT.U_idTR);


                Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(SQLQuery);
                oRecordSet.MoveFirst();
                for (int i = 0; !oRecordSet.EoF ; i++)
                {
                    nodoc = oRecordSet.Fields.Item("U_numDoc").Value.ToString();
                    dentry = oRecordSet.Fields.Item("U_idTR").Value.ToString();
                    revertir(nodoc, dentry);
                    oRecordSet.MoveNext();
                }

                B1.Application.Forms.ActiveForm.Freeze(false);
                B1.Application.SetStatusBarMessage("Cancelación Automática realizada con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error al realizar Cancelación automática:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }

        }
     
    }

}
