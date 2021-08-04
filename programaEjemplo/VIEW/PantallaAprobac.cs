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
        private string ItemActiveMenu = "";

        private string formActual = "";
        SAPbouiCOM.Form AForm = null;
        SAPbouiCOM.Matrix AMatrix = null;

        private int rowsel = 0;   



        public PantallaAprobac()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Aprobac.srf"], "AprRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString())
        {
            formActual = "AprRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();

            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);

            cargar_datos_matriz();
        }

        private void cargar_datos_matriz()
        {
            bool todoOk = true;
            string serror = "";
            formActual = B1.Application.Forms.ActiveForm.UniqueID;
            AForm = B1.Application.Forms.ActiveForm;
            AMatrix = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item("mtxaprob").Specific;
            try
            {
                B1.Application.Forms.ActiveForm.Freeze(true);
                string SQLQuery = string.Empty;
                //SQLQuery = String.Format("SELECT T0.{1} , T1.{3} IdVend, T0.{6}, T0.{7}, DAYS_BETWEEN(T0.{7}, CURRENT_DATE) dias, " + 
                //    " T2.{3} IdAut, T0.{9}, T0.{10}, T0.{11} , T3.{14} , T3.{15}, T3.{16}, T3.{17}, "+
                //    " T3.{18} , T3.{19}, T3.{20}, T3.{21}, T0.{22} "+
                //    " FROM {0} T0 INNER JOIN {2} T1 ON T0.{4} = T1.{5} " +
                //    " LEFT JOIN {2} T2 ON T0.{8} = T2.{5} " +
                //    " INNER JOIN {12} T3 ON T0.{1} = T3.{13} ",
                //                            Constantes.View.CAB_RVT.CAB_RV, //0
                //                            Constantes.View.CAB_RVT.U_numOC,//1
                //                            Constantes.View.ousr.OUSR, //2
                //                            Constantes.View.ousr.uName, //3
                //                            Constantes.View.CAB_RVT.U_idVend,//4
                //                            Constantes.View.ousr.uCode, //5
                //                            Constantes.View.CAB_RVT.U_fechaC, //6
                //                            Constantes.View.CAB_RVT.U_fechaV, //7
                //                            Constantes.View.CAB_RVT.U_idAut, //8
                //                            Constantes.View.CAB_RVT.U_estado, //9
                //                            Constantes.View.CAB_RVT.U_idTR, //10
                //                            Constantes.View.CAB_RVT.U_idTV, //11
                //                            Constantes.View.DET_RVT.DET_RV, //12
                //                            Constantes.View.DET_RVT.U_numOC, //13
                //                            Constantes.View.DET_RVT.U_codArt, //14
                //                            Constantes.View.DET_RVT.U_articulo, //15
                //                            Constantes.View.DET_RVT.U_codCli, //16
                //                            Constantes.View.DET_RVT.U_cliente, //17
                //                            Constantes.View.DET_RVT.U_cant, //18
                //                            Constantes.View.DET_RVT.U_onHand, //19
                //                            Constantes.View.DET_RVT.U_estado, //20
                //                            Constantes.View.DET_RVT.U_idTV, //21                                                
                //                            Constantes.View.CAB_RVT.U_comment//22
                //                            );


                SAPbouiCOM.CheckBox cboxPer = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxPer).Specific;
                string condPer = String.Empty;
                if (cboxPer.Checked == true)
                {
                    SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                    SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                    condPer = desde.Value!="" && hasta.Value != "" ?Constantes.View.CAB_RVT.U_fechaC + " between '" + desde.Value + "' AND ' " + hasta.Value + "'":"" ;
                }

                SAPbouiCOM.CheckBox cboxCli = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxCli).Specific;
                string condCli = String.Empty;
                if (cboxCli.Checked == true)
                {
                    SAPbouiCOM.ComboBox cli = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbCli).Specific;
                    string selCli = cli.Selected != null  ? cli.Selected.Value.ToString() : "";
                    condCli = selCli != "" ? Constantes.View.DET_RVT.U_codCli + " = '" + selCli + "'": condCli;
                }

                SAPbouiCOM.CheckBox cboxArt = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxArt).Specific;
                string condArt = String.Empty;
                if (cboxArt.Checked == true)
                {
                    SAPbouiCOM.ComboBox art = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbArt).Specific;
                    string selArt = art.Selected != null ? art.Selected.Value.ToString() : "";
                    condArt = selArt != "" ? Constantes.View.DET_RVT.U_codArt + " = '" + selArt + "'": condArt;
                }

                SAPbouiCOM.CheckBox cboxVend = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cboxVend).Specific;
                string condVend = String.Empty;
                if (cboxVend.Checked == true)
                {
                    SAPbouiCOM.ComboBox vend = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.cbVend).Specific;
                    string selVend = vend.Selected != null ? vend.Selected.Value.ToString() : "";
                    condVend = selVend != "" ? Constantes.View.CAB_RVT.U_idVend + " = '" + selVend + "'" : condVend;
                }

                string cadw = "";
                cadw = condPer != String.Empty || condCli != String.Empty || condArt != String.Empty || condVend != String.Empty ? " WHERE " : "";
                cadw = cadw + (condPer  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condPer  : "");
                cadw = cadw + (condCli  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condCli  : "");
                cadw = cadw + (condArt  != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condArt  : "");
                cadw = cadw + (condVend != String.Empty ? (cadw.Length == 7 ? "" : " AND ") + condVend : "");

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
                      " T0.{8}, T2.{3} U_aut, T0.{9}, T0.{10}, T0.{11} " +
                //,T0.{22} " +
                      " FROM {0} T0 INNER JOIN {2} T1 ON T0.{4} = T1.{5} " +
                      " LEFT JOIN {2} T2 ON T0.{8} = T2.{5} {24}  {23}  {25}",
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
                for (int i = 1; !rsCards.EoF; i++)
                {
                    AMatrix.AddRow(1);
                    AMatrix.Columns.Item(1).Cells.Item(i).Specific.Value = fields.Item("U_numDoc").Value.ToString();
                    AMatrix.Columns.Item(2).Cells.Item(i).Specific.Value = fields.Item("U_IdVend").Value.ToString();
                    AMatrix.Columns.Item(3).Cells.Item(i).Specific.Value = fields.Item("U_vend").Value.ToString();
                    AMatrix.Columns.Item(4).Cells.Item(i).Specific.Value = fields.Item("U_fechaC").Value.ToString("yyyyMMdd");
                    AMatrix.Columns.Item(5).Cells.Item(i).Specific.Value = fields.Item("U_fechaV").Value.ToString("yyyyMMdd");
                    AMatrix.Columns.Item(6).Cells.Item(i).Specific.Value = fields.Item("U_diasv").Value.ToString();
                    if (Int32.Parse(fields.Item("U_diasv").Value.ToString()) <= 5) 
                    { 
                        AMatrix.CommonSetting.SetCellFontColor(i, 6, 255); 
                    }
                    else
                    {
                        AMatrix.CommonSetting.SetCellFontColor(i, 6, 0); 
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
                    AMatrix.Columns.Item(11).Cells.Item(i).Specific.Value = fields.Item("U_idTV").Value.ToString();
                    //AMatrix.Columns.Item(12).Cells.Item(i).Specific.Value = fields.Item("U_comment").Value.ToString();


                    rsCards.MoveNext();
                }
                AMatrix.AutoResizeColumns();
                SAPbouiCOM.Column oColumn = AMatrix.Columns.Item("numDoc");
                oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);
                
                B1.Application.Forms.ActiveForm.Freeze(false);
            }
            catch (Exception ex)
            {
                   todoOk = false;
                   serror = ex.Message;
                   throw;
            }
            if (todoOk)
            {
                B1.Application.SetStatusBarMessage("Articulos cargados con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
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
                                        if (cboxPer.Checked)
                                        {
                                            SAPbouiCOM.EditText desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtDesde).Specific;
                                            SAPbouiCOM.EditText hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.aprobac.txtHasta).Specific;
                                            if (desde.Value.ToString() != " " && hasta.Value.ToString() != "")
                                            {
                                                cargar_datos_matriz();
                                            }
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
                                                cargar_datos_matriz();
                                            }
                                            break;
                                        case "cbArt":
                                            {
                                                cargar_datos_matriz();
                                            }
                                            break;
                                        case "cbVend":
                                            {
                                                cargar_datos_matriz();
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
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = null;
                                                oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxPer").Specific;
                                                SAPbouiCOM.EditText desde = null;
                                                desde = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtDesde").Specific;
                                                SAPbouiCOM.EditText hasta = null;
                                                hasta = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txtHasta").Specific;
                                                //desde.Item.Enabled = !oCbox.Checked;
                                                //hasta.Item.Enabled = !oCbox.Checked;
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxArt":
                                            {
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = null;
                                                oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxArt").Specific;
                                                SAPbouiCOM.ComboBox oCombo = null;
                                                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbArt").Specific;
                                                //oCombo.Item.Visible = !oCbox.Checked;
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxVend":
                                            {
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = null;
                                                oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxVend").Specific;
                                                SAPbouiCOM.ComboBox oCombo = null;
                                                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbVend").Specific;
                                                //oCombo.Item.Visible = !oCbox.Checked;
                                                cargar_datos_matriz();
                                            }
                                            break;

                                        case "cboxCli":
                                            {
                                                // Activar busqueda por articulo  
                                                SAPbouiCOM.CheckBox oCbox = null;
                                                oCbox = (SAPbouiCOM.CheckBox)B1.Application.Forms.ActiveForm.Items.Item("cboxCli").Specific;
                                                SAPbouiCOM.ComboBox oCombo = null;
                                                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbCli").Specific;
                                                //oCombo.Item.Visible = !oCbox.Checked;
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
                                                string nodoc = AMatrix.Columns.Item(1).Cells.Item(pVal.Row).Specific.Value;
                                                new VIEW.PantallaRegistro(false, nodoc);
                                            }
                                            break;

                                    }
                                }
                                break;

                            case BoEventTypes.et_ITEM_PRESSED:
                                {

                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.registro.btn_crear:
                                            {
                                                //SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                                //if (btn_crear.Caption == "Actualizar")
                                                //{
                                                //    Guardar_Solicitud();
                                                //    BubbleEvent = false;
                                                //}

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
                                                B1.Application.SetStatusBarMessage("Error Fecha Desde <= Fecha Hasta", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                BubbleEvent = false;
                                            }
                                    }
                                    break;

                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
     
    }

}
