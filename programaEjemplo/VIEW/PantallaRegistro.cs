using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SAPbobsCOM;
using SSIFramework;
using SSIFramework.DI.Attributes;
using SSIFramework.Utilidades;
using System.Threading;


namespace ventaRT.VIEW
{
    class PantallaRegistro : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        private string ItemActiveMenu = "";

        private string formActual = "";
        SAPbouiCOM.Form SForm = null;
        SAPbouiCOM.Matrix SMatrix = null;

        SAPbouiCOM.DBDataSource oDbLinesDataSource = null;
        SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;


        List<string> lineasdel = new List<string>();
       

        public PantallaRegistro()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Registro.srf"], "ventaRT_Registro" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            formActual = "ventaRT_Registro" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();

            
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);

            
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);

 
            cargar_info_inicial();
        }

       

        // Metodos Override

        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
              BubbleEvent = true;
              if (B1.Application.Forms.ActiveForm.UniqueID == formActual)
              {
                if (pVal.BeforeAction)
                {
                    BubbleEvent = true;
                    switch (pVal.MenuUID)
                    {
                        case "1282":    // Crear      
                            insertar_solicitud();
                            BubbleEvent = false;
                            break;
                        case "1281":    // Buscar                      
                            preparar_modo_Find();
                            BubbleEvent = false;
                            break;
                        case "1283":    // Eliminar                     
                            eliminar_solicitud();
                            BubbleEvent = false;
                            break;
                        case "1292":   //ADICIONAR LINEA
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.registro.mtx:
                                    SMatrix.AddRow(1, SMatrix.RowCount);
                                    SMatrix.ClearRowData(SMatrix.RowCount);
                                    SMatrix.FlushToDataSource();
                                    SMatrix.LoadFromDataSource();
                                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                    btn_crear.Caption = "Actualizar";
                                    SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                        case "1293":  //BORRAR LINEA
                            switch (ItemActiveMenu)
                            {
                                //ejemplo con una matrix 
                                case ventaRT.Constantes.View.registro.mtx:

                                    int nRow = (int)SMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    nRow = nRow == -1 ? SMatrix.RowCount : nRow ;
                                    if (nRow > 0)
                                    {

                                        SMatrix.GetLineData(nRow);
                                        string lindel = oDbLinesDataSource.GetValue("code", nRow- 1);
                                        //oDbLinesDataSource.Offset
                                        lineasdel.Add(lindel);
                                        SMatrix.DeleteRow(nRow);
                                        SMatrix.FlushToDataSource();
                                        SMatrix.LoadFromDataSource();
                                        SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                        btn_crear.Caption = "Actualizar";
                                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    }
                                    BubbleEvent = false;

                                    break;
                            }
                            break;
                        case "1290":    // Primero                      
                            activar_primero();
                            break;
                        case "1289":    // Ant                      
                            activar_anterior();
                            break;
                        case "1288":    // Sig                      
                            activar_posterior();
                            break;
                        case "1291":    // Ultimo                      
                            activar_ultimo();
                            break;
                    }
                    BubbleEvent = false;
                }
              }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Ejecutando Menu" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }


        private void ThisSapApiForm_OnAfterRightClick(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (eventInfo.FormUID == formActual)
                {
                    ItemActiveMenu = eventInfo.ItemUID;
                    if (eventInfo.BeforeAction && eventInfo.ItemUID == ventaRT.Constantes.View.registro.mtx)
                    {
                        SForm.EnableMenu("1292", true); //Activar Agregar Linea
                        SForm.EnableMenu("1293", true); //Activar Borrar Linea 
                    }
                    else
                    {
                        SForm.EnableMenu("1292", false); //Desctivar Agregar Linea
                        SForm.EnableMenu("1293", false); //Desactivar Borrar Linea 
                    }
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Activando Opciones Menu" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
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

                            case BoEventTypes.et_ITEM_PRESSED:
                                {

                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.registro.btn_crear:
                                            {
                                                switch (B1.Application.Forms.ActiveForm.Mode)
                                                {
                                                    case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                                                        {
                                                            SAPbouiCOM.ComboBox oCombox = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                                            string buscado = oCombox.Selected.Value == null ? " 0" : oCombox.Selected.Value.ToString();
                                                            if (buscado != "0")
                                                            {
                                                                cargar_solicitud(buscado, true);
                                                            }
                                                            BubbleEvent = false;
                                                            break;
                                                        }
                                                    case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                                                        {
                                                            guardar_solicitud();
                                                            BubbleEvent = false;
                                                            break;
                                                        }
                                                    case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                                                        {
                                                            guardar_solicitud();
                                                            BubbleEvent = false;
                                                            break;
                                                        }
                                                }
                                                break;
                                            }

                                    }
                                    break;

                                }

                            case BoEventTypes.et_VALIDATE:
                                {
                                    if (pVal.InnerEvent == false && pVal.ItemUID == "mtx" && pVal.ColUID == "cant")
                                    {
                                        string codArt = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        string codCli = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codCli").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        if (codArt != "" && codCli != "" && pVal.Row == SMatrix.RowCount)
                                        {
                                            string tempnum = SMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific.Value.ToString();
                                            if (Double.Parse(tempnum) == 0.00)
                                            { SMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific.Value = "1"; }
                                            SMatrix.AddRow(1, pVal.Row);
                                            SMatrix.ClearRowData(SMatrix.RowCount);
                                            SMatrix.FlushToDataSource();
                                            SMatrix.LoadFromDataSource();
                                            //SMatrix.Columns.Item("codArt").Cells.Item(SMatrix.RowCount).Click(BoCellClickType.ct_Double, 0);
                                            //SMatrix.Columns.Item(5).Cells.Item(pVal.Row + 1).Specific.Value = "1";


                                        }
                                    }


                                }
                                break;

                            case BoEventTypes.et_CHOOSE_FROM_LIST:
                                {
                                    if (pVal.InnerEvent == true)
                                    {

                                        SAPbouiCOM.ChooseFromList oCFL;

                                        SAPbouiCOM.IChooseFromListEvent CFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;

                                        string CFL_Id = CFLEvent.ChooseFromListUID;
                                        oCFL = SForm.ChooseFromLists.Item(CFL_Id);
                                        if (pVal.FormTypeEx.Substring(0, 10) == "ventaRT_Re" && CFLEvent.SelectedObjects != null)
                                        {
                                            if (pVal.ItemUID == "mtx" && pVal.ColUID == "codArt")
                                            {
                                                bool Ok = true;
                                                string artsel = CFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString();
                                                string codcli = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codCli").Cells.Item(pVal.Row).Specific).Value.ToString();
                                                // Validar que no existan repetidos earticulo y cliente en el documento
                                                if (artsel != "" && codcli != "" && !validar_art_cliente_unicos(artsel, codcli, pVal.Row))
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error Datos Repetidos: Articulo y Cliente deben ser unicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                    BubbleEvent = false;
                                                }
                                                // Validar que tenga existencia en la Bodega Principal CD
                                                if (Ok)
                                                {
                                                    if (!(obtener_exist_articulo(artsel) > 0))
                                                    {
                                                        Ok = false;
                                                        B1.Application.SetStatusBarMessage("Error el Articulo no tienen disponibilidad en la Bodega Principal", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                }
                                                if (Ok)
                                                {
                                                    int nRow = (int)SMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                                    nRow = nRow == -1 ? pVal.Row : nRow - 1;
                                                    SMatrix.FlushToDataSource();
                                                    oDbLinesDataSource.SetValue("U_CodArt", nRow - 1, artsel);
                                                    //oDbLinesDataSource.Offset
                                                    oDbLinesDataSource.SetValue("U_articulo", nRow - 1, CFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString());
                                                    oDbLinesDataSource.SetValue("U_cant", nRow - 1, obtener_exist_articulo(artsel).ToString());
                                                    oDbLinesDataSource.SetValue("U_onHand", nRow - 1, obtener_exist_articulo(artsel).ToString());
                                                    SMatrix.LoadFromDataSource();
                                                    SMatrix.Columns.Item("codCli").Cells.Item(nRow).Click();
                                                }
                                            }
                                            if (pVal.ItemUID == "mtx" && pVal.ColUID == "codCli")
                                            {
                                                bool Ok = true;
                                                string codart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                                string clisel = CFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString();
                                                if (codart != "" && clisel != "" && !validar_art_cliente_unicos(codart, clisel, pVal.Row))
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error Datos Repetidos: Articulo y Cliente deben ser unicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                    BubbleEvent = false;
                                                }
                                                if (Ok)
                                                {
                                                    int nRow = (int)SMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                                    nRow = nRow == -1 ? pVal.Row : nRow - 1;
                                                    SMatrix.FlushToDataSource();
                                                    oDbLinesDataSource.SetValue("U_CodCli", nRow - 1, CFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString());
                                                    oDbLinesDataSource.SetValue("U_cliente", nRow - 1, CFLEvent.SelectedObjects.GetValue("CardName", 0).ToString());
                                                    SMatrix.LoadFromDataSource();

                                                    SMatrix.Columns.Item("cant").Cells.Item(nRow).Click();
                                                }

                                            }
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    else
                    {
                        // Antes de Accion

                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                {
                                    // Rellenando combo de busqueda
                                    if (pVal.ItemUID == "cbnd")
                                    {
                                        SAPbouiCOM.ComboBox oCombo = null;
                                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                        string SQLQuery = string.Empty;
                                        SQLQuery = String.Format("SELECT {0}, {2} FROM {1}",
                                                                            Constantes.View.CAB_RVT.Code,
                                                                            Constantes.View.CAB_RVT.CAB_RV,
                                                                            Constantes.View.CAB_RVT.U_fechaC);

                                        llenar_combo_id(oCombo, SQLQuery);
                                    }
                                }
                                break;

                            case BoEventTypes.et_ITEM_PRESSED:
                                {

                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.registro.btn_crear:
                                            {
                                                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                                if (btn_crear.Caption == "Actualizar")
                                                {
                                                    guardar_solicitud();
                                                    BubbleEvent = false;
                                                }

                                            }
                                            break;
                                    }
                                }
                                break;


                            case BoEventTypes.et_VALIDATE:
                                {
                                    if (pVal.InnerEvent == false && pVal.ItemUID == "mtx")
                                    {

                                        string codart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        string codcli = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codCli").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        switch (pVal.ColUID)
                                        {
                                            case "codArt":
                                                {
                                                    if (codart == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error Codigo Articulo es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (codart != "" && codcli != "" && !validar_art_cliente_unicos(codart, codcli, pVal.Row))
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error Datos Repetidos: Articulo y Cliente deben ser unicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = false;
                                                        }
                                                    }
                                                }
                                                break;
                                            case "codCli":
                                                {
                                                    if (codcli == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error Codigo Cliente es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (codart != "" && codcli != "" && !validar_art_cliente_unicos(codart, codcli, pVal.Row))
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error Datos Repetidos: Articulo y Cliente deben ser unicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = false;
                                                        }
                                                    }

                                                }
                                                break;
                                            case "cant":
                                                {
                                                    double cantidad = Double.Parse(((SAPbouiCOM.EditText)SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Specific).Value.ToString());
                                                    double disp = Double.Parse(((SAPbouiCOM.EditText)SMatrix.Columns.Item("onHand").Cells.Item(pVal.Row).Specific).Value.ToString());
                                                    if (cantidad == 0 && disp != 0)
                                                    {
                                                        //SMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific.Value = disp.ToString(); 
                                                        B1.Application.SetStatusBarMessage("Error Cantidad debe ser superior a 0", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (cantidad > disp)
                                                        {
                                                            //SMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific.Value = disp.ToString();
                                                            B1.Application.SetStatusBarMessage("Error Cantidad > Disponibilidad", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = false;
                                                        }

                                                    }

                                                }
                                                break;
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
                throw ex;
            }

        }

         
        // Metodos No Override

        private void cargar_info_inicial()
        {
            //SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
            //oCombo.Item.Visible = false;
            SForm = B1.Application.Forms.ActiveForm;
            SMatrix = SForm.Items.Item("mtx").Specific;
            formActual = B1.Application.Forms.ActiveForm.UniqueID;

            SForm.EnableMenu("1290", true); SForm.EnableMenu("1289", true);
            SForm.EnableMenu("1288", true); SForm.EnableMenu("1291", true);

            SForm.EnableMenu("1282", true); SForm.EnableMenu("1283", true);
            SForm.EnableMenu("1281", true);

            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
            oDbLinesDataSource = SForm.DataSources.DBDataSources.Item("@DET_RSTV");

            // FILTRAR LAS SOLICITUDES DEL USUARIO ACTUAL
            SAPbouiCOM.Conditions orCons = new SAPbouiCOM.Conditions();
            SAPbouiCOM.Condition orCon = orCons.Add();
            orCon.Alias = "U_idVend" ;
            orCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            orCon.CondVal = B1.Company.UserName;
            oDbHeaderDataSource.Query(orCons);
            if (oDbHeaderDataSource.Size >0)
            {
                oDbHeaderDataSource.Offset = 0;
                //oDbHeaderDataSource.Query();


                // FILTRAR LAS LINES DE SOLICITUD ACTUAL
                SAPbouiCOM.Conditions olCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition olCon = olCons.Add();
                olCon.Alias = "U_numOC";
                olCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                olCon.CondVal = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                oDbLinesDataSource.Query(olCons);
                SMatrix.LoadFromDataSource();


                //cargar_lineas(oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset));
            }

            //if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            //{
            //    preparar_modo_Find();
            //}


        }

        private bool insertar_solicitud()
        {

            bool todoOk = true;
            string serror = "";            
            try {
                    SForm = B1.Application.Forms.ActiveForm;
                    SMatrix = SForm.Items.Item("mtx").Specific;

                    B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                        int norecord = obtener_ultimo_ID("CA") + 1;
               
                        //Insertando nuevo record

                        oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size - 1;
                        oDbHeaderDataSource.Query();
                        oDbHeaderDataSource.InsertRecord(oDbHeaderDataSource.Size);
                        oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size-1;

                        DateTime fc = DateTime.Now.Date;
                        DateTime fv = fc.AddDays(10);

                        

                        oDbHeaderDataSource.SetValue("U_numDoc", norecord, norecord.ToString());
                        oDbHeaderDataSource.SetValue("U_IdVend", norecord, obtener_Vendedor());
                        oDbHeaderDataSource.SetValue("U_vend", norecord, obtener_NameVendedor());
                        oDbHeaderDataSource.SetValue("U_fechaC", norecord, fc.ToString("yyyyMMdd"));
                        oDbHeaderDataSource.SetValue("U_fechaV", norecord, fv.ToString("yyyyMMdd"));
                        oDbHeaderDataSource.SetValue("U_estado", norecord, "Nueva");
                        oDbHeaderDataSource.SetValue("U_comment", norecord, "");


                        SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                        SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                        SAPbouiCOM.EditText txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                        SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_aut).Specific;
                        SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                        SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                        SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                        SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                        SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                        SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                        SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                        SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                        SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;

                        txt_numoc.Value = norecord.ToString();
                        txt_idvend.Value = obtener_Vendedor();
                        txt_vend.Value = obtener_NameVendedor(); 
                        txt_idaut.Value = "";
                        txt_aut.Value = "";
                        txt_idtv.Value = "";
                        txt_idtr.Value = "";
                        txt_fechac.Value = fc.ToString("yyyyMMdd");
                        txt_fechav.Value = fv.ToString("yyyyMMdd");
                        txt_com.Value = "";
                        txt_estado.Value = "Nueva" ;
                        btn_crear.Caption = "Crear";
                        mtx.Clear();
                        mtx.AddRow(1, 1);
                        mtx.ClearRowData(1);
                    
                        //mtx.Columns.Item(5).Cells.Item(1).Specific.Value = "1";
  
                }
                catch (Exception ex)
                {
                    todoOk = false;
                    B1.Application.SetStatusBarMessage("Error creando solicitud: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }
                return todoOk;
        }

        private bool preparar_modo_Find()
        {
            bool todoOk = true;
            int borrado = 0;


            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede activar Modo Busqueda porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;
            }
            else
            {
                try
                {
                    if (B1.Company.InTransaction || SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || SForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        int respuesta = B1.Application.MessageBox("Desea cancelar los datos modificados? ", 1, "OK", " Cancelar");
                        if (respuesta == 1)
                        {
                            if (B1.Company.InTransaction)
                            {
                                B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }


                            if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {

                                oDbHeaderDataSource.RemoveRecord(oDbHeaderDataSource.Size - 1);
                                borrado = 1;
                                oDbHeaderDataSource.Offset = 0;
                                oDbHeaderDataSource.Query();
                            }
                            todoOk = true;
                        }
                        else { todoOk = false; }

                    }
                    if (todoOk)
                    {


                        SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                        SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                        SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                        SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                        SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                        SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                        SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                        SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                        
                        SAPbouiCOM.EditText txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                        SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;

                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                        btn_crear.Caption = "Buscar";
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;


                        oCombo.Item.Visible = true;
                        oCombo.Active = true;
                        txt_idvend.Value = "";
                        txt_vend.Value = "";
                        txt_numoc.Value = "";
                        txt_fechac.Value = "";
                        txt_fechav.Value = "";
                        txt_estado.Value = "";
                        txt_com.Value = "";
                        txt_idaut.Value = "";
                        txt_aut.Value = "";


                        //oCombo.Item.Enabled = true;
                        //SMatrix.Item.Enabled = false;
                        //txt_com.Item.Enabled = false;

                        //oCombo.Active = true;
                        //SMatrix.Item.Enabled = false;
                        //txt_com.Item.Enabled = false;




                    }
                }
                catch (Exception ex)
                {
                    todoOk = false;
                    B1.Application.SetStatusBarMessage("Error preparando busqueda: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }
                return todoOk;
            }
        
 
        }

        private void activar_primero()
        {

            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {

                    //oDbHeaderDataSource.Offset = 0;
                    //oDbHeaderDataSource.Query();
                    //Cargar_Solicitud(oDbHeaderDataSource.GetValue("U_numDoc", 0), false);
                    cargar_solicitud("0", false);
                    B1.Application.SetStatusBarMessage("Movimiento al Primero ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {

                }
            }
         }

        private void activar_anterior()
        {


            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    if (oDbHeaderDataSource.Offset > 0)
                    {
                        oDbHeaderDataSource.Offset--;
                       // oDbHeaderDataSource.Query();
                       // Cargar_Solicitud(oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset), false);

                        cargar_solicitud(oDbHeaderDataSource.Offset.ToString(), false);
                    }

                    B1.Application.SetStatusBarMessage("Movimiento al Anterior ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error en Movimiento al Ultimo ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    throw ex;
                }  
            }
  
        }

        private void activar_posterior()
        {

            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    oDbHeaderDataSource.Offset++;
                    //oDbHeaderDataSource.Query();
                    //Cargar_Solicitud(oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset), false);
                    cargar_solicitud(oDbHeaderDataSource.Offset.ToString(), false);

                    B1.Application.SetStatusBarMessage("Movimiento al Siguiente ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {

                }
            }
  


        }

        private void activar_ultimo()
      {

            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size - 1;
                    //oDbHeaderDataSource.Query();
                    //Cargar_Solicitud(oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset), false);
                    cargar_solicitud(oDbHeaderDataSource.Offset.ToString(), false);
                    B1.Application.SetStatusBarMessage("Movimiento al Ultimo ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {

                }
            }



      }

        private bool eliminar_solicitud()
        {
            bool todoOk = true;
            string serror = "";

            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede eliminar porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;
            }
            else
            {
                try
                {
                    // Eliminar documento 
                    SAPbouiCOM.EditText snumOC = (SAPbouiCOM.EditText)SForm.Items.Item("txt_numoc").Specific;
                    string abuscar = snumOC.Value.ToString();

                    string SQLQuery = String.Format("SELECT {0} FROM {1}",
                                        Constantes.View.CAB_RVT.U_numOC,
                                        Constantes.View.CAB_RVT.CAB_RV);

                    Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oRecordSet.DoQuery(SQLQuery);
                    oRecordSet.MoveFirst();
                    bool encontrado = false;
                    int i;
                    for (i = 0; !oRecordSet.EoF && !encontrado; i++)
                    {
                        encontrado = oRecordSet.Fields.Item("U_numDoc").Value.ToString() == abuscar;
                        oRecordSet.MoveNext();
                    }

                    if (encontrado)
                    {
                        oDbHeaderDataSource.RemoveRecord(i - 1);

                        SQLQuery = String.Format("DELETE FROM {1} WHERE {0} = '{2}' ",
                                        Constantes.View.CAB_RVT.U_numOC,
                                        Constantes.View.CAB_RVT.CAB_RV,
                                        abuscar);
                        oRecordSet.DoQuery(SQLQuery);

                        // Borrar lineas detalle

                        SQLQuery = String.Format("DELETE FROM {1} WHERE {0} = '{2}' ",
                                        Constantes.View.DET_RVT.U_numOC,
                                        Constantes.View.DET_RVT.DET_RV,
                                        abuscar);
                        oRecordSet.DoQuery(SQLQuery);

                        if (oDbHeaderDataSource.Offset == 0) { activar_primero(); }
                        else { activar_anterior(); }
                    }
                    else
                    {
                        todoOk = false;
                        serror = "Documento No Encontrado";
                    }
                }
                catch (Exception ex)
                {
                    serror = ex.Message;
                    todoOk = false;
                    throw ex;
                }
                if (todoOk)
                {
                    B1.Application.SetStatusBarMessage("Solicitud eliminada con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                else
                {
                    B1.Application.SetStatusBarMessage("Error eliminando solicitud: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }
                return todoOk;
            }

        }

        private bool guardar_solicitud()
        {
            bool todoOk = true;
            string serror = "";
            string sCode = ""; string sName = "";
            int iRet;
            try
            {
 

                SAPbobsCOM.UserTable UTDoc = B1.Company.UserTables.Item("CAB_RSTV");
                SAPbobsCOM.UserTable UTLines = B1.Company.UserTables.Item("DET_RSTV");
                //SForm.Freeze(true);

                try {
                      // Salvando documento 
                        SAPbouiCOM.EditText snumOC = (SAPbouiCOM.EditText)SForm.Items.Item("txt_numoc").Specific;

                        int norecord =  Int32.Parse(snumOC.Value.ToString());
                        sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                        string sfechav = oDbHeaderDataSource.GetValue("U_fechaV",oDbHeaderDataSource.Offset);
                        string sfechac = oDbHeaderDataSource.GetValue("U_fechaC", oDbHeaderDataSource.Offset);
                        string sestado = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                        string scom = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                        string svend = oDbHeaderDataSource.GetValue("U_idVend", oDbHeaderDataSource.Offset);
                        string snvend = oDbHeaderDataSource.GetValue("U_vend", oDbHeaderDataSource.Offset);
                        string saut = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                        string snaut = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                        string sidtr = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                        string sidtv = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);

                        // Guardando en la UserTable
                        B1.Company.StartTransaction();
                        if (UTDoc.GetByKey(sCode))
                        {
                            //UPDATE
                            UTDoc.UserFields.Fields.Item("U_fechaC").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechac);
                            UTDoc.UserFields.Fields.Item("U_numDoc").Value = sCode;
                            UTDoc.UserFields.Fields.Item("U_idVend").Value = svend;
                            UTDoc.UserFields.Fields.Item("U_fechaV").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechav);
                            UTDoc.UserFields.Fields.Item("U_estado").Value = sestado;
                            UTDoc.UserFields.Fields.Item("U_comment").Value = scom;
                            UTDoc.UserFields.Fields.Item("U_vend").Value = snvend;
                            UTDoc.UserFields.Fields.Item("U_idAut").Value = saut;
                            UTDoc.UserFields.Fields.Item("U_aut").Value = snaut;
                            UTDoc.UserFields.Fields.Item("U_idTR").Value = sidtr;
                            UTDoc.UserFields.Fields.Item("U_idTV").Value = sidtv;

                            iRet = UTDoc.Update();
                            todoOk = (iRet == 0);
                        }
                        else
                        {
                            //INSERT
                            UTDoc.Code = sCode;
                            UTDoc.Name = sCode;
                            UTDoc.UserFields.Fields.Item("U_fechaC").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechac);
                            UTDoc.UserFields.Fields.Item("U_numDoc").Value = sCode;
                            UTDoc.UserFields.Fields.Item("U_idVend").Value = svend;
                            UTDoc.UserFields.Fields.Item("U_fechaV").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechav);
                            UTDoc.UserFields.Fields.Item("U_estado").Value = sestado;
                            UTDoc.UserFields.Fields.Item("U_comment").Value = scom;
                            UTDoc.UserFields.Fields.Item("U_vend").Value = snvend;
                            UTDoc.UserFields.Fields.Item("U_idAut").Value = saut;
                            UTDoc.UserFields.Fields.Item("U_aut").Value = snaut;
                            UTDoc.UserFields.Fields.Item("U_idTR").Value = sidtr;
                            UTDoc.UserFields.Fields.Item("U_idTV").Value = sidtv;

                            iRet = UTDoc.Add();
                            todoOk = (iRet == 0);
                        }
                    
                    
                        ////Guardando con instrucciones SQL
                        // //Buscar si existe ese codigo para update

                        //Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        //string SQLQuery = String.Format("SELECT {0} FROM {1} WHERE {0} = '{2}'",
                        //                Constantes.View.CAB_RVT.U_numOC,
                        //                Constantes.View.CAB_RVT.CAB_RV,
                        //                sCode); 

                        //oRecordSet.DoQuery(SQLQuery);

                        //oRecordSet.MoveFirst();

                        //if (!oRecordSet.EoF)
                        //{
                        //    // UPDATE
                        //    SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}'  WHERE {0} = '{3}' ",
                        //                     Constantes.View.CAB_RVT.U_numOC,
                        //                     Constantes.View.CAB_RVT.CAB_RV,
                        //                     Constantes.View.CAB_RVT.U_comment,
                        //                     sCode, scom);
                        //    oRecordSet.DoQuery(SQLQuery);
                        //}
                        //else
                        //{
                        //    // INSERT


                        //    DateTime fc = DateTime.Now.Date;
                        //    DateTime fv = fc.AddDays(10);
                        //    SQLQuery = String.Format("INSERT INTO {0} ({7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}) "+
                        //    " VALUES('{1}','{2}','{3}','{4}','{5}','{6}','','','','{1}','{1}',' {20}' ) ",
                        //                     Constantes.View.CAB_RVT.CAB_RV,
                        //                     sCode, 
                        //                     svend, 
                        //                     fc.ToString("yyyyMMdd"),
                        //                     fv.ToString("yyyyMMdd"),
                        //                     sestado,
                        //                     scom,
                        //                     Constantes.View.CAB_RVT.U_numOC,
                        //                     Constantes.View.CAB_RVT.U_idVend,
                        //                     Constantes.View.CAB_RVT.U_fechaC,
                        //                     Constantes.View.CAB_RVT.U_fechaV,
                        //                     Constantes.View.CAB_RVT.U_estado,
                        //                     Constantes.View.CAB_RVT.U_comment,
                        //                     Constantes.View.CAB_RVT.U_idAut,
                        //                     Constantes.View.CAB_RVT.U_idTR,
                        //                     Constantes.View.CAB_RVT.U_idTV,
                        //                     Constantes.View.CAB_RVT.Code,
                        //                     Constantes.View.CAB_RVT.Name,
                        //                     Constantes.View.CAB_RVT.U_vend,
                        //                     Constantes.View.CAB_RVT.U_aut,
                        //                     snvend,snaut);
                        //    oRecordSet.DoQuery(SQLQuery);
                        //}
                }
                catch (Exception ex)
                {
                        if (B1.Company.InTransaction)
                        {
                            B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        serror = ex.Message;
                        todoOk = false;
                }
                finally
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    UTDoc = null;
                }

                //Salvando lineas del documento
                if (SMatrix != null && todoOk)
                {
                    int norecord2 = obtener_ultimo_ID("DE") ;
  
                    SMatrix.FlushToDataSource();
                    for(int i=0; i <= oDbLinesDataSource.Size-1; i++)
                    {

                        // Obteniendo texto de los campos de DbDataSource
                        string sCodeL = oDbLinesDataSource.GetValue("Code", i);
                        string sNameL = oDbLinesDataSource.GetValue("Name" ,i);
                        string scodart = oDbLinesDataSource.GetValue("U_codArt",i);
                        string sart = oDbLinesDataSource.GetValue("U_articulo",i);
                        string scodcli = oDbLinesDataSource.GetValue("U_codCli",i);
                        string sccli = oDbLinesDataSource.GetValue("U_cliente",i);
                        string scant = oDbLinesDataSource.GetValue("U_cant",i);
                        string sdisp = oDbLinesDataSource.GetValue("U_onHand", i);

                        if (scodart != "" && scodcli!= "" && scant!="")
                        {
                            try
                            {
                                // Guardando en la UserTable
                                B1.Company.StartTransaction();
                                if (UTLines.GetByKey(sCodeL))
                                {
                                    //UPDATE
                                    UTLines.UserFields.Fields.Item("U_codArt").Value = scodart;
                                    UTLines.UserFields.Fields.Item("U_articulo").Value = sart;
                                    UTLines.UserFields.Fields.Item("U_codCli").Value = scodcli;
                                    UTLines.UserFields.Fields.Item("U_cliente").Value = sccli;
                                    UTLines.UserFields.Fields.Item("U_cant").Value = Double.Parse(scant) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_onHand").Value = Double.Parse(sdisp) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_numOC").Value = sCode;
                                    iRet = UTLines.Update();
                                    todoOk = (iRet == 0);
                                }
                                else
                                {
                                    //INSERT
                                    norecord2 = norecord2 + 1;
                                    sCodeL = norecord2.ToString();
                                    UTLines.Code = sCodeL;
                                    UTLines.Name = sCodeL;
                                    UTLines.UserFields.Fields.Item("U_codArt").Value = scodart;
                                    UTLines.UserFields.Fields.Item("U_articulo").Value = sart;
                                    UTLines.UserFields.Fields.Item("U_codCli").Value = scodcli;
                                    UTLines.UserFields.Fields.Item("U_cliente").Value = sccli;
                                    UTLines.UserFields.Fields.Item("U_cant").Value = Double.Parse(scant)/1000000.00;
                                    UTLines.UserFields.Fields.Item("U_onHand").Value = Double.Parse(sdisp) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_numOC").Value = sCode;
                                    iRet = UTLines.Add();
                                    todoOk = (iRet == 0);
                                }
                            }
                            catch (Exception ex)
                            {
                                if (B1.Company.InTransaction)
                                {
                                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                serror = ex.Message;
                                todoOk = false;
                            }
                            finally
                            {
                                if (todoOk) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);}
                                
                            }
                        }

  
                    }
                    UTLines = null;
                    //oDbLineDataSource.Query();
                   // SMatrix.LoadFromDataSource();

                }
                else {todoOk = false;}
            }
            catch (Exception ex)
            {
                todoOk = false;
                serror = ex.Message;
                throw;
            }
            finally {
                //SForm.Freeze(false);
                System.GC.Collect();
            }

            if (todoOk)
            {
                todoOk = eliminar_filas_borradas();
            }
 

            if (todoOk){
               B1.Application.SetStatusBarMessage("Solicitud procesada con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
               SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
               SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)SForm.Items.Item("1").Specific;
               btn_crear.Caption = "OK";
            }
            else {
                B1.Application.SetStatusBarMessage("Error guardando solicitud: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
                

            return todoOk;

        }

        private bool eliminar_filas_borradas()
        {

            bool todoOk = true;
            string SQLQuery = String.Empty;
            try
            {
                SMatrix.LoadFromDataSource();
                if (lineasdel !=null)
                {
                    for (int i = 0; i < lineasdel.Count ; i++)
                    {

                        SQLQuery = String.Format("DELETE FROM {1} WHERE {0} = '{2}' ",
                                        Constantes.View.DET_RVT.Code,
                                        Constantes.View.DET_RVT.DET_RV,
                                        lineasdel[i]);
                        Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsCards.DoQuery(SQLQuery);
                    }
                }

            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error sincronizando eliminados " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                todoOk = false;
                throw;
            }

            finally
            {
                lineasdel.Clear();
                System.GC.Collect();
            }

            return todoOk;
        }

        private bool cargar_lineas(string noDoc)
        {
           bool todoOk = true;

           string serror = "";
           if (noDoc != "")
           {
               try
               {
                   SForm.Freeze(true);
                   // --------------------------con conditions en la dbdatasource


                       // FILTRAR LAS LINES DE SOLICITUD ACTUAL
                       SAPbouiCOM.Conditions olCons = new SAPbouiCOM.Conditions();
                       SAPbouiCOM.Condition olCon = olCons.Add();
                       olCon.Alias = "U_numOC";
                       olCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                       olCon.CondVal = noDoc;
                       oDbLinesDataSource.Query(olCons);
                       SMatrix.LoadFromDataSource();




                   // --------------------------con SQL Select
                   //String strSQL = String.Format("SELECT {0}, {1}, {2},{3},{4}, {5}, {6}, {7}, {8}, {10}" +
                   //                                    " FROM {9} Where {7}='{11}'",
                   //                                Constantes.View.DET_RVT.U_codArt, //0
                   //                                Constantes.View.DET_RVT.U_articulo, //1
                   //                                Constantes.View.DET_RVT.U_codCli, //2
                   //                                Constantes.View.DET_RVT.U_cliente, //3
                   //                                Constantes.View.DET_RVT.U_cant, //4
                   //                                Constantes.View.DET_RVT.U_estado, //5
                   //                                Constantes.View.DET_RVT.U_idTV,//6
                   //                                Constantes.View.DET_RVT.U_numOC,//7
                   //                                Constantes.View.DET_RVT.Code,//8
                   //                                Constantes.View.DET_RVT.DET_RV,//9
                   //                                Constantes.View.DET_RVT.U_onHand,//10
                   //                                noDoc);  //11

                   //Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                   //rsCards.DoQuery(strSQL);
                   //SMatrix.Clear();
                   //SAPbobsCOM.Fields fields = rsCards.Fields;
                   //rsCards.MoveFirst();
                   //for (int i = 1; !rsCards.EoF; i++)
                   //{
                   //    SMatrix.AddRow(1, 1);
                   //    SMatrix.Columns.Item(1).Cells.Item(i).Specific.Value = fields.Item("U_codArt").Value.ToString();
                   //    SMatrix.Columns.Item(2).Cells.Item(i).Specific.Value = fields.Item("U_articulo").Value.ToString();
                   //    SMatrix.Columns.Item(3).Cells.Item(i).Specific.Value = fields.Item("U_codCli").Value.ToString();
                   //    SMatrix.Columns.Item(4).Cells.Item(i).Specific.Value = fields.Item("U_cliente").Value.ToString();
                   //    SMatrix.Columns.Item(5).Cells.Item(i).Specific.Value = fields.Item("U_cant").Value.ToString();
                   //    SMatrix.Columns.Item(6).Cells.Item(i).Specific.Value = fields.Item("U_onHand").Value.ToString();
                   //    SMatrix.Columns.Item(7).Cells.Item(i).Specific.Checked = fields.Item("U_estado").Value.ToString()=="A";
                   //    SMatrix.Columns.Item(8).Cells.Item(i).Specific.Value = fields.Item("U_idTV").Value.ToString();
                   //    SMatrix.Columns.Item(9).Cells.Item(i).Specific.Value = fields.Item("U_numOC").Value.ToString();
                   //    SMatrix.Columns.Item(10).Cells.Item(i).Specific.Value = fields.Item("code").Value.ToString();                       
                        
                   //    rsCards.MoveNext();
                   //}
                  SMatrix.AutoResizeColumns();
                  // SMatrix.FlushToDataSource();


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
                   B1.Application.SetStatusBarMessage("Error cargando lineas: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
               }

               SForm.Freeze(false);
               return todoOk;
           }
           else { return true; }
        }

        private bool cargar_solicitud(string noDoc, bool posicion)
        {

            bool todoOk = true;

            string serror = "";


            if (oDbHeaderDataSource.Size == 0)
            {
                return insertar_solicitud();
            }
            else
            {
                if (noDoc != "")
                {



                    SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                    SAPbouiCOM.EditText txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                    SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                    SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                    SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                    SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                    SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                    SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                    SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                    SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item("mtx").Specific;
                    SAPbouiCOM.ComboBox oCombox = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                    SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_aut).Specific;
 

                    try
                    {




                        if (B1.Company.InTransaction || SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || SForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            int respuesta = B1.Application.MessageBox("Desea cancelar los datos modificados? ", 1, "OK", " Cancelar");
                            if (respuesta == 1)
                            {
                                if (B1.Company.InTransaction)
                                {
                                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                }


                                if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    oDbHeaderDataSource.RemoveRecord(oDbHeaderDataSource.Size - 1);
                                }
                                todoOk = true;
                            }
                            else { todoOk = false; }
                        }
                        if (todoOk)
                        {
                            int nuevaposic = 0;
                            if (!posicion)
                            {
                                //Buscando posicion fisica sino es invocado dese el Find
                                //string SQLQuery = String.Format("SELECT {0} FROM {1}",
                                //                    Constantes.View.CAB_RVT.U_numOC,
                                //                    Constantes.View.CAB_RVT.CAB_RV);

                                //Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                //oRecordSet.DoQuery(SQLQuery);
                                //oRecordSet.MoveFirst();
                                //bool encontrado = false;
                                //int i;
                                //for (i = 0; !oRecordSet.EoF && !encontrado; i++)
                                //{
                                //    encontrado = oRecordSet.Fields.Item("U_numDoc").Value.ToString() == noDoc;
                                //    oRecordSet.MoveNext();
                                //}

                                //if (encontrado)
                                //{
                                //    nuevaposic = i - 1;
                                //}
                                nuevaposic = Int32.Parse(noDoc);
                            }
                            else
                            {
                                nuevaposic = Int32.Parse(noDoc) - 1; //Viene increm del Find 
                            }

                            nuevaposic = nuevaposic < 0 ? 0 : nuevaposic;
                            oDbHeaderDataSource.Offset = nuevaposic;

                            oDbHeaderDataSource.Query();
 
                            txt_numoc.Value = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                            txt_idvend.Value = oDbHeaderDataSource.GetValue("U_idVend", oDbHeaderDataSource.Offset);
                            txt_idaut.Value = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                            txt_vend.Value = oDbHeaderDataSource.GetValue("U_vend", oDbHeaderDataSource.Offset);
                            //txt_vend.Value = txt_vend.Value.ToString() == "" ? obtener_NameVendedor():txt_vend.Value;
                            oDbHeaderDataSource.SetValue("U_vend", oDbHeaderDataSource.Offset, txt_vend.Value.ToString());
                            txt_aut.Value = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                            txt_idtv.Value = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);
                            txt_idtr.Value = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                            txt_fechac.Value = oDbHeaderDataSource.GetValue("U_fechaC", oDbHeaderDataSource.Offset);
                            txt_fechav.Value = oDbHeaderDataSource.GetValue("U_fechaV", oDbHeaderDataSource.Offset);
                            txt_com.Value = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                            txt_estado.Value = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                            txt_estado.Value = txt_estado.Value == "N" ? " Nueva" : "Revisada";

                        }
                    }
                    catch (Exception ex)
                    {
                        todoOk = false;
                        serror = ex.Message;
                        throw;
                    }

                    if (todoOk)
                    {
                        B1.Application.SetStatusBarMessage("Solicitud cargada con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        todoOk = cargar_lineas(txt_numoc.Value.ToString());
                        if (todoOk)
                        {
                            SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            btn_crear.Caption = "OK";
                            txt_com.Item.Enabled = true;
                            txt_com.Active = true;
                            mtx.Item.Enabled = true;
                            oCombox.Item.Visible = false;
                        }
                        else
                        {
                            B1.Application.SetStatusBarMessage("Error cargando Solicitud ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        }

                    }
                    else
                    {
                        B1.Application.SetStatusBarMessage("Error cargando lineas: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    }

                    //SForm.Freeze(false);
                    return todoOk;
                }
                else { return true; }
            }
        }

        private double obtener_exist_articulo(string codart)
        {
            double exist = 0.00;
            try
            {
                //String strSQL = String.Format("SELECT T0.{0},T0.{2},T1.{3} FROM {4} T0 INNER JOIN {5} T1"  +
                //    " ON T0.{2} = T1.{2} " +
                //    " WHERE contains(T0.{1},'%{6}%') AND T0.{2}='CD'",
                //          Constantes.View.oitw.OnHand,     
                //          Constantes.View.oitw.ItemCode,
                //          Constantes.View.owhs.WhsCode,
                //          Constantes.View.owhs.WhsName,
                //          Constantes.View.oitw.OITW,
                //          Constantes.View.owhs.OWHS,
                //          codart);
                String strSQL = String.Format("SELECT {0} FROM {3} " +
                    " WHERE contains({1},'%{4}%') AND {2}='{5}'  ",
                          Constantes.View.oitw.OnHand,
                          Constantes.View.oitw.ItemCode,
                          Constantes.View.oitw.WhsCode,
                          Constantes.View.oitw.OITW,
                          codart,
                          "CD");
                Recordset rsResult = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsResult.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsResult.Fields;
                rsResult.MoveFirst();
                if (!rsResult.EoF)
                { 
                    exist = Double.Parse(rsResult.Fields.Item("OnHand").Value.ToString());
                    //string wc = rsResult.Fields.Item("WhsCode").Value.ToString();
                    //string wn = rsResult.Fields.Item("WhsName").Value.ToString();
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Stock Disponible" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
            return exist;
        }

        private bool validar_art_cliente_unicos( string art, string cli, int row)
        {
            bool todoOK = true;
            if(SMatrix.RowCount > 1)
            {
                try
                {
                    // Validar contra la misma matriz porque cuando es nuevo solo datos en linea, 
                    // No fisicos en la BD
                    int creg = 0;
                    for (int i = 1; i <= SMatrix.RowCount && creg < 1; i++)
                    {
                        if ((i != row) &&
                            (SMatrix.Columns.Item(1).Cells.Item(i).Specific).Value.ToString() == art &&
                            (SMatrix.Columns.Item(3).Cells.Item(i).Specific).Value.ToString() == cli)
                        {
                            creg++;
                        }
                    }
                    todoOK = (creg < 1);

                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error validando campos repetidos" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    todoOK = false;
                    throw;
                }
            }

            return todoOK;
        }

        private int obtener_ultimo_ID(string tipo)
        {
            int CodeNumCA = 0;
            int CodeNumDE = 0;
            if (tipo == "CA")
            {

                String strSQL = String.Format("SELECT TOP 1 CAST(T0.{0} AS INT) AS nd FROM {1} T0 ORDER BY CAST(T0.{0} AS INT) DESC",
                                    Constantes.View.CAB_RVT.U_numOC,
                                    Constantes.View.CAB_RVT.CAB_RV);

                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);

                string Code = rsCards.Fields.Item("nd").Value.ToString();

                //probar cuando la tabla este vacia, osea el primero registro y no haya otro anterior
                if (Code != "")
                {
                    CodeNumCA = Convert.ToInt32(Code);

                }
                return CodeNumCA;
            }
            else
            {

                String strSQL = String.Format("SELECT TOP 1 CAST(T0.{0} AS INT) AS nl FROM {1} T0 ORDER BY CAST(T0.{0} AS INT) DESC",
                                    Constantes.View.DET_RVT.Code,
                                    Constantes.View.DET_RVT.DET_RV);

                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);

                string Code = rsCards.Fields.Item("nl").Value.ToString();
                if (Code != "")
                {
                    CodeNumDE = Convert.ToInt32(Code);

                }

                return CodeNumDE;
            }




        }

        public void llenar_combo_id(SAPbouiCOM.ComboBox oCombo, string SqlQuery)
        {
            SAPbobsCOM.Recordset oRecordSet = null;

            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(SqlQuery);

            oCombo.ValidValues.Add("0", "Seleccione Documento:");

            for (int i = 1; !oRecordSet.EoF;i++ )
            {
                oCombo.ValidValues.Add(i.ToString(), oRecordSet.Fields.Item(0).Value.ToString() + " ("+ oRecordSet.Fields.Item(1).Value.ToString("dd/MM/yyyy")+")" );
                oRecordSet.MoveNext();

            }


            //oCombo.Select("0", (SAPbouiCOM.BoSearchKey)(0));
            oCombo.Item.DisplayDesc = false;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);

        }

        private string obtener_NameVendedor()
        {
            try
            {
                string usrCurrent = B1.Company.UserName;
                String strSQL = String.Format("SELECT {0}, {1}   FROM {2} Where contains({0},'%{3}%')",
                          Constantes.View.ousr.uCode,
                          Constantes.View.ousr.uName,
                          Constantes.View.ousr.OUSR,
                          usrCurrent);
                Recordset rsUsers = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsUsers.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsUsers.Fields;
                rsUsers.MoveFirst();
                string User_Name = rsUsers.Fields.Item("U_NAME").Value.ToString();
                return User_Name;
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Vendedor", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }

        private string obtener_Vendedor()
        {
            try
            {
                string usrCurrent = B1.Company.UserName;
                String strSQL = String.Format("SELECT {0},{1}  FROM {2} Where contains({0},'%{3}%')",
                          Constantes.View.ousr.uCode,
                          Constantes.View.ousr.uName,
                          Constantes.View.ousr.OUSR,
                          usrCurrent);
                Recordset rsUsers = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsUsers.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsUsers.Fields;

                string User_Code = rsUsers.Fields.Item("USER_CODE").Value.ToString();
                string User_Name = rsUsers.Fields.Item("U_NAME").Value.ToString();
                return User_Code;
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Vendedor", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }

        
    }
}
