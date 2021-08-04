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
        private SAPbouiCOM.Form SForm = null;
        private SAPbouiCOM.Matrix SMatrix = null;

        private SAPbouiCOM.DBDataSource oDbLinesDataSource = null;
        private SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;

        private bool registrar = true;
        private List<string> lineasdel = new List<string>();
        private int rowsel = 0;
        private int indice = 0;
        private string docaprob = "";

        private SAPbouiCOM.ComboBox oCombo = null;
        private SAPbouiCOM.EditText txt_numoc = null;
        private SAPbouiCOM.EditText txt_fechac = null;
        private SAPbouiCOM.EditText txt_fechav = null;
        private SAPbouiCOM.EditText txt_estado = null;
        private SAPbouiCOM.EditText txt_idvend = null;
        private SAPbouiCOM.EditText txt_vend = null;
        private SAPbouiCOM.EditText txt_idaut = null;
        private SAPbouiCOM.EditText txt_aut = null;
        private SAPbouiCOM.EditText txt_idtv = null;
        private SAPbouiCOM.EditText txt_idtr = null;
        private SAPbouiCOM.EditText txt_com = null;
        private SAPbouiCOM.Matrix mtx = null;
        private SAPbouiCOM.Button btn_crear = null;
        private SAPbouiCOM.Button btn_cancel = null;
        private SAPbouiCOM.Button btn_autorizar = null;
        private SAPbouiCOM.Button btn_tr = null;
        private SAPbouiCOM.Button btn_cancelar = null;
        private SAPbouiCOM.Button btn_tv = null;


       

        public PantallaRegistro(bool registro = true, string doc="" )
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Registro.srf"], "SolRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            formActual = "SolRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);
            registrar = registro;
            docaprob = doc;  
            cargar_inicial();
        }

       

        // Metodos Override

        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
              BubbleEvent = true;
              if (B1.Application.Forms.ActiveForm.UniqueID == formActual && registrar)
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
                        //case "1281":    // Buscar                      
                        //    preparar_modo_Find();
                        //    BubbleEvent = false;
                        //    break;
                        case "1283":    // Eliminar                     
                            eliminar_solicitud();
                            BubbleEvent = false;
                            break;
                        case "1292":   //Adicionar linea
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.registro.mtx:
                                    insertar_linea_solic();
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                        case "1293":  //Borrar linea
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.registro.mtx:
                                    borrar_linea_solic();
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
                if (eventInfo.FormUID == formActual && registrar)
                {
                    ItemActiveMenu = eventInfo.ItemUID;
                    if (eventInfo.BeforeAction && eventInfo.ItemUID == ventaRT.Constantes.View.registro.mtx)
                    {
                        SForm.EnableMenu("1292", true); //Activar Agregar Linea
                        SForm.EnableMenu("1293", true); //Activar Borrar Linea 
                        rowsel = eventInfo.Row;
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
                B1.Application.SetStatusBarMessage("Error Activando Menu" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                            case BoEventTypes.et_COMBO_SELECT:
                                {
                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.registro.cbnd:
                                            {
                                                SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                                btn_crear.Caption = "OK";
                                                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                                string buscado = oCombo.Selected.Value == null ? " 0" : oCombo.Selected.Value.ToString();
                                                if (buscado != "0")
                                                {
                                                    indice = Int32.Parse(buscado);
                                                    cargar_solicitud(buscado, true);
                                                }
                                                BubbleEvent = false;
                                                break;
                                            }
                                    }
                                }
                                break;

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
                                                            oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                                            string buscado = oCombo.Selected.Value == null ? " 0" : oCombo.Selected.Value.ToString();
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

                                        case Constantes.View.registro.btn_autorizar:
                                            {
                                                aprobar_solicitud();
                                                break;
                                            }

                                        case Constantes.View.registro.btn_TR:
                                            {
                                                transferir_aceptados();
                                                break;
                                            }

                                    }
                                    break;
                                }

                            case BoEventTypes.et_CHOOSE_FROM_LIST:
                                {
                                    if (pVal.InnerEvent == true)
                                    {
                                        SAPbouiCOM.ChooseFromList oCFL;
                                        SAPbouiCOM.IChooseFromListEvent CFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;
                                        string CFL_Id = CFLEvent.ChooseFromListUID;
                                        oCFL = SForm.ChooseFromLists.Item(CFL_Id);
                                        if (pVal.FormTypeEx.Substring(0, 5) == "SolRT" && CFLEvent.SelectedObjects != null)
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
                                        string SQLQuery = string.Empty;
                                        string usrCurrent = B1.Company.UserName;
                                        SQLQuery = String.Format("SELECT CAST(T0.{0} AS INT) AS ND, {2} FROM {1} T0 " +
                                        " WHERE {3} = '{4}'  ORDER BY CAST(T0.{0} AS INT) ASC",
                                                                            Constantes.View.CAB_RVT.U_numOC,
                                                                            Constantes.View.CAB_RVT.CAB_RV,
                                                                            Constantes.View.CAB_RVT.U_fechaC,
                                                                            Constantes.View.CAB_RVT.U_idVend,
                                                                            usrCurrent);
                                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
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

        private void cargar_inicial()
        {

            SForm = B1.Application.Forms.ActiveForm;
            SMatrix = SForm.Items.Item("mtx").Specific;
            formActual = B1.Application.Forms.ActiveForm.UniqueID;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
            oDbLinesDataSource = SForm.DataSources.DBDataSources.Item("@DET_RSTV");


            SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            SAPbouiCOM.Button btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;
            SAPbouiCOM.Button btn_autorizar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_autorizar).Specific;
            SAPbouiCOM.Button btn_tr = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TR).Specific;
            SAPbouiCOM.Button btn_cancelar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancelar).Specific;
            SAPbouiCOM.Button btn_tv = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TV).Specific;
            oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
            txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
            txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;




            if (registrar)
            {
                // Vendedor
                oCombo.Active = true;
                btn_autorizar.Item.Visible = false;
                btn_tr.Item.Visible = false;
                btn_cancelar.Item.Visible = false;
                btn_tv.Item.Visible = false;

                SForm.EnableMenu("1290", true); SForm.EnableMenu("1289", true);
                SForm.EnableMenu("1288", true); SForm.EnableMenu("1291", true);
                SForm.EnableMenu("1282", true); SForm.EnableMenu("1283", true);
                //SForm.EnableMenu("1281", true); 
                SForm.EnableMenu("1281", false);  //buscar

            }
            else
            {
                // Autorizador

                btn_cancel.Item.Visible = false;

                SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                mtx.Item.Enabled = true;
                // Desactivar todas las columnas de la matriz menos el estado
                SAPbouiCOM.Column oColumn = SMatrix.Columns.Item("codArt");
                oColumn.Editable = false;
                oColumn = SMatrix.Columns.Item("codCli");
                oColumn.Editable = false;
                oColumn = SMatrix.Columns.Item("cant");
                oColumn.Editable = false;

                txt_com.Item.Enabled = true;
                txt_com.Active = true;
                oCombo.Item.Visible = false;

                SForm.EnableMenu("1290", false); SForm.EnableMenu("1289", false);
                SForm.EnableMenu("1288", false); SForm.EnableMenu("1291", false);
                SForm.EnableMenu("1282", false); SForm.EnableMenu("1283", false);
                //SForm.EnableMenu("1281", true); 
                SForm.EnableMenu("1281", false);  //buscar

                
            }


            SForm.Freeze(true);


            if (registrar)
            {
                // Vendedor

                oCombo.Item.Click(BoCellClickType.ct_Regular);
            }
            else
            {
                // Autorizador
  
                // Filtrar documento a aprobar
                SAPbouiCOM.Conditions orCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition orCon = orCons.Add();
                orCon.Alias = "U_numDoc";
                orCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                orCon.CondVal = docaprob;
                oDbHeaderDataSource.Query(orCons);

                // Filtrar lineas del documento
                SAPbouiCOM.Conditions olCons2 = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition olCon2 = olCons2.Add();
                olCon2.Alias = "U_numOC";
                olCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                olCon2.CondVal = docaprob;
                oDbLinesDataSource.Query(olCons2);
                SMatrix.LoadFromDataSource();
                // Recargar DocNum de Transferencia o Devolucion
                string dentry = "";
                for (int i = 1; i <= SMatrix.RowCount; i++)
                {
                    dentry = (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value.ToString();
                    (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value = obtener_DocNum(dentry);
                }
                SMatrix.AutoResizeColumns();


                SMatrix.AutoResizeColumns();
                SAPbouiCOM.Column oColumn = SMatrix.Columns.Item("codArt");
                oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);

                txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                txt_estado.Value = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                txt_estado.Value = txt_estado.Value == "" ? "N" : txt_estado.Value;

                txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                txt_idtr.Value = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                txt_idtr.Value = obtener_DocNum(txt_idtr.Value);

                string estadoactual = txt_estado.Value.ToString().Substring(0, 1) ;
                txt_estado.Value = obtener_Estado(estadoactual);
                btn_autorizar.Item.Enabled = estadoactual == "N";
                btn_tr.Item.Enabled = estadoactual == "A";
                btn_cancelar.Item.Enabled = estadoactual == "T";
                btn_tv.Item.Enabled = estadoactual == "C";

                oColumn = SMatrix.Columns.Item("estado");
                oColumn.Editable = (estadoactual == "A" || estadoactual == "C");
            }
            SForm.Freeze(false);
        }

        private bool insertar_solicitud()
        {
            bool todoOk = true;
            try {
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


                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                    txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                    txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                    txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                    txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                    txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                    txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                    txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                    SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                    SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                    txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                    mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                    SAPbouiCOM.Button btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;
                    SAPbouiCOM.Button btn_autorizar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_autorizar).Specific;
                    SAPbouiCOM.Button btn_tr = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TR).Specific;
                    SAPbouiCOM.Button btn_cancelar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancelar).Specific;
                    SAPbouiCOM.Button btn_tv = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TV).Specific;


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



            oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
            txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
            txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
            txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
            txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
            txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
            txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
            txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
            SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
            SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
            txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
            SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            SAPbouiCOM.Button btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;
            SAPbouiCOM.Button btn_autorizar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_autorizar).Specific;
            SAPbouiCOM.Button btn_tr = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TR).Specific;
            SAPbouiCOM.Button btn_cancelar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancelar).Specific;
            SAPbouiCOM.Button btn_tv = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TV).Specific;


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
                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        btn_crear.Caption = "Buscar";
                        
                        // inhabilitando todo
                        txt_idvend.Value = "";
                        txt_vend.Value = "";
                        txt_numoc.Value = "";
                        txt_fechac.Value = "";
                        txt_fechav.Value = "";
                        txt_estado.Value = "";
                        txt_com.Value = "";
                        txt_idaut.Value = "";
                        txt_aut.Value = "";
                        SMatrix.Item.Enabled = false;
                        txt_com.Item.Enabled = false;
                        oCombo.Active = true;
                        //oCombo.Item.Click(BoCellClickType.ct_Regular);
                        //oCombo.Item.Enabled = true;
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
                    indice = 1;
                    //cargar_solicitud(indice.ToString(), false);
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    oCombo.Select(indice.ToString(), BoSearchKey.psk_ByValue);
                    //cargar_solicitud("0", false);
                    B1.Application.SetStatusBarMessage("Movimiento al Primero ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error en Movimiento al Primero ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    throw ex;
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
                    if (indice > 0)
                    {
                        //oDbHeaderDataSource.Offset--;
                        //cargar_solicitud(oDbHeaderDataSource.Offset.ToString(), false);
                        indice--;
                        //cargar_solicitud(indice.ToString(), false);
                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                        oCombo.Select(indice.ToString(), BoSearchKey.psk_ByValue);
                    }
                    B1.Application.SetStatusBarMessage("Movimiento al Anterior ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error en Movimiento al Anterior ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                    //oDbHeaderDataSource.Offset++;
                    //cargar_solicitud(oDbHeaderDataSource.Offset.ToString(), false);
                    if (indice < oDbHeaderDataSource.Size)
                    {
                        indice++;
                        //cargar_solicitud(indice.ToString(), false);
                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                        oCombo.Select(indice.ToString(), BoSearchKey.psk_ByValue);
                        B1.Application.SetStatusBarMessage("Movimiento al Siguiente ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    }
                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error en Movimiento al Siguiente ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    throw ex;
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
                    //oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size - 1;
                    //cargar_solicitud(oDbHeaderDataSource.Offset.ToString(), false);
                    //indice = oDbHeaderDataSource.Size - 1;
                    //cargar_solicitud(indice.ToString(), false);
                    indice = oDbHeaderDataSource.Size;
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    oCombo.Select(indice.ToString(), BoSearchKey.psk_ByValue);
                    B1.Application.SetStatusBarMessage("Movimiento al Ultimo ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error en Movimiento al Ultimo ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    throw ex;
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
                    string abuscar = txt_numoc.Value.ToString();

                    if (ya_Procesada(abuscar))
                    {
                        todoOk = false;
                        serror = "Solicitud Procesada, no se puede eliminar..";
                    }
                    else
                    {
                        string SQLQuery = String.Format("SELECT {0}, {2} FROM {1}",
                                              Constantes.View.CAB_RVT.U_numOC,
                                              Constantes.View.CAB_RVT.CAB_RV,
                                              Constantes.View.CAB_RVT.U_estado);

                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery(SQLQuery);
                        oRecordSet.MoveFirst();
                        bool encontrado = false;
                        string estadoactual = "";
                        int i;
                        for (i = 0; !oRecordSet.EoF && !encontrado; i++)
                        {
                            encontrado = oRecordSet.Fields.Item("U_numDoc").Value.ToString() == abuscar;
                            estadoactual = oRecordSet.Fields.Item("U_estado").Value.ToString();
                            oRecordSet.MoveNext();
                        }

                        if (encontrado)
                        {
                            // Validar que sea Nueva sino no se puede borrar

                            if (estadoactual != "N")
                            {
                                todoOk = false;
                                serror = "Documento en Proceso, no se puede Eliminar";
                            }
                            else
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

                        }
                        else
                        {
                            todoOk = false;
                            serror = "Documento No Encontrado";
                        }
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
                    B1.Application.SetStatusBarMessage("Solicitud eliminada con exito", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    B1.Application.SetStatusBarMessage("Error eliminando solicitud: " + serror, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                        int norecord =  Int32.Parse(txt_numoc.Value.ToString());
                        sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                        string sfechav = oDbHeaderDataSource.GetValue("U_fechaV",oDbHeaderDataSource.Offset);
                        string sfechac = oDbHeaderDataSource.GetValue("U_fechaC", oDbHeaderDataSource.Offset);
                        string sestado = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                        sestado = (sestado=="" ?"N" :sestado).Substring(0,1);
                        //sestado = "N";
                        string scom = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                        string svend = oDbHeaderDataSource.GetValue("U_idVend", oDbHeaderDataSource.Offset);
                        string snvend = oDbHeaderDataSource.GetValue("U_vend", oDbHeaderDataSource.Offset);
                        string saut = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                        string snaut = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                        string sidtr = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                        string sidtv = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);

                        //// Guardando en la UserTable
                        //B1.Company.StartTransaction();
                        //if (UTDoc.GetByKey(sCode))
                        //{
                        //    //UPDATE
                        //    UTDoc.UserFields.Fields.Item("U_fechaC").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechac);
                        //    UTDoc.UserFields.Fields.Item("U_numDoc").Value = sCode;
                        //    UTDoc.UserFields.Fields.Item("U_idVend").Value = svend;
                        //    UTDoc.UserFields.Fields.Item("U_fechaV").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechav);
                        //    UTDoc.UserFields.Fields.Item("U_estado").Value = sestado;
                        //    UTDoc.UserFields.Fields.Item("U_comment").Value = scom;
                        //    UTDoc.UserFields.Fields.Item("U_vend").Value = snvend;
                        //    UTDoc.UserFields.Fields.Item("U_idAut").Value = saut;
                        //    UTDoc.UserFields.Fields.Item("U_aut").Value = snaut;
                        //    UTDoc.UserFields.Fields.Item("U_idTR").Value = sidtr;
                        //    UTDoc.UserFields.Fields.Item("U_idTV").Value = sidtv;

                        //    iRet = UTDoc.Update();
                        //    todoOk = (iRet == 0);
                        //}
                        //else
                        //{
                        //    //INSERT
                        //    UTDoc.Code = sCode;
                        //    UTDoc.Name = sCode;
                        //    UTDoc.UserFields.Fields.Item("U_fechaC").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechac);
                        //    UTDoc.UserFields.Fields.Item("U_numDoc").Value = sCode;
                        //    UTDoc.UserFields.Fields.Item("U_idVend").Value = svend;
                        //    UTDoc.UserFields.Fields.Item("U_fechaV").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(sfechav);
                        //    UTDoc.UserFields.Fields.Item("U_estado").Value = sestado;
                        //    UTDoc.UserFields.Fields.Item("U_comment").Value = scom;
                        //    UTDoc.UserFields.Fields.Item("U_vend").Value = snvend;
                        //    UTDoc.UserFields.Fields.Item("U_idAut").Value = saut;
                        //    UTDoc.UserFields.Fields.Item("U_aut").Value = snaut;
                        //    UTDoc.UserFields.Fields.Item("U_idTR").Value = sidtr;
                        //    UTDoc.UserFields.Fields.Item("U_idTV").Value = sidtv;

                        //    iRet = UTDoc.Add();
                        //    todoOk = (iRet == 0);
                        //}


                        //Guardando con instrucciones SQL
                        //Buscar si existe ese codigo para update

                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string SQLQuery = String.Format("SELECT {0} FROM {1} WHERE {0} = '{2}'",
                                        Constantes.View.CAB_RVT.U_numOC,
                                        Constantes.View.CAB_RVT.CAB_RV,
                                        sCode);

                        oRecordSet.DoQuery(SQLQuery);

                        oRecordSet.MoveFirst();

                        if (!oRecordSet.EoF)
                        {
                            // UPDATE
                            SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}', {5}='{6}'  WHERE {0} = '{3}' ",
                                             Constantes.View.CAB_RVT.U_numOC,   //0
                                             Constantes.View.CAB_RVT.CAB_RV,    //1
                                             Constantes.View.CAB_RVT.U_comment, //2
                                             sCode,                             //3
                                             scom,                              //4
                                             Constantes.View.CAB_RVT.U_estado,  //5
                                             sestado);                          //6
   
                            oRecordSet.DoQuery(SQLQuery);
                        }
                        else
                        {
                            // INSERT

                            DateTime fc = DateTime.Now.Date;
                            DateTime fv = fc.AddDays(10);
                            SQLQuery = String.Format("INSERT INTO {0} ({7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}) " +
                            " VALUES('{1}','{2}','{3}','{4}','{5}','{6}','{21}','{23}','{24}','{1}','{1}','{20}', '{21}' ) ",
                                             Constantes.View.CAB_RVT.CAB_RV,    //0
                                             sCode, //1
                                             svend, //2
                                             fc.ToString("yyyyMMdd"), //3
                                             fv.ToString("yyyyMMdd"), //4
                                             sestado, //5
                                             scom, //6
                                             Constantes.View.CAB_RVT.U_numOC, //7
                                             Constantes.View.CAB_RVT.U_idVend, //8
                                             Constantes.View.CAB_RVT.U_fechaC, //9
                                             Constantes.View.CAB_RVT.U_fechaV, //10
                                             Constantes.View.CAB_RVT.U_estado, //11
                                             Constantes.View.CAB_RVT.U_comment, //12
                                             Constantes.View.CAB_RVT.U_idAut, //13
                                             Constantes.View.CAB_RVT.U_idTR, //14
                                             Constantes.View.CAB_RVT.U_idTV, //15
                                             Constantes.View.CAB_RVT.Code, //16
                                             Constantes.View.CAB_RVT.Name, //17
                                             Constantes.View.CAB_RVT.U_vend, //18
                                             Constantes.View.CAB_RVT.U_aut, //19
                                             snvend, //20
                                             saut, //21
                                             snaut,  //22
                                             sidtr,  //23
                                             sidtv //24
                                             );
                            oRecordSet.DoQuery(SQLQuery);
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
                    if (B1.Company.InTransaction)
                    {
                        B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
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
                        string sestad = oDbLinesDataSource.GetValue("U_estado", i);

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
                                    UTLines.UserFields.Fields.Item("U_estado").Value = sestad;
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

                    // FILTRAR LAS LINES DE SOLICITUD ACTUAL
                    SAPbouiCOM.Conditions olCons2 = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition olCon2 = olCons2.Add();
                    olCon2.Alias = "U_numOC";
                    olCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    olCon2.CondVal = noDoc;
                    oDbLinesDataSource.Query(olCons2);
                    SMatrix.LoadFromDataSource();
                    // Recargar DocNum de Transferencia o Devolucion
                    string dentry = ""; 
                    for (int i = 1; i <= SMatrix.RowCount; i++)
                    {
                        dentry = (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value.ToString();
                        (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value = obtener_DocNum(dentry);
                    }
                    SMatrix.AutoResizeColumns();
                    SAPbouiCOM.Column oColumn = SMatrix.Columns.Item("codArt");
                    oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);
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
                 SForm.Freeze(true);
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

                            oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                            txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                            txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                            txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                            txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                            txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                            txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                            txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                            txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                            SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                            SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                            txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                            mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                            txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;



                            int nuevaposic = 0;
                            if (!posicion)
                            {
                                // Navegacion normal
                                //nuevaposic = Int32.Parse(noDoc);

                                // nuevo invento
                                //Buscando posicion fisica
                                nuevaposic = Int32.Parse(noDoc)+1;
                                string SQLQuery = String.Format("SELECT TOP {2} CAST({0} AS INT) AS ND" +
                                    " FROM {1} ORDER BY CAST({0} AS INT)  ASC",
                                                    Constantes.View.CAB_RVT.U_numOC,
                                                    Constantes.View.CAB_RVT.CAB_RV, nuevaposic.ToString());
                                Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oRecordSet.DoQuery(SQLQuery);
                                oRecordSet.MoveLast();
                                string nuevodoc = "";

                                if (!oRecordSet.EoF)
                                {
                                    nuevodoc = oRecordSet.Fields.Item("ND").Value.ToString();
                                    //Buscando posicion fisica
                                    SQLQuery = String.Format("SELECT {0} FROM {1}",
                                                        Constantes.View.CAB_RVT.U_numOC,
                                                        Constantes.View.CAB_RVT.CAB_RV);
                                    oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    oRecordSet.DoQuery(SQLQuery);
                                    oRecordSet.MoveFirst();
                                    bool encontrado = false;
                                    int i;
                                    for (i = 0; !oRecordSet.EoF && !encontrado; i++)
                                    {
                                        encontrado = oRecordSet.Fields.Item("U_numDoc").Value.ToString() == nuevodoc;
                                        oRecordSet.MoveNext();
                                    }
                                    if (encontrado)
                                    {
                                        nuevaposic = i - 1;
                                    }
                                }
                            }
                            else
                            {
                                //Buscando posicion fisica
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
                                    encontrado = oRecordSet.Fields.Item("U_numDoc").Value.ToString() == noDoc;
                                    oRecordSet.MoveNext();
                                }
                                if (encontrado)
                                {
                                    nuevaposic = i - 1;
                                }
                            }


                            // Vendedor
                            // FILTRAR LAS SOLICITUDES DEL USUARIO ACTUAL
                            SAPbouiCOM.Conditions orCons = new SAPbouiCOM.Conditions();
                            SAPbouiCOM.Condition orCon = orCons.Add();
                            orCon.Alias = "U_idVend";
                            orCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            orCon.CondVal = B1.Company.UserName;

                            // Carga Inicial de Datos si no esta hecha
                            if (docaprob == "")
                            {
                                oDbHeaderDataSource.Query(orCons);
                                docaprob=oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                            }
  
                            // Carga de la Solicitud Encontrada

                            nuevaposic = nuevaposic < 0 ? 0 : nuevaposic;

                            if (oDbHeaderDataSource.Offset != nuevaposic)
                            {
                                oDbHeaderDataSource.Offset = nuevaposic;
                                oDbHeaderDataSource.Query(orCons);
                            }
                            txt_numoc.Value = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                            //txt_idvend.Value = oDbHeaderDataSource.GetValue("U_idVend", oDbHeaderDataSource.Offset);
                            //txt_idaut.Value = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                            //txt_vend.Value = oDbHeaderDataSource.GetValue("U_vend", oDbHeaderDataSource.Offset);
                            ////txt_vend.Value = txt_vend.Value.ToString() == "" ? obtener_NameVendedor():txt_vend.Value;
                            //oDbHeaderDataSource.SetValue("U_vend", oDbHeaderDataSource.Offset, txt_vend.Value.ToString());
                            //txt_aut.Value = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                            //txt_idtv.Value = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);
                            //txt_idtr.Value = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                            //txt_fechac.Value = oDbHeaderDataSource.GetValue("U_fechaC", oDbHeaderDataSource.Offset);
                            //txt_fechav.Value = oDbHeaderDataSource.GetValue("U_fechaV", oDbHeaderDataSource.Offset);
                            //txt_com.Value = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                            txt_estado.Value = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                            txt_estado.Value = obtener_Estado(txt_estado.Value);
                            txt_idtr.Value = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                            txt_idtr.Value = obtener_DocNum(txt_idtr.Value);
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
                            SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                            btn_crear.Caption = "OK";
                            //oCombox.Select(noDoc,BoSearchKey.psk_Index);
                            mtx.Item.Enabled = true;
                            txt_com.Item.Enabled = true;
                            txt_com.Active = true;

                            //oCombox.Item.Visible = false;
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

                    SForm.Freeze(false);
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
                //oCombo.ValidValues.Add(i.ToString(), oRecordSet.Fields.Item(0).Value.ToString() + " ("+ oRecordSet.Fields.Item(1).Value.ToString("dd/MM/yyyy")+")" );
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString(), oRecordSet.Fields.Item(1).Value.ToString("dd/MM/yyyy"));
                oRecordSet.MoveNext();

            }
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

        private bool ya_Procesada(string nodoc)
        {
            try
            {
                String strSQL = String.Format("SELECT {1} FROM {2}  Where {0}='{3}'",
                                    Constantes.View.CAB_RVT.U_numOC,
                                    Constantes.View.CAB_RVT.U_estado,
                                    Constantes.View.CAB_RVT.CAB_RV,
                                    nodoc);

                Recordset rsUsers = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsUsers.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsUsers.Fields;
                rsUsers.MoveFirst();
                if (rsUsers.EoF)
                {
                    return false;
                }
                else
                {
                    string estado = rsUsers.Fields.Item("U_estado").Value.ToString();
                    return estado.Substring(0,1) != "N" ;
                }

            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Autorizaciones", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }

        private void borrar_linea_solic()
        {
            SForm.Freeze(true);
            //int nRow = (int)UMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            //nRow = nRow == -1 ? UMatrix.RowCount : nRow ;
            if (rowsel > 0)
            {
                SMatrix.GetLineData(rowsel);
                //  Verificando si tiene autorizadas
                string nodoc = oDbLinesDataSource.GetValue("U_numOC", rowsel - 1);
                if (ya_Procesada(nodoc))
                {
                    B1.Application.SetStatusBarMessage("Ese artículo no se puede borrar pues ya se procesó su Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    string lindel = oDbLinesDataSource.GetValue("code", rowsel - 1);
                    lineasdel.Add(lindel);
                    SMatrix.DeleteRow(rowsel);
                    SMatrix.FlushToDataSource();
                    SMatrix.LoadFromDataSource();
                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                    btn_crear.Caption = "Actualizar";
                    SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            SForm.Freeze(false);

            }

        private void insertar_linea_solic()
        {
            SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            btn_crear.Caption = "Actualizar";
            SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            SMatrix.AddRow(1, SMatrix.RowCount);
            SMatrix.ClearRowData(SMatrix.RowCount);
            SMatrix.FlushToDataSource();
            SMatrix.LoadFromDataSource();
            SMatrix.Columns.Item(1).Cells.Item(SMatrix.RowCount).Click(BoCellClickType.ct_Double);
        }

        private string obtener_Estado(string abrev)
        {
            string resultado = "";
            switch (abrev)
            {
                case "N":
                    {
                        resultado = "Nueva";
                    }
                    break;

                case "A":
                    {
                        resultado = "Aprob";
                    }
                    break;

                case "T":
                    {
                        resultado = "Trans";
                    }
                    break;

                case "C":
                    {
                        resultado = "Canc";
                    }
                    break;

                case "D":
                    {
                        resultado = "Dev";
                    }
                    break;

                case "F":
                    {
                        resultado = "Final";
                    }
                    break;
            }
            return resultado;
        }

        private void aprobar_solicitud()
        {
            try
            {
                SForm.Freeze(true);
                string sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                string scom = "\n - Aprobada: " +DateTime.Now.Date.ToString("ddMMyyyy");
                string sestado = "A";
                Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3}='{6}', {7} = '{9}', {8} = '{10}'  FROM {1} WHERE {0} = '{4}' ",
                                         Constantes.View.CAB_RVT.U_numOC,   //0
                                         Constantes.View.CAB_RVT.CAB_RV,    //1
                                         Constantes.View.CAB_RVT.U_comment, //2
                                         Constantes.View.CAB_RVT.U_estado,  //3
                                         sCode,                             //4
                                         scom,                              //5
                                         sestado,                          //6
                                         Constantes.View.CAB_RVT.U_idAut, //7
                                         Constantes.View.CAB_RVT.U_aut,  //8
                                         obtener_Vendedor(),          //9
                                         obtener_NameVendedor());   //10
                oRecordSet.DoQuery(SQLQuery);

                //Aprobar todos los articulos
                sestado = "Y";
                SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}' FROM {1} WHERE {0} = '{3}' ",
                                         Constantes.View.DET_RVT.U_numOC,   //0
                                         Constantes.View.DET_RVT.DET_RV,    //1
                                         Constantes.View.DET_RVT.U_estado,   //2
                                         sCode,                             //3
                                         sestado);                          //4
                oRecordSet.DoQuery(SQLQuery);

                cargar_inicial();

                SForm.Freeze(false);
                B1.Application.SetStatusBarMessage("Solicitud procesada con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error aprobando solicitud: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
         }

        private void transferir_aceptados()
        {
            bool todoOk = true;
            int result = 0;
            try
            {
                SForm.Freeze(true);
                string sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                GC.Collect();
                B1.Company.StartTransaction();
                SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                 
                doctransf.DocDate = DateTime.Today;
                doctransf.TaxDate = DateTime.Today;
                doctransf.FromWarehouse = "CD";
                doctransf.ToWarehouse = "SHOWROOM";
                doctransf.JournalMemo = "Generado por Addons VentasRT Solicitud: " + sCode;

                if(SMatrix.RowCount > 1)
                {
                    string artcurrent = "";
                    string art = "";
                    double totalart = 0.00;
                    int cantlines = 1;
                    int linestransf = 0;
                    for (int i = 1; i <= SMatrix.RowCount; i++)
                    {
                        art = (SMatrix.Columns.Item(1).Cells.Item(i).Specific).Value.ToString();
                        if (artcurrent != art)
                        {
                            if (artcurrent != "" && obtener_exist_articulo(artcurrent)>0)
                            {
                                if (cantlines > 1)
                                {
                                    result = doctransf.Lines.Count;
                                    doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                    doctransf.Lines.Add();
                                    doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                    linestransf++;
                                }
                                cantlines++;
                                doctransf.Lines.ItemCode = artcurrent;
                                doctransf.Lines.ItemDescription = (SMatrix.Columns.Item(2).Cells.Item(i-1).Specific).Value.ToString();
                                doctransf.Lines.Quantity = totalart;
                                doctransf.Lines.FromWarehouseCode = "CD";
                                doctransf.Lines.WarehouseCode = "SHOWROOM";
                            }
                            artcurrent = art;
                            totalart = Double.Parse((SMatrix.Columns.Item(5).Cells.Item(i).Specific).Value.ToString()) / 1000000.00; ;
                        }
                        else
                        {
                            totalart += Double.Parse((SMatrix.Columns.Item(5).Cells.Item(i).Specific).Value.ToString()) / 1000000.00;
                        }
                    }
                    // Adicionar ultima fila
                    if (artcurrent != "" && obtener_exist_articulo(artcurrent) > 0)
                    {
                        if (cantlines > 1)
                        {
                            result = doctransf.Lines.Count;
                            doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                            doctransf.Lines.Add();
                            doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                            linestransf++;
                        }
                        doctransf.Lines.ItemCode = artcurrent;
                        doctransf.Lines.ItemDescription = (SMatrix.Columns.Item(2).Cells.Item(SMatrix.RowCount).Specific).Value.ToString();
                        doctransf.Lines.Quantity = totalart;
                        doctransf.Lines.FromWarehouseCode = "CD";
                        doctransf.Lines.WarehouseCode = "SHOWROOM";


                        result = doctransf.Add();
                        GC.Collect();
                        todoOk = (result == 0) && (linestransf >0);
                    }
                }


                if (todoOk)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    string newkey = B1.Company.GetNewObjectKey();

                    //Actualizar datos de Transferencia en Solicitud
                    string scom = "\n - Transferida: " + DateTime.Now.Date.ToString("ddMMyyyy") +" DocNum:" + newkey;
                    string sestado = "T";
                    Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3}='{6}', {7} = '{8}' FROM {1} WHERE {0} = '{4}' ",
                                             Constantes.View.CAB_RVT.U_numOC,   //0
                                             Constantes.View.CAB_RVT.CAB_RV,    //1
                                             Constantes.View.CAB_RVT.U_comment, //2
                                             Constantes.View.CAB_RVT.U_estado,  //3
                                             sCode,                             //4
                                             scom,                              //5
                                             sestado,                          //6
                                             Constantes.View.CAB_RVT.U_idTR,  //7
                                             newkey);                         //8

                    oRecordSet.DoQuery(SQLQuery);

                    //Actualizar datos de Transferencia en articulos autorizados
                    sestado = "Y";
                    SQLQuery = String.Format("UPDATE {1} SET {5} = '{6}' FROM {1} WHERE {0} = '{3}' AND {2} = '{4}' ",
                                             Constantes.View.DET_RVT.U_numOC,   //0
                                             Constantes.View.DET_RVT.DET_RV,    //1
                                             Constantes.View.DET_RVT.U_estado,   //2
                                             sCode,                             //3
                                             sestado,                         //4
                                             Constantes.View.DET_RVT.U_idTV,   //5
                                             newkey);                           //6

                    oRecordSet.DoQuery(SQLQuery);

                    cargar_inicial();
                }
                else
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                SForm.Freeze(false);
                B1.Application.SetStatusBarMessage("Solicitud Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Transferiendo solicitud autorizada: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
        }  
      
    }
}
