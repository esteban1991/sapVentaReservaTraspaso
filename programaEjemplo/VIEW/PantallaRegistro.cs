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
using System.Windows.Forms;


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
        private List<string> lineasnodisp = new List<string>();
        private int rowsel = 0;
        private int indice = 0;
        private string docaprob = "";
        private string tractual = "";
        private string cominicial = "";
        private bool cabinserted = false;

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
        private SAPbouiCOM.EditText txt_log = null;
        private SAPbouiCOM.Matrix mtx = null;
        private SAPbouiCOM.Button btn_crear = null;
        private SAPbouiCOM.Button btn_cancel = null;
        private SAPbouiCOM.Button btn_autorizar = null;
        private SAPbouiCOM.Button btn_tr = null;
        private SAPbouiCOM.Button btn_cancelar = null;
        private SAPbouiCOM.Button btn_tv = null;
        private PantallaAprobac fa = null;


        public PantallaRegistro(PantallaAprobac faref, bool registro = true, string doc="" )
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Registro.srf"], "SolRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            formActual = "SolRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);
            registrar = registro;
            docaprob = doc;
            fa = faref;
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
                        case "1292":   //Adicionar linea
                            if (ItemActiveMenu == ventaRT.Constantes.View.registro.mtx)
                            {
                                insertar_linea_solic();
                                BubbleEvent = false;
                            }
                            break;
                        case "1293":  //Borrar linea
                            if (ItemActiveMenu== ventaRT.Constantes.View.registro.mtx)
                            {
                                borrar_linea_solic();
                                BubbleEvent = false;
                            }
                            break;
                        case "1290":    // Primero                      
                            activar_primero();
                            BubbleEvent = false;
                            break;
                        case "1289":    // Ant                      
                            activar_anterior();
                            BubbleEvent = false;
                            break;
                        case "1288":    // Sig                      
                            activar_posterior();
                            BubbleEvent = false;
                            break;
                        case "1291":    // Ultimo                      
                            activar_ultimo();
                            BubbleEvent = false;
                            break;
                        case "773":    // Pegar                     
                            insertar_lineas_necesarias();
                            break;
                    }
                    //BubbleEvent = true;
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
                ItemActiveMenu = eventInfo.ItemUID;
                rowsel = eventInfo.Row;
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
                            case BoEventTypes.et_FORM_CLOSE:
                                {
                                    if (!registrar && fa != null && encontrar_formulario())
                                    {
                                        fa.ThisSapApiForm.Form.Select();
                                        fa.cargar_datos_matriz();
                                    }
                                }
                                break;

                            case BoEventTypes.et_COMBO_SELECT:
                                {
                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.registro.cbnd:
                                            {
                                                if (!cabinserted)
                                                {
                                                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                                    string buscado = oCombo.Selected.Value == null ? "0" : oCombo.Selected.Value.ToString();
                                                    if (buscado != "0")
                                                    {
                                                        //indice = Int32.Parse(buscado);
                                                        cargar_solicitud(buscado, true);
                                                    }
                                                }
                                                else { cabinserted = false; }
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
                                                            if (registrar)
                                                            {
                                                                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                                                string buscado = oCombo.Selected.Value == null ? " 0" : oCombo.Selected.Value.ToString();
                                                                if (buscado != "0")
                                                                {
                                                                    cargar_solicitud(buscado, true);
                                                                }
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
                                                procesar_solicitud(false,true);
                                                break;
                                            }

                                        case Constantes.View.registro.btn_TR:
                                            {
                                                transferir(false,true);
                                                break;
                                            }
                                        case Constantes.View.registro.btn_cancelar:
                                            {
                                                procesar_solicitud(false,false);
                                                break;
                                            }

                                        case Constantes.View.registro.btn_TV:
                                            {
                                                transferir(false,false);
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
                                                // Validar que no existan repetidos articulo y cliente en el documento
                                                if (artsel != "" && codcli != "" && !validar_art_cliente_unicos(artsel, codcli, pVal.Row))
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error Datos Repetidos: Artículo y Cliente deben ser únicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                    BubbleEvent = false;
                                                }
                                                // Validar que tenga existencia en la Bodega Principal CD
                                                if (Ok)
                                                {
                                                    if (!(obtener_exist_articulo(artsel, "CD") > 0))
                                                    {
                                                        Ok = false;
                                                        B1.Application.SetStatusBarMessage("Error el Artículo no tienen disponibilidad en la Bodega Principal", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                                                    oDbLinesDataSource.SetValue("U_cant", nRow - 1, "1");
                                                    oDbLinesDataSource.SetValue("U_onHand", nRow - 1, obtener_exist_articulo(artsel, "CD").ToString());
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
                                                    B1.Application.SetStatusBarMessage("Error Datos Repetidos: Artículo y Cliente deben ser únicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                                            if (pVal.ItemUID == "mtx" && pVal.ColUID == "cant")
                                            {
                                                bool Ok = true;
                                                string codart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                                string codcli = CFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString();
                                                if (codart != "" && codcli != "" && !validar_art_cliente_unicos(codart, codcli, pVal.Row))
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error Datos Repetidos: Artículo y Cliente deben ser únicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                    BubbleEvent = false;
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
                                                if (btn_crear.Caption == "Actualizar" || btn_crear.Caption == "Crear") 
                                                {
                                                    guardar_solicitud();
                                                    BubbleEvent = false;
                                                }
                                                else
                                                {
                                                    if (btn_crear.Caption == "Buscar" && registrar)
                                                    {
                                                        oCombo.Item.Click(BoCellClickType.ct_Regular);
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
                                    if (pVal.InnerEvent == false && pVal.ItemUID == "mtx")
                                    {
                                        string codart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        string codcli = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codCli").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        string cantart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        switch (pVal.ColUID)
                                        {
                                            case "codArt":
                                                {
                                                    if (codart == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error Código Artículo es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (codart != "" && codcli != "" && !validar_art_cliente_unicos(codart, codcli, pVal.Row))
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error Datos Repetidos: Articulo y Cliente deben ser únicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = false;
                                                        }
                                                    }
                                                }
                                                break;
                                            case "codCli":
                                                {
                                                    if (codcli == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error Código Cliente es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (codart != "" && codcli != "" && !validar_art_cliente_unicos(codart, codcli, pVal.Row))
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error Datos Repetidos: Artículo y Cliente deben ser únicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = false;
                                                        }
                                                    }
                                                }
                                                break;
                                            case "cant":
                                                {
                                                    if (codart != "" && codcli != "" && !validar_art_cliente_unicos(codart, codcli, pVal.Row))
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error Datos Repetidos: Artículo y Cliente deben ser únicos por Solicitud", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value = "" ;
                    
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (cantart == "")
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error: Cantidad debe ser superior a  0 ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = false;
                                                        }
                                                        else
                                                        {
                                                            double cantidad = Double.Parse(((SAPbouiCOM.EditText)SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Specific).Value.ToString());
                                                            double disp = Double.Parse(((SAPbouiCOM.EditText)SMatrix.Columns.Item("onHand").Cells.Item(pVal.Row).Specific).Value.ToString());
                                                            if (cantidad == 0 && disp != 0)
                                                            {
                                                                B1.Application.SetStatusBarMessage("Error Cantidad debe ser superior a 0", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                BubbleEvent = false;
                                                            }
                                                            else
                                                            {
                                                                if (cantidad > disp)
                                                                {
                                                                    B1.Application.SetStatusBarMessage("Error Cantidad > Disponibilidad", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                    BubbleEvent = false;
                                                                }

                                                            }
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
                string errormsg = ex.Message == "" ? "Error del Addon VentaRT, notifique a su administrador" : ex.Message;
                B1.Application.SetStatusBarMessage("Error: " + errormsg, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }

        }

         
        // Metodos No Override

        private void cargar_inicial()
        {
            Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            // borrando tablas !!!
            //string SQLQuery = String.Empty;
            // Borrar lineas detalle
            //SQLQuery = String.Format("DELETE FROM {0}",
            //                  Constantes.View.DET_RVT.DET_RV,
            //                   Constantes.View.DET_RVT.U_numOC);
            //oRecordSet.DoQuery(SQLQuery);

            //SQLQuery = String.Format("DELETE FROM {0}",
            //                       Constantes.View.CAB_RVT.CAB_RV,
            //                        Constantes.View.CAB_RVT.U_numOC);

            //oRecordSet.DoQuery(SQLQuery);

            //string t = "3";
            //string e = "D" ;
            //string SQLQuery = String.Format("UPDATE {0} SET {3} = '{4}'  FROM {0} WHERE {1} = '{2}' ",
            //                       Constantes.View.CAB_RVT.CAB_RV,
            //                       Constantes.View.CAB_RVT.U_numOC, t,
            //                       Constantes.View.CAB_RVT.U_estado, e);
            //oRecordSet.DoQuery(SQLQuery);

            SForm = B1.Application.Forms.ActiveForm;
            SMatrix = SForm.Items.Item("mtx").Specific;
            formActual = B1.Application.Forms.ActiveForm.UniqueID;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
            oDbLinesDataSource = SForm.DataSources.DBDataSources.Item("@DET_RSTV");

            btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;
            btn_autorizar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_autorizar).Specific;
            btn_tr = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TR).Specific;
            btn_cancelar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancelar).Specific;
            btn_tv = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TV).Specific;
            oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
            txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            txt_log = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_log).Specific;
            mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
            txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
            txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
            txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
            txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
            txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
            txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
            txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_aut).Specific;
            txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
            txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
            txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
            txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_aut).Specific;

            SForm.EnableMenu("4870", false);  //filtrar matriz
            SForm.EnableMenu("8802", false);  //maxim matriz
            oCombo.Item.AffectsFormMode = false;

            txt_estado.Item.AffectsFormMode = false;
            txt_idtr.Item.AffectsFormMode = false;
            txt_idtv.Item.AffectsFormMode = false;
            txt_idaut.Item.AffectsFormMode = false;
            txt_aut.Item.AffectsFormMode = false;
            txt_log.Item.AffectsFormMode = false;


            if (registrar)
            {
                // Vendedor
                oCombo.Active = true;
                btn_autorizar.Item.Visible = false;
                btn_tr.Item.Visible = false;
                btn_cancelar.Item.Visible = false;
                btn_tv.Item.Visible = false;

                SForm.EnableMenu("1290", true); SForm.EnableMenu("1289", true); // mov entre registros
                SForm.EnableMenu("1288", true); SForm.EnableMenu("1291", true);
                SForm.EnableMenu("1282", true); // crear solicitud
                SForm.EnableMenu("1281", false);  //buscar
                SForm.EnableMenu("772", true); SForm.EnableMenu("773", true); //copiar y pegar

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
                SForm.EnableMenu("1282", false); 
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

                SAPbouiCOM.Column oColumn = SMatrix.Columns.Item("codArt");
                oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);

                cominicial = txt_com.Value;

                txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                txt_estado.Value = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                txt_estado.Value = txt_estado.Value == "" ? "R" : txt_estado.Value;

                txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                txt_idtr.Value = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                txt_idtr.Value = obtener_DocNum(txt_idtr.Value);

                txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                txt_idtv.Value = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);
                txt_idtv.Value = obtener_DocNum(txt_idtv.Value);

                string estadoactual = txt_estado.Value.ToString().Substring(0, 1) ;
                txt_estado.Value = obtener_Estado(estadoactual);
             

                oColumn = SMatrix.Columns.Item("estado");
                oColumn.Editable = (estadoactual == "A" || estadoactual == "C");

                // Recargar DocNum de Transferencia o Devolucion
                string dentry = "";
                int cantautoriz = 0;
                for (int i = 1; i <= SMatrix.RowCount; i++)
                {
                    dentry = (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value.ToString();
                    (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value = obtener_DocNum(dentry);
                    SMatrix.CommonSetting.SetCellEditable(i, 7, (estadoactual == "A" ||estadoactual == "C")) ;
                    if ((SMatrix.Columns.Item(7).Cells.Item(i).Specific).Checked ) {cantautoriz ++;}
                }
                SMatrix.AutoResizeColumns();

                btn_autorizar.Item.Enabled = estadoactual == "R";
                btn_tr.Item.Enabled = estadoactual == "A" && cantautoriz > 0;
                btn_cancelar.Item.Enabled = estadoactual == "R"  ;
                btn_tv.Item.Enabled = estadoactual == "C" && cantautoriz > 0;
            }
            SForm.Freeze(false);
        }

        private bool insertar_solicitud()
        {
            bool todoOk = true;
            try {

                    bool contraercombo = (txt_numoc.Value.ToString() == "");
                    mtx.Item.Enabled = true;
                    txt_com.Item.Enabled = true;
                    txt_com.Active = true;

                    SForm.EnableMenu("1292", true); //Activar Agregar Linea
                    SForm.EnableMenu("1293", true); //Activar Borrar Linea 

                    B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    int norecord = obtener_ultimo_ID("CA") + 1;
               
                    //Insertando nuevo record
                    // FILTRAR LAS SOLICITUDES DEL USUARIO ACTUAL
                    SAPbouiCOM.Conditions orCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition orCon = orCons.Add();
                    orCon.Alias = "U_idVend";
                    orCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    orCon.CondVal = B1.Company.UserName;

                    oDbHeaderDataSource.Query(orCons);
                    oDbHeaderDataSource.InsertRecord(oDbHeaderDataSource.Size);
                    oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size-1;

                    DateTime fc = DateTime.Now.Date;
                    //fc = fc.AddDays(-30);  // de prueba para q se cancele automatico
                    DateTime fv = fc.AddDays(10);

                    oDbHeaderDataSource.SetValue("U_numDoc", norecord, norecord.ToString());
                    oDbHeaderDataSource.SetValue("U_IdVend", norecord, obtener_Vendedor());
                    oDbHeaderDataSource.SetValue("U_vend", norecord, obtener_NameVendedor());
                    oDbHeaderDataSource.SetValue("U_fechaC", norecord, fc.ToString("yyyyMMdd"));
                    oDbHeaderDataSource.SetValue("U_fechaV", norecord, fv.ToString("yyyyMMdd"));
                    oDbHeaderDataSource.SetValue("U_estado", norecord, "Reservada");
                    oDbHeaderDataSource.SetValue("U_comment", norecord, "");
                    oDbHeaderDataSource.SetValue("U_logs", norecord, "");

                    mtx.Clear();
                    mtx.AddRow(1, 1);
                    mtx.ClearRowData(1);
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
                    txt_log.Value = "";
                    txt_estado.Value = "Reservada" ;
                    btn_crear.Caption = "Crear";

                    // adicionar ultima fila a combo de busqueda

                    if (contraercombo)
                    {
                        cabinserted = true;
                        oCombo.Item.Click(BoCellClickType.ct_Collapsed);
                    }
                    string nvalue = norecord.ToString();
                    oCombo.ValidValues.Add(
                        nvalue,
                        fc.ToString("yyyyMMdd"));

                    indice = oCombo.ValidValues.Count;
                    cabinserted = true;
                    oCombo.Select(nvalue, BoSearchKey.psk_ByValue);

                    mtx.Columns.Item("codArt").Cells.Item(1).Click();
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
            txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_aut).Specific;
            txt_log = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_log).Specific;
            txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
            txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
            txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
            btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;
            btn_autorizar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_autorizar).Specific;
            btn_tr = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TR).Specific;
            btn_cancelar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancelar).Specific;
            btn_tv = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_TV).Specific;


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
                                oCombo.ValidValues.Remove(txt_numoc.Value.ToString(), BoSearchKey.psk_ByValue); 
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
                        txt_log.Value = "";
                        txt_idaut.Value = "";
                        txt_aut.Value = "";
                        SMatrix.Item.Enabled = false;
                        txt_com.Item.Enabled = false;
                        oCombo.Active = true;
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
            if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    indice = 1;
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
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
            if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    if (indice > 1)
                    {
                        indice--;
                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                        oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
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

            if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    if (indice < oDbHeaderDataSource.Size)
                    {
                        indice++;
                        if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && indice > 1) { indice--; }
                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                        oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
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
            if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
            {
                B1.Application.SetStatusBarMessage("No se puede mover porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                try
                {
                    indice = oDbHeaderDataSource.Size;
                    if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && indice > 1) { indice--; }
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
                    B1.Application.SetStatusBarMessage("Movimiento al Último ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error en Movimiento al Último ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    throw ex;
                }
            }
      }

        private bool eliminar_solicitud()
        {
            bool todoOk = true;
            string serror = "";

            if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("En Modo Insertar no procede eliminacion... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;
            }

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

                                // eliminar linea en combo seleccion
                                oCombo.ValidValues.Remove(abuscar,BoSearchKey.psk_ByValue);

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
            bool concurrente = false;
            string serror = "";
            string sCode = ""; string sName = "";
            int iRet;
            string sidtr = "";
            todoOk = validar_art_cliente_unicos_todos();
            if (!todoOk) { serror = "Datos Repetidos (Articulo y Cliente deben ser Unicos para una Solicitud)"; }
            else
            {

                try
                {
                    SAPbobsCOM.UserTable UTDoc = B1.Company.UserTables.Item("CAB_RSTV");
                    SAPbobsCOM.UserTable UTLines = B1.Company.UserTables.Item("DET_RSTV");
                    //SForm.Freeze(true);

                    try
                    {
                        // Salvando documento 
                        int norecord = Int32.Parse(txt_numoc.Value.ToString());
                        sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                        string sfechav = oDbHeaderDataSource.GetValue("U_fechaV", oDbHeaderDataSource.Offset);
                        string sfechac = oDbHeaderDataSource.GetValue("U_fechaC", oDbHeaderDataSource.Offset);
                        string sestado = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                        sestado = (sestado == "" ? "R" : sestado).Substring(0, 1);

                        string scom = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                        string slog = oDbHeaderDataSource.GetValue("U_logs", oDbHeaderDataSource.Offset);
                        string svend = oDbHeaderDataSource.GetValue("U_idVend", oDbHeaderDataSource.Offset);
                        string snvend = oDbHeaderDataSource.GetValue("U_vend", oDbHeaderDataSource.Offset);
                        string saut = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                        string snaut = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                        sidtr = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                        string sidtv = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);

                        //Guardando con instrucciones SQL
                        //Buscar si existe ese codigo para update
                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string SQLQuery = String.Format("SELECT {0},{3} FROM {1} WHERE {0} = '{2}'",
                                        Constantes.View.CAB_RVT.U_numOC,
                                        Constantes.View.CAB_RVT.CAB_RV,
                                        sCode,
                                        Constantes.View.CAB_RVT.U_idVend);

                        oRecordSet.DoQuery(SQLQuery);
                        oRecordSet.MoveFirst();
                        if (!oRecordSet.EoF)
                        {
                            concurrente =  (oRecordSet.Fields.Item("U_idVend").Value.ToString() != svend);
                        }

                        if (!oRecordSet.EoF && !concurrente)
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

                            // comprobando ultimo id x si concurrencia
                            int norecordant = norecord;
                            norecord = obtener_ultimo_ID("CA") + 1;

                            if (norecordant != norecord)
                            {
                                oCombo.ValidValues.Remove(sCode, BoSearchKey.psk_ByValue);
                                sCode = norecord.ToString();
                                oDbHeaderDataSource.SetValue("U_numDoc", norecordant, sCode);
                                oCombo.ValidValues.Add(
                                    sCode,
                                    fc.ToString("yyyyMMdd"));
                                indice = oCombo.ValidValues.Count;
                                cabinserted = true;
                                //oCombo.Select(sCode, BoSearchKey.psk_ByValue);
                            }


                            //fc = fc.AddDays(-30);  // de prueba para q se cancele automatico
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
                        int norecord2 = obtener_ultimo_ID("DE");

                        SMatrix.FlushToDataSource();
                        for (int i = 0; i <= oDbLinesDataSource.Size - 1; i++)
                        {

                            // Obteniendo texto de los campos de DbDataSource
                            string sCodeL = oDbLinesDataSource.GetValue("Code", i);
                            string sNameL = oDbLinesDataSource.GetValue("Name", i);
                            string scodart = oDbLinesDataSource.GetValue("U_codArt", i);
                            string sart = oDbLinesDataSource.GetValue("U_articulo", i);
                            string scodcli = oDbLinesDataSource.GetValue("U_codCli", i);
                            string sccli = oDbLinesDataSource.GetValue("U_cliente", i);
                            string scant = oDbLinesDataSource.GetValue("U_cant", i);
                            string sdisp = oDbLinesDataSource.GetValue("U_onHand", i);
                            string sestad = oDbLinesDataSource.GetValue("U_estado", i);

                            if (scodart != "" && scodcli != "" && scant != "")
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
                                        UTLines.UserFields.Fields.Item("U_cant").Value = Double.Parse(scant) / 1000000.00;
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
                                    if (todoOk) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); }

                                }
                            }


                        }
                        UTLines = null;
                    }
                    else { todoOk = false; }
                }
                catch (Exception ex)
                {
                    todoOk = false;
                    serror = ex.Message;
                    throw;
                }
                finally
                {
                    //SForm.Freeze(false);
                    System.GC.Collect();
                }
            }

            if (todoOk)
            {
                todoOk = eliminar_filas_borradas();
            }
 
            // Transfiriendo 
            if (todoOk)
            {
                if (registrar)
                {
                    if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        // Transferir al Virtual
                        transferir(true, false);
                    }
                    else
                    {
                        // Cancelar y Transferir 
                        revertir(sCode, tractual);
                    }
                }

            }

            if (todoOk){
               B1.Application.SetStatusBarMessage("Solicitud guardada con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
               SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
               SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)SForm.Items.Item("1").Specific;
               btn_crear.Caption = "OK";
               if (registrar)
               {
                   // recargando para actualizar posicion de la navegacion
                   oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                   oCombo.Select(sCode, BoSearchKey.psk_ByValue);
               }
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
                   B1.Application.SetStatusBarMessage("Artículos cargados con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
               }
               else
               {
                   B1.Application.SetStatusBarMessage("Error cargando líneas: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                                int respuesta = B1.Application.MessageBox("¿Desea cancelar los datos modificados? ", 1, "OK", " Cancelar");
                                if (respuesta == 1)
                                {
                                    if (B1.Company.InTransaction)
                                    {
                                        B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                    }
                                    if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        if (oCombo.ValidValues.Count == oDbHeaderDataSource.Size+1)
                                        { 
                                            oCombo.ValidValues.Remove(txt_numoc.Value.ToString(), BoSearchKey.psk_ByValue); 
                                        }
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
                            txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                            txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                            txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                            txt_log = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_log).Specific;
                            mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                            string usrCurrent = B1.Company.UserName;

                                int nuevaposic = 0;
                                if (!posicion)
                                {
                                    // Navegacion normal
                                    //Buscando posicion fisica
                                    nuevaposic = Int32.Parse(noDoc) + 1;
                                    
                                    string SQLQuery = String.Format("SELECT TOP {2} CAST({0} AS INT) AS ND" +
                                        " FROM {1} WHERE {3} = '{4}' ORDER BY CAST({0} AS INT)  ASC",
                                                        Constantes.View.CAB_RVT.U_numOC,  //0
                                                        Constantes.View.CAB_RVT.CAB_RV, //1
                                                        nuevaposic.ToString(),///2
                                                        Constantes.View.CAB_RVT.U_idVend,//3
                                                        usrCurrent);                        //4                                

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
                                    string SQLQuery = String.Format("SELECT {0} FROM {1}  WHERE {2} = '{3}' ",
                                                        Constantes.View.CAB_RVT.U_numOC,  //0
                                                        Constantes.View.CAB_RVT.CAB_RV,  //1
                                                        Constantes.View.CAB_RVT.U_idVend,//2
                                                        usrCurrent);                     //3 

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
                                        indice = indice == 0 ? nuevaposic + 1 : indice;
                                    }
                                }
                                // Vendedor
                                // FILTRAR LAS SOLICITUDES DEL USUARIO ACTUAL
                                SAPbouiCOM.Conditions orCons = new SAPbouiCOM.Conditions();
                                SAPbouiCOM.Condition orCon = orCons.Add();
                                orCon.Alias = "U_idVend";
                                orCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                orCon.CondVal = B1.Company.UserName;

                                oDbHeaderDataSource.Query(orCons);

                                // Carga de la Solicitud Encontrada

                                nuevaposic = nuevaposic < 0 ? 0 : nuevaposic;
                                oDbHeaderDataSource.Offset = nuevaposic;

                                oDbHeaderDataSource.Query(orCons);

                                // Carga Inicial de Datos si no esta hecha
                                if (docaprob == "")
                                {
                                    docaprob = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                                }

                                txt_numoc.Value = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                                txt_estado.Value = oDbHeaderDataSource.GetValue("U_estado", oDbHeaderDataSource.Offset);
                                txt_idtr.Value = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                                txt_idtv.Value = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);
                                txt_fechac.Value = oDbHeaderDataSource.GetValue("U_fechaC", oDbHeaderDataSource.Offset);
                                txt_fechav.Value = oDbHeaderDataSource.GetValue("U_fechaV", oDbHeaderDataSource.Offset);
                                txt_idvend.Value = oDbHeaderDataSource.GetValue("U_idVend", oDbHeaderDataSource.Offset);
                                txt_vend.Value = oDbHeaderDataSource.GetValue("U_vend", oDbHeaderDataSource.Offset);
                                txt_idaut.Value = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                                txt_aut.Value = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                                txt_com.Value = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                                txt_log.Value = oDbHeaderDataSource.GetValue("U_logs", oDbHeaderDataSource.Offset);
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
                        B1.Application.SetStatusBarMessage("Solicitud cargada con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        todoOk = cargar_lineas(txt_numoc.Value.ToString());
                        if (todoOk)
                        {
                            bool procesar = true;

                                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                btn_crear.Caption = "OK";
                                txt_estado.Value = obtener_Estado(txt_estado.Value);
                                tractual = txt_idtr.Value.ToString();
                                txt_idtr.Value = obtener_DocNum(txt_idtr.Value);
                                txt_idtv.Value = obtener_DocNum(txt_idtv.Value);

                                procesar = !ya_Procesada(txt_numoc.Value.ToString());

                            SForm.EnableMenu("1292", registrar && procesar); //Activar Agregar Linea
                            SForm.EnableMenu("1293", registrar && procesar); //Activar Borrar Linea 
                            txt_log.Active = true;
                            mtx.Item.Enabled = procesar;
                            txt_com.Item.Enabled = procesar;
                            txt_com.Active = procesar;
                        }
                        else
                        {
                            B1.Application.SetStatusBarMessage("Error cargando Solicitud ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        }

                    }
                    else
                    {
                        B1.Application.SetStatusBarMessage("Error cargando líneas: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    }

                    SForm.Freeze(false);
                    return todoOk;
                }
                else { return true; }
            }
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

        private bool validar_art_cliente_unicos_todos()
        {
            bool todoOK = true;
            string art = "";
            string cli = "";

            if (SMatrix.RowCount > 1)
            {
                try
                {
                    for (int j = 1; j <= SMatrix.RowCount && todoOK; j++)
                    {
                        art = (SMatrix.Columns.Item(1).Cells.Item(j).Specific).Value.ToString();
                        cli = (SMatrix.Columns.Item(3).Cells.Item(j).Specific).Value.ToString();
                        todoOK = validar_art_cliente_unicos(art, cli, j);
                    }
                }
                catch (Exception ex)
                {
                    B1.Application.SetStatusBarMessage("Error validando campos repetidos" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    todoOK = false;
                    throw;
                }
            }
            if (!todoOK)
            {
                B1.Application.SetStatusBarMessage("Error: Artículo y Cliente Repetidos: " + art + "-" + cli , SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                    return estado.Substring(0,1) != "R" ;
                }

            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Autorizaciones", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
        }

        private void borrar_linea_solic()
        {
           string abuscar = txt_numoc.Value.ToString();
           if (registrar && !ya_Procesada(abuscar))
           {
               SForm.Freeze(true);
               if (rowsel > 0)
               {
                   if (rowsel <= oDbLinesDataSource.Size) // Verificar si la linea ya ha sido salvada al dbsource
                   {
                       string lindel = oDbLinesDataSource.GetValue("code", rowsel - 1);
                       lineasdel.Add(lindel);
                   }
                   SMatrix.DeleteRow(rowsel);
                   SMatrix.FlushToDataSource();
                   SMatrix.LoadFromDataSource();
                   SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                   btn_crear.Caption = SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE ? "Actualizar" : btn_crear.Caption;
                   if (SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                   {
                       SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                   }

               }
               SForm.Freeze(false);
           }
           else
           {
               B1.Application.SetStatusBarMessage("No se puede borrar líneas porque ya es una Solicitud Procesada", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
           }
      }

        private void insertar_linea_solic()
        {

           string abuscar = txt_numoc.Value.ToString();
           if (registrar && !ya_Procesada(abuscar))
           {
               SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
               mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
               btn_crear.Caption = SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE ? "Actualizar" : btn_crear.Caption;
               if (SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
               {
                   SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
               }

               int posfila = mtx.RowCount + 1;
               mtx.AddRow(1, posfila);
               mtx.ClearRowData(posfila);
               if (SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
               {
                   mtx.FlushToDataSource();
                   mtx.LoadFromDataSource();
               }
               mtx.Columns.Item(1).Cells.Item(posfila).Click(BoCellClickType.ct_Double);
           }
           else
           {
               B1.Application.SetStatusBarMessage("No se puede insertar líneas porque ya es una Solicitud Procesada", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
           }
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

        private void procesar_solicitud(bool crear, bool aprobar)
        {
            try
            {
                bool continuar = true;
                string terror = "";
                if (!crear && !aprobar)
                {
                    continuar = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset) != "";
                    if (!continuar) { terror = "Al cancelar, es recomendable comentar la causa....."; }
                    else
                    {
                        // chequear que el comentario sea diferente
                        continuar = (cominicial != oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset));
                        if (!continuar) { terror = "Al cancelar, es recomendable cambiar el  comentario existente....."; }
                    }
                }
                if (!continuar)
                {
                    int respuesta = B1.Application.MessageBox(terror, 1, "OK");
                }
                else
                {
                    SForm.Freeze(true);
                    string sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                    string scom = crear ? "Reservada" : (aprobar ? "Aprobada: " : "Cancelada: ") + DateTime.Now.Date.ToString("dd/MM/yyyy");
                    string sestado = crear ? "R" : (aprobar ? "A" : "C");
                    Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    // Buscando logs actual
                    string SQLQuery = String.Format("SELECT {2} FROM {1} WHERE {0} = '{3}' ",
                                               Constantes.View.CAB_RVT.U_numOC,
                                               Constantes.View.CAB_RVT.CAB_RV,
                                               Constantes.View.CAB_RVT.U_logs,
                                               sCode);

                    oRecordSet.DoQuery(SQLQuery);
                    oRecordSet.MoveFirst();
                    string logant = "";
                    if (!oRecordSet.EoF)
                    {
                        logant = oRecordSet.Fields.Item("U_logs").Value;
                    }

                    scom = logant + "\r\n-" + scom + "\r\n";

                    SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3}='{6}', {7} = '{9}', {8} = '{10}'  FROM {1} WHERE {0} = '{4}' ",
                                             Constantes.View.CAB_RVT.U_numOC,   //0
                                             Constantes.View.CAB_RVT.CAB_RV,    //1
                                             Constantes.View.CAB_RVT.U_logs, //2
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
                    //sestado = (crear || aprobar) ? "Y" : "N" ;
                    sestado = "Y";
                    //string wherecancel = (crear ||aprobar) ? "" : " AND " + Constantes.View.DET_RVT.U_estado + " = 'Y' ";
                    //SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}' FROM {1} WHERE {0} = '{3}' {5} ",
                    SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}' FROM {1} WHERE {0} = '{3}'  ",
                                             Constantes.View.DET_RVT.U_numOC,   //0
                                             Constantes.View.DET_RVT.DET_RV,    //1
                                             Constantes.View.DET_RVT.U_estado,   //2
                                             sCode,                             //3
                                             sestado);                          //4
                    // wherecancel);                     //5
                    oRecordSet.DoQuery(SQLQuery);

                    cargar_inicial();

                    SForm.Freeze(false);
                    B1.Application.SetStatusBarMessage("Solicitud " + (crear ? "Reservada:" : (aprobar ? "Autorizada" : "Cancelada")) + " con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error " + (crear ? "reservando" : (aprobar? "autorizando" : " cancelando")) + " solicitud: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
         }

        private void transferir(bool crear, bool aprobar)
        {
            bool todoOk = true;
            int result = 0;
            string tv = "";
            string terror = "";
            try
            {
                SForm.Freeze(true);
                string sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                GC.Collect();
                B1.Company.StartTransaction();
                SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                doctransf.DocDate = DateTime.Today;
                doctransf.TaxDate = DateTime.Today;
                // Serie Primaria
                doctransf.Series = 27;
                doctransf.FromWarehouse = crear ? "CD" : "CD_RSV";
                doctransf.ToWarehouse = crear ? "CD_RSV" : "CD";
                doctransf.JournalMemo = "Addons VentasRT al " + (crear? "Reservar" : (aprobar ? "Aprobar" : "Cancelar")) + " Solic:" + sCode;

                SAPbouiCOM.Column oColumn = SMatrix.Columns.Item("codArt");
                oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);

                if (SMatrix.RowCount > 0)
                {
                    string artcurrent = "";
                    string art = "";
                    bool autoriz = false;
 
                    double totalart = 0.00;
                    int cantlines = 1;
                    int linestransf = 0;
                    double disponible = 0;
                    for (int i = 1; i <= SMatrix.RowCount; i++)
                    {
                        art = (SMatrix.Columns.Item(1).Cells.Item(i).Specific).Value.ToString();
                        autoriz = (SMatrix.Columns.Item(7).Cells.Item(i).Specific).Checked;
                        tv = (SMatrix.Columns.Item(8).Cells.Item(i).Specific).Value.ToString();
                        //if ((crear) || ((autoriz == aprobar) && (aprobar ? true : tv != "")))
                        if (crear || (!crear && autoriz))
                        {
                            if (artcurrent != art)
                            {
                                if (artcurrent != "")
                                {
                                    disponible = obtener_exist_articulo(artcurrent, crear ? "CD" : "CD_RSV");
                                    if (disponible < totalart)
                                    {
                                        // Procesar los articulos no disponibles a la hora de transferir
                                        lineasnodisp.Add(artcurrent);
                                    }
                                    else
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
                                        doctransf.Lines.ItemDescription = (SMatrix.Columns.Item(2).Cells.Item(i - 1).Specific).Value.ToString();
                                        doctransf.Lines.Quantity = totalart;
                                        doctransf.Lines.FromWarehouseCode = crear ? "CD" : "CD_RSV";
                                        doctransf.Lines.WarehouseCode = crear ? "CD_RSV" : "CD";
                                    }
                                }
                                artcurrent = art;
                                totalart = Double.Parse((SMatrix.Columns.Item(5).Cells.Item(i).Specific).Value.ToString()) / 1000000.00; ;
                            }
                            else
                            {
                                totalart += Double.Parse((SMatrix.Columns.Item(5).Cells.Item(i).Specific).Value.ToString()) / 1000000.00;
                            }
                        }

                    }
                    // Adicionar ultima fila
                    autoriz = (SMatrix.Columns.Item(7).Cells.Item(SMatrix.RowCount).Specific).Checked;
                    tv = (SMatrix.Columns.Item(8).Cells.Item(SMatrix.RowCount).Specific).Value.ToString();
                    //if ((crear) || ((autoriz == aprobar) && (aprobar ? true : tv != "")))
                    if (crear || (!crear && autoriz))
                    {
                        if (artcurrent != "")
                        {
                            disponible = obtener_exist_articulo(artcurrent, crear ? "CD" : "CD_RSV");
                            if (disponible < totalart)
                            {
                                // Procesar los articulos no disponibles a la hora de transferir
                                lineasnodisp.Add(artcurrent);
                            }
                            else
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
                                doctransf.Lines.ItemDescription = (SMatrix.Columns.Item(2).Cells.Item(SMatrix.RowCount).Specific).Value.ToString();
                                doctransf.Lines.Quantity = totalart;
                                doctransf.Lines.FromWarehouseCode = crear ? "CD" : "CD_RSV";
                                doctransf.Lines.WarehouseCode = crear ? "CD_RSV" : "CD";
                            }
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
                        terror = "No existen artículos disponibles. "+infonodisp;
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
                        string scom = (crear ? "Reservada" : (aprobar?"Aprobada":"Cancelada")) + " Transferida: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + " DocNum:" + obtener_DocNum(newkey) + infonodisp;
                        string sestado = crear ? "R" : (aprobar ? "T" : "D") ;
                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        // Buscando logs actual
                        string SQLQuery = String.Format("SELECT {2} FROM {1} WHERE {0} = '{3}' ",
                                                   Constantes.View.CAB_RVT.U_numOC,
                                                   Constantes.View.CAB_RVT.CAB_RV,
                                                   Constantes.View.CAB_RVT.U_logs,
                                                   sCode);

                        oRecordSet.DoQuery(SQLQuery);
                        oRecordSet.MoveFirst();
                        string logant = "";
                        if (!oRecordSet.EoF)
                        {
                            logant = oRecordSet.Fields.Item("U_logs").Value;
                        }

                        scom = logant + "\r\n-" + scom + "\r\n";

                        SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3}='{6}', {7} = '{8}' FROM {1} WHERE {0} = '{4}' ",
                                                 Constantes.View.CAB_RVT.U_numOC,   //0
                                                 Constantes.View.CAB_RVT.CAB_RV,    //1
                                                 Constantes.View.CAB_RVT.U_logs, //2
                                                 Constantes.View.CAB_RVT.U_estado,  //3
                                                 sCode,                             //4
                                                 scom,                              //5
                                                 sestado,                          //6
                                                 crear ? Constantes.View.CAB_RVT.U_idTR : Constantes.View.CAB_RVT.U_idTV,  //7
                                                 newkey);                         //8

                        oRecordSet.DoQuery(SQLQuery);

                        //Actualizar datos de Transferencia en articulos autorizados o cancelados
                        //sestado = (crear || aprobar) ? "Y" : "N" ;
                        sestado =  "Y" ;
                        //string wherecancel = (crear || aprobar) ? "" : " AND "  + Constantes.View.DET_RVT.U_idTV + " != '' " ;
                        //string wherecancel = (crear) ? "" : " AND " + Constantes.View.DET_RVT.U_idTV + " != '' ";
                        //SQLQuery = String.Format("UPDATE {1} SET {5} = '{6}' FROM {1} WHERE {0} = '{3}' AND {2} = '{4}' {7}",
                        SQLQuery = String.Format("UPDATE {1} SET {5} = '{6}' FROM {1} WHERE {0} = '{3}' AND {2} = '{4}' ",
                                                 Constantes.View.DET_RVT.U_numOC,   //0
                                                 Constantes.View.DET_RVT.DET_RV,    //1
                                                 Constantes.View.DET_RVT.U_estado,   //2
                                                 sCode,                             //3
                                                 sestado,                         //4
                                                 Constantes.View.DET_RVT.U_idTV,   //5
                                                 newkey);                           //6
                                                 //wherecancel);                     //7

                        oRecordSet.DoQuery(SQLQuery);

                        //cancelar_filas_nodisp(newkey, crear, tv);
                        cancelar_filas_nodisp(newkey, crear, "" );
                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)SForm.Items.Item("1").Specific;
                        btn_crear.Caption = "OK";
                        if (crear)
                        { cargar_solicitud(sCode, true); }
                        else {cargar_inicial();}

                        B1.Application.SetStatusBarMessage("Solicitud" + (crear ? "Reservada" : (aprobar ? "Autorizada" : "Cancelada"))  + " Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    }
                    else
                    {
                        
                        B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud " + (crear ? "Reservada" : (aprobar? "Autorizada" : "Cancelada"))  , SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    }
 
                }
                else
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud " + (crear ? "Reservada " : (aprobar ? "Autorizada " : "Cancelada ")) + terror , SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }

                SForm.Freeze(false);

            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud " + (crear ? "Reservada" : (aprobar ? "Autorizada" : "Cancelada")) + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
        }

        private void revertir(string sCode, string docentry)
        {
            bool todoOk = true;
            int result = 0;
            string tv = "";
            string terror = "";

            try
            {
                SForm.Freeze(true);
                GC.Collect();
                B1.Company.StartTransaction();
                SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                doctransf.DocDate = DateTime.Today;
                doctransf.TaxDate = DateTime.Today;
                // Serie Primaria
                doctransf.Series = 27;
                doctransf.FromWarehouse = "CD_RSV";
                doctransf.ToWarehouse =  "CD";
                doctransf.JournalMemo = "Addons VentasRT al Revertir Solic:" + sCode;

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
                        string scom = "Reservada Revertida: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + " DocNum:" + obtener_DocNum(newkey) + infonodisp;

                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        // Buscando logs actual
                        string SQLQuery = String.Format("SELECT {2} FROM {1} WHERE {0} = '{3}' ",
                                                   Constantes.View.CAB_RVT.U_numOC,
                                                   Constantes.View.CAB_RVT.CAB_RV,
                                                   Constantes.View.CAB_RVT.U_logs,
                                                   sCode);

                        oRecordSet.DoQuery(SQLQuery);
                        oRecordSet.MoveFirst();
                        string logant = "";
                        if (!oRecordSet.EoF)
                        {
                            logant = oRecordSet.Fields.Item("U_logs").Value;
                        }

                        scom = logant + "\r\n-" + scom + "\r\n";
                        

                        SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}' FROM {1} WHERE {0} = '{3}' ",
                                                 Constantes.View.CAB_RVT.U_numOC,   //0
                                                 Constantes.View.CAB_RVT.CAB_RV,    //1
                                                 Constantes.View.CAB_RVT.U_logs,    //2
                                                 sCode,                             //3
                                                 scom);                              //4
                        oRecordSet.DoQuery(SQLQuery);
                        B1.Application.SetStatusBarMessage("Solicitud Reservada Revertida Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        // Transferir la actualizada
                        transferir(true, false);
                    }
                    else
                    {

                        B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud Reservada Revertida", SAPbouiCOM.BoMessageTime.bmt_Long, true);
                    }
                }
                else
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud Reservada Revertida " +terror , SAPbouiCOM.BoMessageTime.bmt_Long, true);


                    //Actualizar logs en Solicitud
                    string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                    string slog = "Error:No pudo ser Revertida por no tener disponibilidad: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + infonodisp;
                    string scom = "Solicitud sin disponibilidad al intentar Revertir: " + DateTime.Now.Date.ToString("dd/MM/yyyy");

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

                SForm.Freeze(false);


            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Transfiriendo Solicitud Solicitud Reservada Revertida" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
        }

        private bool cancelar_filas_nodisp(string newkey, bool crear , string oldkey)
        {
            bool todoOk = true;
            string SQLQuery = String.Empty;
            try
            {
                string filasnodisp = string.Join("-",lineasnodisp);
                if (lineasnodisp != null && lineasnodisp.Count>0)
                {
                    for (int i = 0; i < lineasnodisp.Count; i++)
                    {
                        string sestado = "N" ;
                        //string nuevaidtv = crear ? "" : "{7}" ;
                        string nuevaidtv =  "";
                        SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3} = '{8}' FROM {1} WHERE {0} = '{4}' AND {3} = '{6}' ",
                                        Constantes.View.DET_RVT.U_codArt,      //0
                                        Constantes.View.DET_RVT.DET_RV,    //1 
                                        Constantes.View.DET_RVT.U_estado,  //2
                                        Constantes.View.DET_RVT.U_idTV,    //3
                                        lineasnodisp[i],                   //4
                                        sestado,                           //5
                                        newkey,                          //6
                                        oldkey,                       //7
                                        nuevaidtv);                    //8
                                     
                        Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsCards.DoQuery(SQLQuery);
                    }
                    int respuesta = B1.Application.MessageBox("Los artículos "+filasnodisp+" no están disponibles, por tanto, se cancela su transferencia", 1, "OK", " Cancelar");
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

        private bool encontrar_formulario()
        {
            bool encontrado = false;

            try
            {
                for (int i = 0; i < B1.Application.Forms.Count && !encontrado; i++)
                {
                    encontrado = (B1.Application.Forms.Item(i).UniqueID == fa.ThisSapApiForm.Form.UniqueID);
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error buscando formulario " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            return encontrado;
        }

        private int cantFilas_clipboard()
        {
            int filas = 0;

            try
            {

                string textocopiado = GetClipBoardData();
                filas = textocopiado.Split('\n').Length;
                
                //string textocopiado = Clipboard.GetText();
                //filas = textocopiado.Split('\n').Length;
  

                //string[] cb = textocopiado.Split(new string[1] { "\r\n" }, StringSplitOptions.None);
                //filas = cb.Length;



            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error gestionando clipboard " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                filas = 0;
            }
            return filas;
        }

        private void insertar_lineas_necesarias()
        {

            string abuscar = txt_numoc.Value.ToString();
            if (registrar && !ya_Procesada(abuscar))
            {
                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                btn_crear.Caption = SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE ? "Actualizar" : btn_crear.Caption;
                if (SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                int filasnuevas = cantFilas_clipboard() - (mtx.RowCount - rowsel + 1);
                if (filasnuevas > 0)
                {
                    int posfila = mtx.RowCount + 1;
                    mtx.AddRow(filasnuevas, posfila);
                    for (int i = 0; i < filasnuevas; i++)
                    {
                        mtx.ClearRowData(posfila + i);
                    }

                    if (SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        mtx.FlushToDataSource();
                        mtx.LoadFromDataSource();
                    }
                }


            }
            else
            {
                B1.Application.SetStatusBarMessage("No se puede insertar líneas porque ya es una Solicitud Procesada", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
        }

        private string GetClipBoardData()
        {
            try
            {
                string clipboardData = null;
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate()
                    {
                        try
                        {
                            clipboardData = Clipboard.GetText(TextDataFormat.Text);
                        }

                        catch (Exception ex)
                        {
                            threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                return clipboardData;
            }
            catch (Exception exception)
            {
                return string.Empty;
            }
        }


    }
}

