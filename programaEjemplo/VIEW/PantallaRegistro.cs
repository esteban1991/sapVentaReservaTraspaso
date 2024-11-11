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
using System.Globalization;

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
        private SAPbouiCOM.EditText txt_idcli = null;
        private SAPbouiCOM.EditText txt_cli = null;
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

        public PantallaRegistro(PantallaAprobac faref, bool registro = true, string doc="")
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
            string errorMessage = "";
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
                        {
                            errorMessage =  insertar_solicitud();
                            if (!string.IsNullOrEmpty(errorMessage)) 
                            {
                                HandleError(new Exception(errorMessage));
                            }
               
                            BubbleEvent = false;
                            break;
                        }
                        case "1292":   //Adicionar linea
                            if (ItemActiveMenu == ventaRT.Constantes.View.registro.mtx)
                            {
                                errorMessage =  insertar_linea_solic();
                                if (!string.IsNullOrEmpty(errorMessage)) 
                                {
                                    HandleError(new Exception(errorMessage));
                                }
                                
                                BubbleEvent = false;
                            }
                            break;
                        case "1293":  //Borrar linea
                            if (ItemActiveMenu== ventaRT.Constantes.View.registro.mtx)
                            {
                                errorMessage =  borrar_linea_solic();
                                if (!string.IsNullOrEmpty(errorMessage)) 
                                {
                                    HandleError(new Exception(errorMessage));
                                }
                                
                                BubbleEvent = false;
                            }
                            break;
                        case "1290":    // Primero  
                            errorMessage = activar_primero();
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                HandleError(new Exception(errorMessage));
                            }
                            BubbleEvent = false;
                            break;
                        case "1289":    // Ant                      
                            errorMessage = activar_anterior();
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                HandleError(new Exception(errorMessage));
                            }
                            BubbleEvent = false;
                            break;
                        case "1288":    // Sig  
                            errorMessage = activar_posterior();
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                HandleError(new Exception(errorMessage));
                            }                    
                            BubbleEvent = false;
                            break;
                        case "1291":    // Ultimo 
                            errorMessage = activar_ultimo();
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                HandleError(new Exception(errorMessage));
                            }                     
                             BubbleEvent = false;
                            break;
                        case "773":    // Pegar  
                            errorMessage = insertar_lineas_necesarias();
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                HandleError(new Exception(errorMessage));
                            }                       
                            
                            break;
                    }
                    //BubbleEvent = true;
                }
              }
            }
            catch (Exception ex)
            {
                errorMessage = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message; 
                B1.Application.SetStatusBarMessage("Error ejecutando menú: " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Long, true);
                throw;
            }
        }

        private void ThisSapApiForm_OnAfterRightClick(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorMessage = "";
            try
            {
                ItemActiveMenu = eventInfo.ItemUID;
                rowsel = eventInfo.Row;
            }
            catch (Exception ex)
            {
                errorMessage = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
                B1.Application.SetStatusBarMessage("Error activando menú: " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Long, true);
                throw;
            }
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
                                    this.B1.Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
                                    this.B1.Application.MenuEvent -= new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
                                    if (!registrar && fa != null && encontrar_formulario(out errorMessage))
                                    {
                                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                        fa.ThisSapApiForm.Form.Select();
                                        fa.cargar_datos_matriz();
                                    }
                                    addonGeneral.contadorRegistrosAbiertos--; // Decrementar el contador al cerrar el formulario
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
                                                        errorMessage =  cargar_solicitud(buscado, true);
                                                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
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
                                                                    errorMessage =  cargar_solicitud(buscado, true);
                                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                                }
                                                            }
                                                            BubbleEvent = false;
                                                            break;
                                                        }
                                                    case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                                                        {
                                                            errorMessage =  guardar_solicitud();
                                                            if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                            BubbleEvent = false;
                                                            break;
                                                        }
                                                    case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                                                        {
                                                            errorMessage =  guardar_solicitud();
                                                            if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                            BubbleEvent = false;
                                                            break;
                                                        }
                                                }
                                                break;
                                            }

                                        case Constantes.View.registro.btn_autorizar:
                                            {
                                                errorMessage = procesar_solicitud(false,true);
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                break;
                                            }

                                        case Constantes.View.registro.btn_TR:
                                            {
                                                errorMessage =  transferir(false,true);
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                break;
                                            }
                                        case Constantes.View.registro.btn_cancelar:
                                            {
                                                errorMessage =  procesar_solicitud(false,false);
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                break;
                                            }

                                        case Constantes.View.registro.btn_TV:
                                            {
                                                errorMessage =  transferir(false,false);
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
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
                                                string preciosel = CFLEvent.SelectedObjects.GetValue("AvgPrice", 0).ToString();
                                                double precio = 0.0;
                                                if (!Double.TryParse(preciosel, out precio))
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error: Precio > 0", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                    BubbleEvent = false;
                                                }
                                                // Validar que no existan repetidos articulo en el documento
                                                errorMessage = validar_art_unico(artsel, pVal.Row);
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                                                // Validar que tenga existencia en la Bodega Principal CD
                                                if (Ok)
                                                {
                                                    if (!(obtener_exist_articulo(artsel, "CD", out errorMessage) > 0))
                                                    {
                                                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                        Ok = false;
                                                        B1.Application.SetStatusBarMessage("Error: Artículo no tienen disponibilidad en la Bodega Principal", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                                                    oDbLinesDataSource.SetValue("U_price", nRow - 1, preciosel);
                                                    oDbLinesDataSource.SetValue("U_onHand", nRow - 1, obtener_exist_articulo(artsel, "CD", out errorMessage).ToString());
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                    SMatrix.LoadFromDataSource();
                                                    SMatrix.Columns.Item("cant").Cells.Item(nRow).Click();
                                                }
                                            }

                                            if (pVal.ItemUID == "txt_idcli")
                                            {
                                                bool Ok = true;
                                                string clisel = CFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString();

                                                if (Ok)
                                                {
                                                    oDbHeaderDataSource.SetValue("U_codCli", oDbHeaderDataSource.Offset, CFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString());
                                                    oDbHeaderDataSource.SetValue("U_cliente", oDbHeaderDataSource.Offset, CFLEvent.SelectedObjects.GetValue("CardName", 0).ToString());

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
                                        " WHERE  contains({3},'%{4}%') ORDER BY CAST(T0.{0} AS INT) ASC",
                                                                            Constantes.View.CAB_RVT.U_numDoc,
                                                                            Constantes.View.CAB_RVT.CAB_RV,
                                                                            Constantes.View.CAB_RVT.U_fechaC,
                                                                            Constantes.View.CAB_RVT.U_idVend,
                                                                            usrCurrent);
                                        oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                                        errorMessage  = llenar_combo_id(oCombo, SQLQuery);
                                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
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
                                                    errorMessage =  guardar_solicitud();
                                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
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

                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.registro.txt_idcli:
                                            {
                                                if (pVal.InnerEvent == false)
                                                {
                                                    SAPbouiCOM.EditText codcli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txt_idcli").Specific;
                                                    if (codcli.Value == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error: Código Cliente es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                }
                                            }
                                            break;

                                        case Constantes.View.registro.txt_fechav:
                                            {
                                                if (pVal.InnerEvent == false)
                                                {
                                                    SAPbouiCOM.EditText fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txt_fechac").Specific;
                                                    SAPbouiCOM.EditText fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item("txt_fechav").Specific;

                                                    if (fechav.Value == "" || fechav.Value == null)
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error: Fecha Vencimiento es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                        BubbleEvent = false;
                                                    }
                                                    // TODO VAlidar que sea > fechaC

                                                    else
                                                    {
                                                        DateTime tempdate;
                                                        DateTime tempdate2;
                                                        if (DateTime.TryParseExact(fechac.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out tempdate) &&
                                                            DateTime.TryParseExact(fechav.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out tempdate2))
                                                        {
                                                            if ((DateTime.ParseExact(fechav.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture) -
                                                                    DateTime.ParseExact(fechac.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)).Days < 0)
                                                            {
                                                                B1.Application.SetStatusBarMessage("Error: Fecha Documento <= Vencimiento", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                BubbleEvent = true;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error: Fecha Vencimiento no posee formato adecuado..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                            BubbleEvent = true;
                                                        }
                                                    }

                                                }
                                            }
                                            break;


                                        case Constantes.View.registro.mtx:
                                        {
                                            if (pVal.InnerEvent == false)
                                            {
                                                string codart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                                string cantart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Specific).Value.ToString();
                                                string price = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific).Value.ToString();

                                                double amount_item = 0;
                                                try
                                                {
                                                    double cant = Double.Parse(cantart) / 1000000.00;
                                                    double precio = Double.Parse(price) / 1000000.00;
                                                    amount_item = cant * precio;
                                                }
                                                catch
                                                {
                                                    amount_item = 0;
                                                }


                                                switch (pVal.ColUID)
                                                {
                                                    case "codArt":
                                                        {
                                                            if (codart == "")
                                                            {
                                                                B1.Application.SetStatusBarMessage("Error: Código Artículo es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                BubbleEvent = false;
                                                            }
                                                            else
                                                            {
                                                                errorMessage = validar_art_unico(codart, pVal.Row);
                                                                if (!string.IsNullOrEmpty(errorMessage)) { 
                                                                    HandleError(new Exception(errorMessage)); 
                                                                }

                                                                SMatrix.Columns.Item("amount").Cells.Item(pVal.Row).Specific.Value = amount_item.ToString();
                                                            }
                                                        }
                                                        break;

                                                    case "cant":
                                                        {
                                                            errorMessage = validar_art_unico(codart, pVal.Row);
                                                            if (!string.IsNullOrEmpty(errorMessage)) { 
                                                                HandleError(new Exception(errorMessage));
                                                            }
                                                            else
                                                            {
                                                                if (cantart == "")
                                                                {
                                                                    B1.Application.SetStatusBarMessage("Error: Cantidad es requerida. ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                    BubbleEvent = false;
                                                                }
                                                                else
                                                                {
                                                                    double cantidad = Double.Parse(((SAPbouiCOM.EditText)SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Specific).Value.ToString());
                                                                    double disp = Double.Parse(((SAPbouiCOM.EditText)SMatrix.Columns.Item("onHand").Cells.Item(pVal.Row).Specific).Value.ToString());
                                                                    if (cantidad <= 0 && disp != 0)
                                                                    {
                                                                        B1.Application.SetStatusBarMessage("Error: Cantidad debe ser superior a 0", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                        BubbleEvent = false;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (cantidad > disp)
                                                                        {
                                                                            B1.Application.SetStatusBarMessage("Error: Cantidad > Disponibilidad", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                                            BubbleEvent = false;
                                                                        }
                                                                        else
                                                                        {

                                                                            SMatrix.Columns.Item("amount").Cells.Item(pVal.Row).Specific.Value = amount_item.ToString();
                                                                        }
                                                                    }
                                                                }
                                                                
                                                            }

                                                        }
                                                        break;
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
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
                B1.Application.SetStatusBarMessage("Error: " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Long, true);
                //throw;
            }

        }

         
        // Metodos No Override

        private void HandleError(Exception ex)
        {
            if (B1.Company.InTransaction)
            {
                B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
            string msgError = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
            B1.Application.SetStatusBarMessage("Error: " + msgError, SAPbouiCOM.BoMessageTime.bmt_Long, true);
        }

        private string num_lineas()
        {
            string errorMessage = "";
            try
            {
                for (int i = 1; i <= mtx.RowCount; i++)
                {
                    mtx.Columns.Item(0).Cells.Item(i).Specific.Value = i.ToString();
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Numerar Líneas: " + ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }


        private string cargar_inicial()
        {
            string errorMessage = "";
            try
            {
                Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                txt_idcli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idcli).Specific;
                txt_cli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_cli).Specific;
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

                SAPbouiCOM.Column cAmount; 
                cAmount = (SAPbouiCOM.Column)mtx.Columns.Item("amount"); 
                cAmount.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

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
                    txt_idtr.Value = obtener_DocNum(txt_idtr.Value, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        SForm.Freeze(false);
                        return errorMessage;
                    }

                    txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                    txt_idtv.Value = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);
                    txt_idtv.Value = obtener_DocNum(txt_idtv.Value, out errorMessage);

                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        SForm.Freeze(false);
                        return errorMessage;
                    }

                    string estadoactual = txt_estado.Value.ToString().Substring(0, 1);
                    txt_estado.Value = obtener_Estado(estadoactual);

                    oColumn = SMatrix.Columns.Item("estado");
                    oColumn.Editable = (estadoactual == "A" || estadoactual == "C");

                    // Recargar DocNum de Transferencia o Devolucion
                    string dentry = "";
                    int cantautoriz = 0;
                    for (int i = 1; i <= SMatrix.RowCount; i++)
                    {
                        dentry = (SMatrix.Columns.Item("idTV").Cells.Item(i).Specific).Value.ToString();
                        (SMatrix.Columns.Item("idTV").Cells.Item(i).Specific).Value = obtener_DocNum(dentry, out errorMessage);
                        if (!string.IsNullOrEmpty(errorMessage))
                        {
                            SForm.Freeze(false);
                            return errorMessage;
                        }
                        SMatrix.CommonSetting.SetCellEditable(i, 6, (estadoactual == "A" || estadoactual == "C"));
                        if ((SMatrix.Columns.Item("estado").Cells.Item(i).Specific).Checked) { cantautoriz++; }
                    }
                    SMatrix.AutoResizeColumns();

                    btn_autorizar.Item.Enabled = estadoactual == "R";
                    btn_tr.Item.Enabled = estadoactual == "A" && cantautoriz > 0;
                    btn_cancelar.Item.Enabled = estadoactual == "R";
                    btn_tv.Item.Enabled = estadoactual == "C" && cantautoriz > 0;
                }
                SForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Cargar Inicial: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string insertar_solicitud()
        {
            string errorMessage = "";
            try {
                    bool contraercombo = (txt_numoc.Value.ToString() == "");
                    mtx.Item.Enabled = true;
                    txt_fechav.Item.Enabled = true;
                    txt_idcli.Item.Enabled = true;
                    txt_idcli.Active = true;
                    txt_com.Item.Enabled = true;
                    SForm.EnableMenu("1292", true); //Activar Agregar Linea
                    SForm.EnableMenu("1293", true); //Activar Borrar Linea 

                    B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    int norecord = obtener_ultimo_ID("CA", out errorMessage) + 1;
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        
                        return errorMessage;
                    }
               
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
                    oDbHeaderDataSource.SetValue("U_idVend", norecord, obtener_Vendedor(out errorMessage));
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    oDbHeaderDataSource.SetValue("U_vend", norecord, obtener_NameVendedor(out errorMessage));
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    oDbHeaderDataSource.SetValue("U_codCli", norecord, "");
                    oDbHeaderDataSource.SetValue("U_cliente", norecord, "");
                    oDbHeaderDataSource.SetValue("U_fechaC", norecord, fc.ToString("yyyyMMdd"));
                    oDbHeaderDataSource.SetValue("U_fechaV", norecord, fv.ToString("yyyyMMdd"));
                    oDbHeaderDataSource.SetValue("U_estado", norecord, "Reservada");
                    oDbHeaderDataSource.SetValue("U_comment", norecord, "");
                    oDbHeaderDataSource.SetValue("U_logs", norecord, "");

                    mtx.Clear();
                    mtx.AddRow(1, 1);
                    mtx.ClearRowData(1);
                    txt_numoc.Value = norecord.ToString();
                    txt_idvend.Value = obtener_Vendedor(out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {

                        return errorMessage;
                    }
                    txt_vend.Value = obtener_NameVendedor(out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {

                        return errorMessage;
                    } 
                    txt_idaut.Value = "";
                    txt_aut.Value = "";
                    txt_idcli.Value = "";
                    txt_cli.Value = "";
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
                    errorMessage =  "Insertar Solicitud: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
                }
                return errorMessage;

        }

        private string preparar_modo_Find()
        {
            string errorMessage = "";
            try
            {
                int borrado = 0;
                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                txt_idcli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idcli).Specific;
                txt_cli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_cli).Specific;

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
                    return "Cambiar a Modo Buscar: No se puede activar Modo Búsqueda porque no tiene registros... ";
                }
                else
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

                        }
                        else 
                        { 
                            return "Cambiar a Modo Buscar: Cancelado porque existen datos en edición... "; 
                        }
                    }

                    SForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    txt_fechav.Item.Enabled = false;
                    txt_idcli.Item.Enabled = false;
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
                    txt_idcli.Value = "";
                    txt_cli.Value = "";
                    SMatrix.Item.Enabled = false;
                    txt_com.Item.Enabled = false;
                    oCombo.Active = true;
                    
                }
            }
           catch (Exception ex)
            {
                errorMessage =  "Cambiar a Modo Buscar: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string activar_primero()
        {
            string errorMessage = "";
            try
                {
                if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    return "No se puede mover al Registro Inicial porque no tiene registros... ";
                }
                indice = 1;
                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                errorMessage =  "Mover a Registro Inicial: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string activar_anterior()
        {
            string errorMessage = "";
            try
            {
                if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    return "No se puede mover al Registro Anterior porque no tiene registros..... ";
                }
                if (indice > 1)
                {
                    indice--;
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Mover a Registro Anterior: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string activar_posterior() 
        {
            string errorMessage = "";
            try
            {
                if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    return "No se puede mover al Registro Siguiente porque no tiene registros... ";
                }
                if (indice < oDbHeaderDataSource.Size)
                {
                    indice++;
                    if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && indice > 1) { indice--; }
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Mover a Registro Siguiente: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string activar_ultimo()
        {
            string errorMessage = "";
            try
            {
                if (oDbHeaderDataSource.Size == 0 || (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    return "No se puede mover al Registro Final porque no tiene registros... ";
                }
                indice = oDbHeaderDataSource.Size;
                if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && indice > 1) { indice--; }
                oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                oCombo.Select(indice.ToString(), BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                errorMessage =  "Mover a Registro Final: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private bool ya_Procesada(string nodoc, out string errorMessage)
        {
            errorMessage = "";
            
            try
            {
                String strSQL = String.Format("SELECT {1} FROM {2}  Where {0}='{3}'",
                                    Constantes.View.CAB_RVT.U_numDoc,
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
                errorMessage =  "Determinar Estado: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }  
            return false;
        }        
     
        private string eliminar_solicitud()
        {
            string errorMessage = "";
            try
            {
                string abuscar = txt_numoc.Value.ToString();
                bool estado = ya_Procesada(abuscar, out errorMessage);

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    return errorMessage;
                }

                if (estado)
                {
                    return "Eliminar Solicitud: Solicitud Procesada, no se puede eliminar..";
                }

                string SQLQuery = String.Format("SELECT {0}, {2} FROM {1}",
                                        Constantes.View.CAB_RVT.U_numDoc,
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
                if (!encontrado)
                {
                    return "Eliminar Solicitud: Documento no encontrado";
                }

                // Validar que sea Nueva sino no se puede borrar
                if (estadoactual != "N")
                {
                    return "Eliminar Solicitud: Documento en Proceso, no se puede Eliminar";
                }

                oDbHeaderDataSource.RemoveRecord(i - 1);
                SQLQuery = String.Format("DELETE FROM {1} WHERE {0} = '{2}' ",
                                Constantes.View.CAB_RVT.U_numDoc,
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

                return (oDbHeaderDataSource.Offset == 0) ? activar_primero(): activar_anterior(); 
            }
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Eliminar Solicitud: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;

        }

        private string validar_art_unico( string art, int row)
        {
            string errorMessage = "";
            try
            {
                if(SMatrix.RowCount > 1)
                {
                    int creg = 0;
                    for (int i = 1; i <= SMatrix.RowCount && creg < 1; i++)
                    {
                        if ((i != row) &&
                            (SMatrix.Columns.Item(1).Cells.Item(i).Specific).Value.ToString() == art )
                        {
                            creg++;
                        }
                    }
                    if (creg >= 1)
                    {
                        return "Artículo debe ser único para cada Solicitud";
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Validar Artículo: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string validar_art_unico_todos()
        {
            string errorMessage = "";
            try
            {
                if (SMatrix.RowCount > 1)
                {
                    bool todoOK = true;
                    string art = "";
                    for (int j = 1; j <= SMatrix.RowCount && todoOK; j++)
                    {
                        art = (SMatrix.Columns.Item(1).Cells.Item(j).Specific).Value.ToString();
                        todoOK = string.IsNullOrEmpty(validar_art_unico(art, j));
                    }
                    if (!todoOK)
                    {
                        return "Artículo Repetido: " + art;
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Validando Artículos: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private int obtener_ultimo_ID(string tipo, out string errorMessage)
        {
            errorMessage = "";
            int CodeNum = 0;
            try 
            {
                if (tipo == "CA")
                {
                    String strSQL = String.Format("SELECT TOP 1 CAST(T0.{0} AS INT) AS nd FROM {1} T0 ORDER BY CAST(T0.{0} AS INT) DESC",
                                        Constantes.View.CAB_RVT.U_numDoc,
                                        Constantes.View.CAB_RVT.CAB_RV);

                    Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsCards.DoQuery(strSQL);
                    string Code = rsCards.Fields.Item("nd").Value.ToString();
                    //probar cuando la tabla este vacia, osea el primero registro y no haya otro anterior
                    if (Code != "")
                    {
                        CodeNum = Convert.ToInt32(Code);
                    }
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
                        CodeNum = Convert.ToInt32(Code);
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Generar ID de la Solicitud: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return CodeNum;
        }
       
        private string eliminar_filas_borradas()
        {
            string errorMessage = "";
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
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Eliminar Filas Borradas: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                lineasdel.Clear();
                System.GC.Collect();
            }
            return errorMessage;
        }

        private string guardar_solicitud()
        {
            string errorMessage = "";
            bool concurrente = false;
            string sCode = ""; 
            int iRet = 0;
            string sidtr = "";

            try
            {
                if (!string.IsNullOrEmpty(validar_art_unico_todos()))
                {
                   return "Guardar Solicitud: Datos Repetidos (Artículo debe ser Único para una Solicitud)"; 
                }
                SAPbobsCOM.UserTable UTDoc = B1.Company.UserTables.Item("CAB_RSTV");
                SAPbobsCOM.UserTable UTLines = B1.Company.UserTables.Item("DET_RSTV");
                //SForm.Freeze(true);

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
                string scli = oDbHeaderDataSource.GetValue("U_codCli", oDbHeaderDataSource.Offset);
                string sncli = oDbHeaderDataSource.GetValue("U_cliente", oDbHeaderDataSource.Offset);
                string saut = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                string snaut = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                sidtr = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                string sidtv = oDbHeaderDataSource.GetValue("U_idTV", oDbHeaderDataSource.Offset);

                double amount_total = 0.00;

                if (SMatrix != null)
                {
                    SAPbouiCOM.Column cAmount;
                    cAmount = (SAPbouiCOM.Column)SMatrix.Columns.Item("amount");
                    cAmount.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    string str_amount = cAmount.ColumnSetting.SumValue.ToString().Replace(".", "").Replace(",",".");
                    amount_total = Double.Parse(str_amount) / 1000000.00;
                }



                //Guardando con instrucciones SQL
                //Buscar si existe ese codigo para update
                Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string SQLQuery = String.Format("SELECT {0},{3} FROM {1} WHERE {0} = '{2}'",
                                Constantes.View.CAB_RVT.U_numDoc,
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
                    SQLQuery = String.Format("UPDATE {1} SET {2} = '{4}', {5}='{6}', {7}='{9}', {8}='{10}', {11}='{12}', {13}='{14}'  WHERE {0} = '{3}' ",
                                        Constantes.View.CAB_RVT.U_numDoc,   //0
                                        Constantes.View.CAB_RVT.CAB_RV,    //1
                                        Constantes.View.CAB_RVT.U_comment, //2
                                        sCode,                             //3
                                        scom,                              //4
                                        Constantes.View.CAB_RVT.U_estado,  //5
                                        sestado,                          //6
                                        Constantes.View.CAB_RVT.U_codCli, //7
                                        Constantes.View.CAB_RVT.U_cliente, //8
                                        scli, //9,
                                        sncli, //10
                                        Constantes.View.CAB_RVT.U_fechaV, //11
                                        sfechav, //12
                                        Constantes.View.CAB_RVT.U_amount, //13
                                        amount_total //14
                                        ); 
                    oRecordSet.DoQuery(SQLQuery);
                }
                else
                {
                    // INSERT
                    DateTime fc = DateTime.Now.Date;
                    // comprobando ultimo id x si concurrencia
                    int norecordant = norecord;
                    norecord = obtener_ultimo_ID("CA", out errorMessage) + 1;
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        return errorMessage;
                    }

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
                    SQLQuery = String.Format("INSERT INTO {0} ({7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{25},{26},{29}) " +
                    " VALUES('{1}','{2}','{3}','{4}','{5}','{6}','{21}','{23}','{24}','{1}','{1}','{20}', '{21}','{27}', '{28}', '{30}' ) ",
                                        Constantes.View.CAB_RVT.CAB_RV,    //0
                                        sCode, //1
                                        svend, //2
                                        fc.ToString("yyyyMMdd"), //3
                                        fv.ToString("yyyyMMdd"), //4
                                        sestado, //5
                                        scom, //6
                                        Constantes.View.CAB_RVT.U_numDoc, //7
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
                                        sidtv,  //24
                                        Constantes.View.CAB_RVT.U_codCli, //25
                                        Constantes.View.CAB_RVT.U_cliente, //26
                                        scli, //27
                                        sncli, //28
                                        Constantes.View.CAB_RVT.U_amount, //29
                                        amount_total //30
                                        );
                    oRecordSet.DoQuery(SQLQuery);
                }

                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                UTDoc = null;
                  
                //Salvando lineas del documento
                if (SMatrix != null)
                {
                    int norecord2 = obtener_ultimo_ID("DE", out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        return errorMessage;
                    }
                    SMatrix.FlushToDataSource();
                    for (int i = 0; i <= oDbLinesDataSource.Size - 1; i++)
                    {
                        // Obteniendo texto de los campos de DbDataSource
                        string sCodeL = oDbLinesDataSource.GetValue("Code", i);
                        string sNameL = oDbLinesDataSource.GetValue("Name", i);
                        string scodart = oDbLinesDataSource.GetValue("U_codArt", i);
                        string sart = oDbLinesDataSource.GetValue("U_articulo", i);
                        string scant = oDbLinesDataSource.GetValue("U_cant", i);
                        string sprice = oDbLinesDataSource.GetValue("U_price", i);
                        string sdisp = oDbLinesDataSource.GetValue("U_onHand", i);
                        string sestad = oDbLinesDataSource.GetValue("U_estado", i);

                        double amount_item = 0;
                        try
                        {
                            double cant = Double.Parse(scant) / 1000000.00;
                            double precio = Double.Parse(sprice) / 1000000.00;
                            amount_item = cant * precio;

                            //amount_item = Double.Parse((SMatrix.Columns.Item("cant").Cells.Item(i).Specific).Value.ToString()) * 
                            //      Double.Parse((SMatrix.Columns.Item("price").Cells.Item(i).Specific).Value.ToString());

                        }
                        catch
                        {
                            amount_item = 0;
                        }
                       


                        if (scodart != "" && scant != "" && sprice != "")
                        {
                                 // Guardando en la UserTable

                                B1.Company.StartTransaction();
                                if (UTLines.GetByKey(sCodeL))
                                {
                                    //UPDATE
                                    UTLines.UserFields.Fields.Item("U_codArt").Value = scodart;
                                    UTLines.UserFields.Fields.Item("U_articulo").Value = sart;
                                    UTLines.UserFields.Fields.Item("U_cant").Value = Double.Parse(scant) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_onHand").Value = Double.Parse(sdisp) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_price").Value = Double.Parse(sprice) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_amount").Value = amount_item;
                                    UTLines.UserFields.Fields.Item("U_estado").Value = sestad;
                                    UTLines.UserFields.Fields.Item("U_numOC").Value = sCode;
                                    iRet = UTLines.Update();
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
                                    UTLines.UserFields.Fields.Item("U_cant").Value = Double.Parse(scant) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_onHand").Value = Double.Parse(sdisp) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_price").Value = Double.Parse(sprice) / 1000000.00;
                                    UTLines.UserFields.Fields.Item("U_amount").Value = amount_item;
                                    UTLines.UserFields.Fields.Item("U_numOC").Value = sCode;
                                    iRet = UTLines.Add();
                                }
                                if (iRet != 0)
                                {
                                    if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                                    errorMessage = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : "";
                                    return errorMessage;
                                }
                                if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); }
                        }

                    }
                    
                    UTLines = null;
                }
 
                errorMessage = eliminar_filas_borradas();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    return errorMessage;
                }
        
                // Transfiriendo 
                if (registrar)
                {
                    if (SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                       // Transferir al Virtual
                       errorMessage = transferir(true, false);
                       if (!string.IsNullOrEmpty(errorMessage))
                       {
                           return errorMessage;
                       }
                    }
                    else
                    {
                       // Cancelar y Transferir 
                       errorMessage = revertir(sCode, tractual);
                       if (!string.IsNullOrEmpty(errorMessage))
                       {
                           return errorMessage;
                       }
                    }
                }

            B1.Application.SetStatusBarMessage("Solicitud guardada con éxito", SAPbouiCOM.BoMessageTime.bmt_Long, false);
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
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Guardar Solicitud: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                //SForm.Freeze(false);
                System.GC.Collect();
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
                    errorMessage =  "Obtener DocNum: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
                }  
            }
            return dnum;
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

        private string cargar_lineas(string noDoc)
        {
            string errorMessage = "";
            try
            {
                if (noDoc != "")
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
                        dentry = (SMatrix.Columns.Item("idTV").Cells.Item(i).Specific).Value.ToString();
                        string result = obtener_DocNum(dentry, out errorMessage);
                        if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                        (SMatrix.Columns.Item("idTV").Cells.Item(i).Specific).Value = result;
                        double amount_item =0;
                        try{
                            double cant =  Double.Parse((SMatrix.Columns.Item("cant").Cells.Item(i).Specific).Value.ToString())/ 1000000.00;
                            double precio = Double.Parse((SMatrix.Columns.Item("price").Cells.Item(i).Specific).Value.ToString())/ 1000000.00;
                            amount_item = cant * precio;

                          //amount_item = Double.Parse((SMatrix.Columns.Item("cant").Cells.Item(i).Specific).Value.ToString()) * 
                          //      Double.Parse((SMatrix.Columns.Item("price").Cells.Item(i).Specific).Value.ToString());

                        }
                        catch
                        {
                            amount_item = 0;
                        }
                         (SMatrix.Columns.Item("amount").Cells.Item(i).Specific).Value = amount_item;
                    }



                    SMatrix.AutoResizeColumns();
                    SAPbouiCOM.Column oColumn = SMatrix.Columns.Item("codArt");
                    oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);

                    errorMessage = num_lineas();
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Cargando Datos de Líneas: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string cargar_solicitud(string noDoc, bool posicion)
        {
            string errorMessage = "";
            try
            {
                if (oDbHeaderDataSource.Size == 0)
                {
                    return insertar_solicitud();
                }
                if (noDoc != "")
                {
                    SForm.Freeze(true);
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
                        }
                        else {
                            SForm.Freeze(false);
                            return "Cargar Solicitud: Cancelado por edición activa de datos";
                        }
                    }
                    oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
                    txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                    txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                    txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                    txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                    txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                    txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_vend).Specific;
                    txt_idcli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idcli).Specific;
                    txt_cli = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_cli).Specific;
  
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
                            " FROM {1}  contains({3},'%{4}%') ORDER BY CAST({0} AS INT)  ASC",
                                            Constantes.View.CAB_RVT.U_numDoc,  //0
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
                                                Constantes.View.CAB_RVT.U_numDoc,
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
                        string SQLQuery = String.Format("SELECT {0} FROM {1}  WHERE  contains({2},'%{3}%') ",
                                            Constantes.View.CAB_RVT.U_numDoc,  //0
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
                    orCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN;
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
                    txt_idcli.Value = oDbHeaderDataSource.GetValue("U_codCli", oDbHeaderDataSource.Offset);
                    txt_cli.Value = oDbHeaderDataSource.GetValue("U_cliente", oDbHeaderDataSource.Offset);
                    txt_idaut.Value = oDbHeaderDataSource.GetValue("U_idAut", oDbHeaderDataSource.Offset);
                    txt_aut.Value = oDbHeaderDataSource.GetValue("U_aut", oDbHeaderDataSource.Offset);
                    txt_com.Value = oDbHeaderDataSource.GetValue("U_comment", oDbHeaderDataSource.Offset);
                    txt_log.Value = oDbHeaderDataSource.GetValue("U_logs", oDbHeaderDataSource.Offset);

                    errorMessage =  cargar_lineas(txt_numoc.Value.ToString());
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        SForm.Freeze(false);
                        return errorMessage;
                    }

                    bool procesar = true;
                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                    SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    btn_crear.Caption = "OK";

                    txt_estado.Value = obtener_Estado(txt_estado.Value);
                    tractual = txt_idtr.Value.ToString();
                    txt_idtr.Value = obtener_DocNum(txt_idtr.Value, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        SForm.Freeze(false);
                        return errorMessage;
                    }

                    txt_idtv.Value = obtener_DocNum(txt_idtv.Value, out errorMessage);
                                
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        SForm.Freeze(false);
                        return errorMessage;
                    }

                    procesar = !ya_Procesada(txt_numoc.Value.ToString(), out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        SForm.Freeze(false);
                        return errorMessage;
                    }

                    SForm.EnableMenu("1292", registrar && procesar); //Activar Agregar Linea
                    SForm.EnableMenu("1293", registrar && procesar); //Activar Borrar Linea 
                    txt_log.Active = true;
                    txt_idcli.Item.Enabled = procesar;
                    txt_fechav.Item.Enabled = procesar;
                    mtx.Item.Enabled = procesar;
                    txt_com.Item.Enabled = procesar;
                    txt_idcli.Active = procesar;

                    B1.Application.SetStatusBarMessage("Solicitud cargada con éxito", SAPbouiCOM.BoMessageTime.bmt_Long, false);

                    SForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Cargando Solicitud: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;        

        }

        private double obtener_exist_articulo(string codart, string codwhs, out string errorMessage)
        {
            double exist = 0.00;
            errorMessage = "";
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
                errorMessage =  "Obtener Stock de Artículo: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return exist;
        }
 
        public string llenar_combo_id(SAPbouiCOM.ComboBox oCombo, string SqlQuery)
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
                oRecordSet.MoveFirst();

                oCombo.ValidValues.Add("0", "Seleccione Documento:");

                for (int i = 1; !oRecordSet.EoF; i++)
                {
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString(), oRecordSet.Fields.Item(1).Value.ToString("dd/MM/yyyy"));
                    oRecordSet.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            catch (Exception ex)
            {
                errorMessage =  "Cargando Datos de Lista" +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;         
        }

        private string obtener_NameVendedor(out string errorMessage)
        {
            errorMessage = "";
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
                errorMessage =  "Obtener Nombre Vendedor: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return "";    
        }

        private string obtener_Vendedor(out string errorMessage )
        {
            errorMessage = "";
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
                errorMessage =  "Obtener Vendedor: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return "";  
        }

        private string borrar_linea_solic()
        {
            string errorMessage = "";
            try
            {
                string abuscar = txt_numoc.Value.ToString();
                if (registrar && !ya_Procesada(abuscar, out errorMessage))
                {
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    SForm.Freeze(true);
                    if (rowsel > 0)
                    {
                        if (rowsel <= oDbLinesDataSource.Size) // Verificar si la linea ya ha sido salvada al dbsource
                        {
                            //string lindel = oDbLinesDataSource.GetValue("code", rowsel - 1);
                            string lindel = SMatrix.Columns.Item(4).Cells.Item(rowsel).Specific.Value.ToString();
                            lineasdel.Add(lindel);
                        }
                        SMatrix.DeleteRow(rowsel);
                        SMatrix.FlushToDataSource();
                        SMatrix.LoadFromDataSource();
                        errorMessage = num_lineas();
                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

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
                    return "Borrar Línea: No se puede borrar líneas porque ya es una Solicitud Procesada"; 
                }
            }
           catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Borrar Línea: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
      }

        private string insertar_linea_solic()
        {
            string errorMessage = "";
            try
            {

                string abuscar = txt_numoc.Value.ToString();
                if (registrar && !ya_Procesada(abuscar, out errorMessage))
                {
                    if (!string.IsNullOrEmpty(errorMessage)) {return errorMessage;}
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

                    errorMessage = num_lineas();
                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }

                    mtx.Columns.Item(1).Cells.Item(posfila).Click(BoCellClickType.ct_Double);
                }
                else
                {
                   return "Insertar Línea: No se puede insertar líneas porque ya es una Solicitud Procesada";
                }
            }
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Insertar Línea: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string procesar_solicitud(bool crear, bool aprobar)
        {
            string errorMessage = "";
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
                                               Constantes.View.CAB_RVT.U_numDoc,
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

                    string codVend = obtener_Vendedor(out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)){return errorMessage;}

                    string nameVend = obtener_NameVendedor(out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)){return errorMessage;}

                    SQLQuery = String.Format("UPDATE {1} SET {2} = '{5}', {3}='{6}', {7} = '{9}', {8} = '{10}'  FROM {1} WHERE {0} = '{4}' ",
                                             Constantes.View.CAB_RVT.U_numDoc,   //0
                                             Constantes.View.CAB_RVT.CAB_RV,    //1
                                             Constantes.View.CAB_RVT.U_logs, //2
                                             Constantes.View.CAB_RVT.U_estado,  //3
                                             sCode,                             //4
                                             scom,                              //5
                                             sestado,                          //6
                                             Constantes.View.CAB_RVT.U_idAut, //7
                                             Constantes.View.CAB_RVT.U_aut,  //8
                                             codVend,          //9
                                             nameVend);   //10
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
                    B1.Application.SetStatusBarMessage("Solicitud " + (crear ? "Reservada:" : (aprobar ? "Autorizada" : "Cancelada")) + " con éxito", SAPbouiCOM.BoMessageTime.bmt_Long, false);
                }
            }

            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Procesar Solicitud: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage; 
         }

        private string cancelar_filas_nodisp(string newkey, bool crear , string oldkey)
        {
            string errorMessage = "";
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
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage =  "Cancelar Líneas No Disponibles: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                lineasnodisp.Clear();
                System.GC.Collect();
            }
            return errorMessage;
        }

        private string transferir(bool crear, bool aprobar)
        {
            string errorMessage = "";

            int result = 0;
            string tv = "";

            try
            {
                SForm.Freeze(true);
                string sCode = oDbHeaderDataSource.GetValue("U_numDoc", oDbHeaderDataSource.Offset);
                string sCli = oDbHeaderDataSource.GetValue("U_codCli", oDbHeaderDataSource.Offset);
                string snCli = oDbHeaderDataSource.GetValue("U_cliente", oDbHeaderDataSource.Offset);
                GC.Collect();
                if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                B1.Company.StartTransaction();
                SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                doctransf.DocDate = DateTime.Today;
                doctransf.TaxDate = DateTime.Today;
                doctransf.CardCode = sCli;
                doctransf.CardName = snCli;

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
                        autoriz = (SMatrix.Columns.Item("estado").Cells.Item(i).Specific).Checked;
                        tv = (SMatrix.Columns.Item("idTV").Cells.Item(i).Specific).Value.ToString();
                        //if ((crear) || ((autoriz == aprobar) && (aprobar ? true : tv != "")))
                        if (crear || (!crear && autoriz))
                        {
                            if (artcurrent != art)
                            {
                                if (artcurrent != "")
                                {
                                    disponible = obtener_exist_articulo(artcurrent, crear ? "CD" : "CD_RSV", out errorMessage);
                                    if (!string.IsNullOrEmpty(errorMessage)){
                                        SForm.Freeze(false);
                                        return errorMessage;
                                    }

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
                                        doctransf.Lines.ItemDescription = (SMatrix.Columns.Item("articulo").Cells.Item(i - 1).Specific).Value.ToString();
                                        doctransf.Lines.Quantity = totalart;
                                        doctransf.Lines.FromWarehouseCode = crear ? "CD" : "CD_RSV";
                                        doctransf.Lines.WarehouseCode = crear ? "CD_RSV" : "CD";
                                    }
                                }
                                artcurrent = art;
                                totalart = Double.Parse((SMatrix.Columns.Item("cant").Cells.Item(i).Specific).Value.ToString()) / 1000000.00;                             }
                            else
                            {
                                totalart += Double.Parse((SMatrix.Columns.Item("cant").Cells.Item(i).Specific).Value.ToString()) / 1000000.00;
                            }
                        }
                    }
                    // Adicionar ultima fila
                    autoriz = (SMatrix.Columns.Item("estado").Cells.Item(SMatrix.RowCount).Specific).Checked;
                    tv = (SMatrix.Columns.Item("idTV").Cells.Item(SMatrix.RowCount).Specific).Value.ToString();
                    //if ((crear) || ((autoriz == aprobar) && (aprobar ? true : tv != "")))
                    if (crear || (!crear && autoriz))
                    {
                        if (artcurrent != "")
                        {
                            disponible = obtener_exist_articulo(artcurrent, crear ? "CD" : "CD_RSV", out errorMessage);
                            if (!string.IsNullOrEmpty(errorMessage)){
                                SForm.Freeze(false);
                                return errorMessage;}

                            if (disponible < totalart)
                            {
                                // Procesar los articulos no disponibles a la hora de transferir
                                lineasnodisp.Add(artcurrent);
                            }
                            else
                            {
                                if (cantlines > 1) // || SMatrix.RowCount ==1)
                                {
                                    doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                    doctransf.Lines.Add();
                                    doctransf.Lines.SetCurrentLine(doctransf.Lines.Count - 1);
                                }
                                linestransf++;
                                doctransf.Lines.ItemCode = artcurrent;
                                doctransf.Lines.ItemDescription = (SMatrix.Columns.Item("articulo").Cells.Item(SMatrix.RowCount).Specific).Value.ToString();
                                doctransf.Lines.Quantity = totalart;
                                doctransf.Lines.FromWarehouseCode = crear ? "CD" : "CD_RSV";
                                doctransf.Lines.WarehouseCode = crear ? "CD_RSV" : "CD";
                            }
                        }
                    }

                    GC.Collect();
                    if (linestransf == 0)
                    {
                        if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                        SForm.Freeze(false);
                        return "Transferir Solicitud: No existen artículos disponibles. " + infonodisp;
                    }
                    result = doctransf.Add();
                    if (result != 0)
                    {
                        if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        errorMessage = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : "";
                        SForm.Freeze(false);
                        return "Transferir Solicitud:" + errorMessage;
                    }
                    
                    

                    if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); }
                    string newkey = B1.Company.GetNewObjectKey();
                    if (newkey != "")
                    {
                        //Actualizar datos de Transferencia en Solicitud
                        string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                        string scom = (crear ? "Reservada" : (aprobar?"Aprobada":"Cancelada")) + " y Transferida: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + " DocNum:" + obtener_DocNum(newkey, out errorMessage) + infonodisp;


                        string sestado = crear ? "R" : "D" ;
                        Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        // Buscando logs actual
                        string SQLQuery = String.Format("SELECT {2} FROM {1} WHERE {0} = '{3}' ",
                                                   Constantes.View.CAB_RVT.U_numDoc,
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
                                                 Constantes.View.CAB_RVT.U_numDoc,   //0
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
                        errorMessage =  cancelar_filas_nodisp(newkey, crear, "" );
                        if (!string.IsNullOrEmpty(errorMessage))  {
                            SForm.Freeze(false);
                            return (errorMessage);}

                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)SForm.Items.Item("1").Specific;
                        btn_crear.Caption = "OK";
                        B1.Application.SetStatusBarMessage("Solicitud " + (crear ? "Reservada" : (aprobar ? "Autorizada" : "Cancelada")) + " y Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Long, false);

                        errorMessage = crear ? cargar_solicitud(sCode, true) : cargar_inicial(); }
                        if (!string.IsNullOrEmpty(errorMessage))  {
                            SForm.Freeze(false);
                            return (errorMessage);}
                     }
                SForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage = "Transferir Solicitud:" + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string revertir(string sCode, string docentry)
        {
            string errorMessage = "";
            bool todoOk = true;
            int result = 0;

            string terror = "";

            try
            {
                SForm.Freeze(true);
                GC.Collect();
                string tr = oDbHeaderDataSource.GetValue("U_idTR", oDbHeaderDataSource.Offset);
                if (string.IsNullOrEmpty(tr))
                {
                    
                    B1.Application.SetStatusBarMessage("Solicitud no Reservada, no es necesario Revertir", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                else
                {
                    if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                    B1.Company.StartTransaction();
                    SAPbobsCOM.StockTransfer doctransf = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                    doctransf.DocDate = DateTime.Today;
                    doctransf.TaxDate = DateTime.Today;
                    string sCli = oDbHeaderDataSource.GetValue("U_codCli", oDbHeaderDataSource.Offset);
                    string snCli = oDbHeaderDataSource.GetValue("U_cliente", oDbHeaderDataSource.Offset);


                    doctransf.CardCode = sCli;
                    doctransf.CardName = snCli;
                    // Serie Primaria
                    doctransf.Series = 27;
                    doctransf.FromWarehouse = "CD_RSV";
                    doctransf.ToWarehouse = "CD";
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
                                    disponible = obtener_exist_articulo(artcurrent, "CD_RSV", out errorMessage);
                                    if (!string.IsNullOrEmpty(errorMessage))
                                    {
                                        SForm.Freeze(false);
                                        return errorMessage;
                                    }


                                    if (disponible >= totalart)
                                    {
                                        if (cantlines > 1)
                                        {
                                            
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
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                SForm.Freeze(false);
                                return errorMessage;
                            }

                            if (disponible >= totalart)
                            {
                                if (cantlines > 1)
                                {
                                    
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
                            terror = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : "";
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
                        if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); }
                        string newkey = B1.Company.GetNewObjectKey();
                        if (newkey != "")
                        {
                            //Actualizar datos de Transferencia en Solicitud
                            string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                            string scom = "Reservada Revertida: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + " DocNum:" + obtener_DocNum(newkey, out errorMessage) + infonodisp;
                            if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }

                            Recordset oRecordSet = (SAPbobsCOM.Recordset)B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            // Buscando logs actual
                            string SQLQuery = String.Format("SELECT {2} FROM {1} WHERE {0} = '{3}' ",
                                                       Constantes.View.CAB_RVT.U_numDoc,
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
                                                     Constantes.View.CAB_RVT.U_numDoc,   //0
                                                     Constantes.View.CAB_RVT.CAB_RV,    //1
                                                     Constantes.View.CAB_RVT.U_logs,    //2
                                                     sCode,                             //3
                                                     scom);                              //4
                            oRecordSet.DoQuery(SQLQuery);
                            B1.Application.SetStatusBarMessage("Solicitud Reservada Revertida Transferida con éxito", SAPbouiCOM.BoMessageTime.bmt_Long, false);

                         }
                    }
                    else
                    {
                        if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        errorMessage = "Revertir Solicitud: Error Transfiriendo Solicitud Reservada Revertida " + terror;


                        //Actualizar logs en Solicitud
                        string infonodisp = lineasnodisp != null && lineasnodisp.Count > 0 ? ", No disp:" + string.Join("-", lineasnodisp) : "";
                        string slog = "Error:No pudo ser Revertida por no tener disponibilidad: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + infonodisp;
                        string scom = "Solicitud sin disponibilidad al intentar Revertir: " + DateTime.Now.Date.ToString("dd/MM/yyyy");

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
                    }
                }


                // Transferir la actualizada
                errorMessage = transferir(true, false);

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    SForm.Freeze(false);
                    return errorMessage;
                }


                SForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage = "Revertir Solicitud: " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private bool encontrar_formulario(out string errorMessage)
        {
            bool encontrado = false;
            errorMessage = "";
            try
            {
                for (int i = 0; i < B1.Application.Forms.Count && !encontrado; i++)
                {
                    encontrado = (B1.Application.Forms.Item(i).UniqueID == fa.ThisSapApiForm.Form.UniqueID);
                }
            }

            catch (Exception ex)
            {
                errorMessage = "Buscando Formulario: " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            } 
            return encontrado;
        }

        private int cantFilas_clipboard(out string errorMessage)
        {
            int filas = 0;
            errorMessage = "";
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
                filas = 0;
                errorMessage =  "Gestionando Clipboard " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }

            return filas;
        }

        private string insertar_lineas_necesarias()
        {
            string errorMessage = "";
            try
            {
                string abuscar = txt_numoc.Value.ToString();

                if (registrar && !ya_Procesada(abuscar, out errorMessage))
                {
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }

                    SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                    mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                    btn_crear.Caption = SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE ? "Actualizar" : btn_crear.Caption;
                    if (SForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    int filasnuevas = cantFilas_clipboard(out errorMessage) - (mtx.RowCount - rowsel + 1);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }

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
                    return "Insertar Líneas: No se puede insertar líneas porque ya es una Solicitud Procesada";
                }
            }
            catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage = "Insertar Líneas: " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string GetClipBoardData()
        {
            string errorMessage = "";
            try
            {
                
                string clipboardData = null;
                //Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate()
                    {
                      clipboardData = Clipboard.GetText(TextDataFormat.Text);
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                return clipboardData;
            }

            catch (Exception ex)
            {
                errorMessage = "Gestionar Clipboard: " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }


    }
}

