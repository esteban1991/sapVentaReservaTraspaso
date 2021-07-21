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
    class PantallaRegistroold : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        private string ItemActiveMenu = "";
        private string formActual = "";
       

        public PantallaRegistroold()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Registro.srf"],"ventaRT_Registro" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            ThisSapApiForm.OnAfterItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterItemPressed);
            ThisSapApiForm.OnBeforeItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnBeforeItemPressed);
            ThisSapApiForm.OnAfterValidate += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterValidate);
            this.B1.Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            ThisSapApiForm.OnAfterGotFocus += new _IApplicationEvents_ItemEventEventHandler(OnAfterGotFocus);

            cargarInfoInicial();
        }

        private void OnAfterGotFocus(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            SAPbouiCOM.Grid grid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;
            SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
   
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_GOT_FOCUS:
                        switch (pVal.ItemUID)
                        {
                            case Constantes.View.registro.grid:
                                if (B1.Application.Forms.ActiveForm.Mode ==  SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                            btn_crear.Caption = "Actualizar";
                                            
                                            txt_com.Item.Enabled = true;
                                            txt_numoc.Item.Enabled = false;
                                            grid.Item.Enabled = true;                                            
                                            B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                            BubbleEvent = true;
                                            break;
                                }
                                BubbleEvent = true;
                                break;
                            case Constantes.View.registro.txt_com:
                                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                    btn_crear.Caption = "Actualizar";
                                    grid.Item.Enabled = true;
                                    txt_numoc.Item.Enabled = false;
                                    txt_com.Item.Enabled = true;

                                    B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    BubbleEvent = true;
                                    break;
                                }
                                BubbleEvent = true;
                                break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
            }
        }


        private void ThisSapApiForm_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
         
            try
            {
                if (pVal.EventType == BoEventTypes.et_CLICK)
                {

                    if (pVal.ItemUID == ventaRT.Constantes.View.registro.btn_crear)
                    {

                    }
                }

                
                //    if (pVal.EventType == BoEventTypes.et_CLICK)
            //    {

            //        if (pVal.ItemUID == ContabilizacionDeNominas.Constantes.Views.registro.btn_cancel)
            //        {
            //             B1.Application.Forms.ActiveForm.Close();
            //            BubbleEvent = false;
            //        }
            //    }

               if (FormUID == formActual)
               {
                   //&&pVal.ItemChanged==true
                   if (pVal.BeforeAction == false && pVal.FormMode == 2 )
                   {
                           SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                           SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                           SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                           SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                         //  SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                           SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                           SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                           SAPbouiCOM.Button btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;
                           btn_crear.Item.Enabled = true;
                           //btn_cancel.Item.Enabled = false;
                   }
               }
            //    //if (B1.Application.Forms.ActiveForm.UniqueID == form)
            //    //{
            //    //    switch (pVal.ItemUID)
            //    //    {
            //    //        case "1":
            //    //            break;
            //    //    }
            //    //}

            }
            catch (Exception ex)
            {
                
                throw;
            }

        }

        private void ThisSapApiForm_OnAfterValidate(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
                //SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                switch (pVal.ItemUID)
                {

                   case ventaRT.Constantes.View.registro.grid:
                        if (pVal.ColUID == "codArti" && pVal.ItemChanged == true)
                       {
                           oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 2, true);
                           string ValueChanged = oGrid.DataTable.GetValue("codArti", pVal.Row);
                           if (ValueChanged != "")
                           {
                                BuscarYGuardarArt(ValueChanged, pVal.Row);
                           }

                        }
                        if (pVal.ColUID == "codClie" && pVal.ItemChanged == true)
                        {
                            oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 4, true);
                            string ValueChanged = oGrid.DataTable.GetValue("codClie", pVal.Row);
                            if (ValueChanged != "")
                            {
                                BuscarYGuardarCliente(ValueChanged, pVal.Row);
                            }

                        }

                        break;
                }

            }
            catch (Exception ex)
            {

                B1.Application.SetStatusBarMessage("Error despues de Validar: " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
        }

        private void BuscarYGuardarArt(string ValueChanged, int row)
        {
            try
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;

                String strSQL = String.Format("SELECT {0},{1}  FROM {2} Where contains({0},'%{3}%')",
                                              Constantes.View.oitm.ItemCode,
                                              Constantes.View.oitm.ItemName,
                                              Constantes.View.oitm.OITM,
                                              ValueChanged);
                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsCards.Fields;

                string ItemCode = rsCards.Fields.Item("ItemCode").Value.ToString();
                string ItemName = rsCards.Fields.Item("ItemName").Value.ToString();
                if (ItemCode != "")
                {

                    oGrid.DataTable.SetValue("articulo", row, fields.Item("ItemName").Value.ToString());
                    oGrid.DataTable.SetValue("codArti", row, fields.Item("ItemCode").Value.ToString());
                    oGrid.CommonSetting.SetCellEditable(row + 1, 2, false);
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Buscando Articulo" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }

        }

        private void BuscarYGuardarCliente(string ValueChanged, int row)
        {
            try
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;

                String strSQL = String.Format("SELECT {0},{1}  FROM {2} Where contains({0},'%{3}%')",
                                              Constantes.View.ocrd.CardCode,
                                              Constantes.View.ocrd.CardName,
                                              Constantes.View.ocrd.OCRD,
                                              ValueChanged);
                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsCards.Fields;

                string CardCode = rsCards.Fields.Item("CardCode").Value.ToString();
                string CardName = rsCards.Fields.Item("CardName").Value.ToString();
                if (CardCode != "")
                {

                    oGrid.DataTable.SetValue("cliente", row, fields.Item("CardName").Value.ToString());
                    oGrid.DataTable.SetValue("codClie", row, fields.Item("CardCode").Value.ToString());
                    oGrid.CommonSetting.SetCellEditable(row + 1, 4, false);
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error Buscando Proveedor" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
        }

        private void cargarInfoInicial()
        {

            //String strSQL = "";
            //try
            //{
            //    strSQL = String.Format("DELETE FROM {0} ", Constantes.View.CAB_RVT.CAB_RV);

            //    Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            //    rsCards.DoQuery(strSQL);
            //}
            //catch (Exception ex)
            //{
            //    throw;
            //}
            



                SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
 
                SAPbouiCOM.Grid grid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;

                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    txt_idvend.Value = obtenerVendedor();
                    txt_numoc.Value = (obtenerUltimoID("CA")+1).ToString();
                    txt_estado.Value = "Nueva";
                    txt_com.Value = "";
                    DateTime fc = DateTime.Now.Date;
                    txt_fechac.Value = SSIFramework.Utilidades.GenericFunctions.GetDate(DateTime.Now.ToString("yyyyMMdd"));
                    DateTime fv = fc.AddDays(10);
                    txt_fechav.Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fv.ToString("yyyyMMdd"));
                    btn_crear.Caption = "Crear";
                    grid.DataTable.Rows.Clear();
                    grid.DataTable.Rows.Add();
                    grid.DataTable.SetValue("cant", grid.DataTable.Rows.Count - 1, 1);
                }

                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    btn_crear.Caption = "Actualizar";
                }
                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    
                    txt_fechac.Value = "" ;
                    txt_fechav.Value = "" ;
                    txt_com.Value = "";
                    txt_estado.Value = "";
                    grid.DataTable.Rows.Clear();
                    grid.Item.Enabled = false;
                    txt_numoc.Active = true;
                    txt_numoc.Value = "";
                    txt_numoc.Item.Enabled = true;
                    txt_com.Item.Enabled = false;
                    btn_crear.Caption = "Buscar";
                    
                }
                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    btn_crear.Caption = "OK";
                }
                formActual = B1.Application.Forms.ActiveForm.UniqueID;
                txt_numoc.Active = true;
        }

        DateTime agregoLinea;

        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                BubbleEvent = false;
                if (pVal.BeforeAction == true)
                {
                    SAPbouiCOM.Matrix oMatrix;
                    switch (pVal.MenuUID)
                    {
                        case "1292":   //ADICIONAR LINEA
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.registro.grid:
                                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;
                                    DateTime inicio = DateTime.Now;
                                    TimeSpan duracion = inicio - agregoLinea;
                                    if (agregoLinea.Year == 1 || (duracion.Seconds >= 2 || duracion.Minutes > 1))
                                    {
                                        oGrid.DataTable.Rows.Add();
                                        agregoLinea = DateTime.Now;
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 2, false);
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 4, false);
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 6, false);
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count,7, false);
                                       // oGrid.DataTable.SetValue("cant", oGrid.DataTable.Rows.Count-1, 1); 
                                        oGrid.AutoResizeColumns();
                                    }
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                        case "1293":  //BORRAR LINEA
                            switch (ItemActiveMenu)
                            {
                                //ejemplo con una matrix 
                                case ventaRT.Constantes.View.registro.mtx:
                                    oForm = B1.Application.Forms.ActiveForm;
                                    oMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific);
                                    //SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@UDT");
                                    int nRow = (int)oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    oMatrix.FlushToDataSource();
                                    oMatrix.LoadFromDataSource();
                                    BubbleEvent = false;
                                    break;

                                case ventaRT.Constantes.View.registro.grid:
                                    bool banderita = false;
                                    oForm = B1.Application.Forms.ActiveForm;
                                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;
                                    for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
                                    {
                                        if (oGrid.Rows.IsSelected(i))
                                        {
                                            if (banderita != true)
                                            {
                                                oGrid.DataTable.Rows.Remove(i);
                                                banderita = true;
                                            }
                                        }
                                    }
                                    BubbleEvent = false;
                                    //oGrid.DataTable.Rows.Remove(0);
                                    break;
                            }
                            break;
                        case "1294":
                            BubbleEvent = false;
                            break;
                        case "1282":  //CREAR
                            if (B1.Application.Forms.ActiveForm.UniqueID == formActual)
                            {
                                SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                                SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                                SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                                SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                                SAPbouiCOM.Grid grid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;
                                txt_vend.Value = obtenerVendedor();
                                txt_numoc.Value = (obtenerUltimoID("CA") + 1).ToString();
                                DateTime fc = DateTime.Now.Date;
                                txt_fechac.Value = SSIFramework.Utilidades.GenericFunctions.GetDate(DateTime.Now.ToString("yyyyMMdd"));
                                DateTime fv = fc.AddDays(10);
                                txt_fechav.Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fv.ToString("yyyyMMdd"));
                                btn_crear.Caption = "Crear";
                                //txt_numoc.Item.Enabled = true;
                                txt_numoc.Active = true;
                                txt_com.Value = "";
                                grid.DataTable.Rows.Clear();
                                grid.DataTable.Rows.Add();
                                grid.DataTable.SetValue("cant", grid.DataTable.Rows.Count - 1, 1);
                                txt_estado.Value = "Nueva";
                                B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                                BubbleEvent = false;
                            }
                            break;
                        case "1281":  //    BUSCAR
                            if (B1.Application.Forms.ActiveForm.UniqueID == formActual)
                            {
                                SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                                SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;                                
                                SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                                SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                                SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                                SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                                SAPbouiCOM.Grid grid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;
                                btn_crear.Caption = "Buscar";
                                
                                txt_fechac.Value = "";
                                txt_fechav.Value = "";
                                txt_estado.Value = "";
                                txt_com.Value = "";
                                grid.DataTable.Rows.Clear();
                                grid.Item.Enabled = false;
                                txt_numoc.Item.Enabled = true;
                                txt_numoc.Active = true;
                                txt_numoc.Value = "";
                                txt_com.Item.Enabled = false;
                                B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                
                                
                                BubbleEvent = false;
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void ThisSapApiForm_OnAfterRightClick(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form oForm = B1.Application.Forms.ActiveForm;
                ItemActiveMenu = eventInfo.ItemUID;
                if (eventInfo.ItemUID == ventaRT.Constantes.View.registro.grid)
                {
                    oForm.EnableMenu("1292", true); //Activar Agregar Linea
                    oForm.EnableMenu("1293", true); //Activar Borrar Linea 
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void ThisSapApiForm_OnBeforeItemPressed(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == Constantes.View.registro.btn_crear)
                {
                    switch (B1.Application.Forms.ActiveForm.Mode)
                    {
                          case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                            {
                                Actualizar_Solicitud();

                                BubbleEvent = true;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                string err = ex.Message;
                throw;
            }

        }

        private void ThisSapApiForm_OnAfterItemPressed(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            SAPbouiCOM.Grid grid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.grid).Specific;
            SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            SAPbouiCOM.Button btn_Crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;

            try
            {
                if (pVal.EventType==BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == Constantes.View.registro.btn_crear)
                {
                    switch (B1.Application.Forms.ActiveForm.Mode)
                    {
                        case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                            {
                                Buscar_Cargar_Solicitud();
                                grid.Item.Enabled = true;
                                txt_com.Item.Enabled = true;
                                txt_com.Active = true;
                                txt_numoc.Item.Enabled = false;
                                BubbleEvent = true;
                                break;
                            }
                        case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                            {
                                Insertar_Solicitud();
                                BubbleEvent = true;
                                break;
                            }
                        case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                            {
                                Actualizar_Solicitud();

                                BubbleEvent = true;
                                break;
                            }
                    }
                }
              }
            catch (Exception ex)
            {
                string err = ex.Message;
                throw;
            }

        }

        private void ActualizarCabecerayCerrar()
        {
            SAPbobsCOM.UserTable oUserTableCa;

            SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
            SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
            SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
            SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
            SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
            SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;                                    

            SAPbouiCOM.Button btn_Crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            SAPbouiCOM.Button btn_Cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_cancel).Specific;


            oUserTableCa = B1.Company.UserTables.Item("CAB_RV");

            oUserTableCa.GetByKey(txt_numoc.Value);
            oUserTableCa.UserFields.Fields.Item("U_comment").Value = txt_com.Value.ToString();
            int i = oUserTableCa.Update();
            if (i != 0)
            {
                B1.Application.SetStatusBarMessage("Error Actualizando Solicitud" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            else
            {
                B1.Application.SetStatusBarMessage("Solicitud Actualizada con Exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                //oForm = B1.Application.Forms.Item("edm");
            }
        }

        private void validacionesDespuesDeGuardar(List<clases.detalle_registro> Lineas)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;

            if (Lineas.Count != 0)
            {
                List<ventaRT.clases.detalle_registro> DetalleBD = new List<ventaRT.clases.detalle_registro>();
                List<ventaRT.clases.detalle_registro> lineasaAgregar = new List<ventaRT.clases.detalle_registro>();
                DetalleBD = ObtenerDetalle(txt_numoc.Value);

                if (Lineas.Count == DetalleBD.Count)
                {

                    insertarLineas("", true);
                }
                //sino no contiene las mismas lineas
                else
                {

                    bool noExiste = true;
                    foreach (var item in Lineas)
                    {


                        foreach (var itemDB in DetalleBD)
                        {
                            if (item.numOC == itemDB.numOC)
                            {
                                noExiste = false;

                                //actualizar el articulo
                                ActualizarLineaUnaXUna(item, itemDB.code);
                                break;
                            }
                        }

                        if (noExiste == true)
                        {
                            //agregarArticulo
                            AñadirLineaLineaUnaXUna(item);
                        }
                    }


                    //para ver si hay que eliminar una linea
                    foreach (var item in DetalleBD)
                    {
                        bool ExisteAun = true;
                        foreach (var itemGrid in Lineas)
                        {
                            if (item.numOC == itemGrid.numOC)
                            {
                                ExisteAun = true;
                                break;
                            }
                            else { ExisteAun = false; }
                        }
                        if (ExisteAun == false)
                        {
                            //eliminamos el articulo
                            elimininarArticulo(item.numOC);

                        }
                    }


                }
            }


        }
 
        private void elimininarArticulo(string code)
        {
            SAPbobsCOM.UserTable oUserTableDE;
            oUserTableDE = B1.Company.UserTables.Item("DET_RV");    
            oUserTableDE.GetByKey(code);

            int i = oUserTableDE.Remove();


            if (i != 0)
            {
                B1.Application.SetStatusBarMessage("Error al eliminar : " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
            else
            {
                B1.Application.SetStatusBarMessage("Edición Eliminada", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
           


            }
        }

        private bool AñadirLineaLineaUnaXUna(clases.detalle_registro item)
        {
            bool isOK = true;
            SAPbobsCOM.UserTable oUserTableDE;
            oUserTableDE = B1.Company.UserTables.Item("DET_RV");
            int IDnEXT = obtenerUltimoID("DE");
            IDnEXT++;
            oUserTableDE.Code = IDnEXT.ToString();
            oUserTableDE.Name = IDnEXT.ToString();

            oUserTableDE.UserFields.Fields.Item("U_numOC").Value = item.numOC;
            oUserTableDE.UserFields.Fields.Item("U_codArti").Value = item.codArti;
            oUserTableDE.UserFields.Fields.Item("U_codClie").Value = item.codClie;
            oUserTableDE.UserFields.Fields.Item("U_cant").Value = item.cant;
            oUserTableDE.UserFields.Fields.Item("U_estado").Value = item.estado;
            oUserTableDE.UserFields.Fields.Item("U_idTV").Value = item.idTV;

            if (ValidarDetalle(item.numOC, item.codArti, item.codClie, oUserTableDE.Code))
            {


                int d = oUserTableDE.Add();
                if (d != 0)
                {
                    isOK = false;
                    B1.Application.SetStatusBarMessage("Error, solo se creo el documento cabecera de la oc:  " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
                else
                {
                    B1.Application.SetStatusBarMessage("Exito en la inserción Detalle", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
            }
            else
            {
                isOK = false;
                B1.Application.SetStatusBarMessage("Error, Datos Repetidos ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            return isOK;
        }

        private bool ActualizarLineaUnaXUna(clases.detalle_registro item, string code)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            bool todoOk = true;

            SAPbobsCOM.UserTable oUserTableDE;
            oUserTableDE = B1.Company.UserTables.Item("DET_RV");

            oUserTableDE.GetByKey(code);
            oUserTableDE.UserFields.Fields.Item("U_numOC").Value = item.numOC;
            oUserTableDE.UserFields.Fields.Item("U_codArti").Value = item.codArti;
            oUserTableDE.UserFields.Fields.Item("U_codClie").Value = item.codClie;
            oUserTableDE.UserFields.Fields.Item("U_cant").Value = item.cant.ToString();
            oUserTableDE.UserFields.Fields.Item("U_estado").Value = item.estado;
            oUserTableDE.UserFields.Fields.Item("U_idTV").Value = item.idTV;
            //if (ValidarDetalle(item.numOC, item.codArti, item.codClie, oUserTableDE.Code))
            //{
                int d = oUserTableDE.Update();
                if (d != 0)
                {
                    B1.Application.SetStatusBarMessage("Error, no se cerro el documento  " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    todoOk = false;
                }
                else
                {
                    B1.Application.SetStatusBarMessage("Exito en la actualización Detalle", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    todoOk = true;
                }
            //}
            //else
            //{
            //    B1.Application.SetStatusBarMessage("Error, Datos Repetidos" , SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            //    todoOk = false;
            //}

            return todoOk;
        }


        private bool validacionesAntesDeGuardar()
        {
           bool todoOk = true;

            if (todoOk != false)
            {
                todoOk = verificarLineasParaGuardar();
            }
            return todoOk;
        }

        private bool verificarLineasParaGuardar()
        {
            bool todoOk = true;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
            if (oGrid.DataTable.Rows.Count != 0)
            {
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    if (todoOk == false)
                    {
                        break;
                    }
                    // fila completa --> cuando se revise o autorice
                    //for (int x = 0; x < oGrid.DataTable.Columns.Count; x++)
                    //{

                    //    string valor = Convert.ToString(oGrid.DataTable.GetValue(x, i));
                    //    if (valor == "")
                    //    {
                    //        todoOk = false;
                    //        break;

                    //    }
                    //}
                    // fila para solicitar nada mas cod cart, codClie y cant
                    for (int x = 0; x < 4; x++)
                    {
                        if ((x == 0) || (x == 2) || (x == 4))
                        {
                            string valor = Convert.ToString(oGrid.DataTable.GetValue(x, i));
                            if (valor == "")
                            {
                                todoOk = false;
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                todoOk = false;
            }



            return todoOk;
        }

        private void LlenarPantalla(string code)
        {
            SAPbouiCOM.EditText txt_vend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
            SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
            SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
            SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
            SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
            SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;

            ventaRT.clases.cabecera_registro cabecera = new ventaRT.clases.cabecera_registro();
            cabecera = ObtenerCabecera(code);

            txt_fechac.Value = cabecera.fechaC;
            txt_fechav.Value = cabecera.fechaV;
            txt_estado.Value = cabecera.estado;
            txt_com.Value = cabecera.comment;
            txt_aut.Value = cabecera.idAut;
            txt_idtr.Value = cabecera.idTR;
            txt_idtv.Value = cabecera.idTV;
     

            //B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

        }

        private List<ventaRT.clases.detalle_registro> ObtenerDetalle(string docnum)
        {


            List<ventaRT.clases.detalle_registro> Lineas = new List<ventaRT.clases.detalle_registro>();

            String strSQL = String.Format("SELECT {0}, T1.{9} articulo, {1}, T2.{11} cliente, {2},{3},{4}, {5}, {14} " +
                                           " FROM {6} T0 INNER JOIN {12} T1 ON T0.{0} = T1.{8} INNER JOIN {13} T2 ON T0.{1} = T2.{10}" +
                                             " Where {5}='{7}'",
                                       Constantes.View.DET_RVT.U_codArt, //0
                                       Constantes.View.DET_RVT.U_codCli, //1
                                       Constantes.View.DET_RVT.U_cant, //2
                                       Constantes.View.DET_RVT.U_estado, //3
                                       Constantes.View.DET_RVT.U_idTV,//4
                                       Constantes.View.DET_RVT.U_numOC,//5
                                       Constantes.View.DET_RVT.DET_RV,//6
                                       docnum, //7
                                       Constantes.View.oitm.ItemCode,  //8
                                       Constantes.View.oitm.ItemName,  //9
                                       Constantes.View.ocrd.CardCode, //10
                                       Constantes.View.ocrd.CardName, //11
                                       Constantes.View.oitm.OITM,  //12
                                       Constantes.View.ocrd.OCRD, //13
                                       Constantes.View.DET_RVT.Code);  //14

  
            Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsCards.DoQuery(strSQL);

            if (rsCards.RecordCount != 0)
            {
                rsCards.MoveFirst();

                for (int i = 1; !rsCards.EoF; i++)
                {
                    ventaRT.clases.detalle_registro detalles = new ventaRT.clases.detalle_registro();
                    SAPbobsCOM.Fields fields = rsCards.Fields;
                    detalles.numOC= fields.Item(Constantes.View.DET_RVTabla.U_numOC).Value.ToString();
                    detalles.codArti = fields.Item(Constantes.View.DET_RVTabla.U_codArt).Value.ToString();
                    detalles.articulo = fields.Item(Constantes.View.DET_RVTabla.articulo).Value.ToString();
                    detalles.codClie = fields.Item(Constantes.View.DET_RVTabla.cliente).Value.ToString();
                    detalles.cant = fields.Item(Constantes.View.DET_RVTabla.U_cant).Value;
                    detalles.estado = fields.Item(Constantes.View.DET_RVTabla.U_estado).Value.ToString();
                    detalles.idTV = fields.Item(Constantes.View.DET_RVTabla.U_idTV).Value.ToString();
                    detalles.code = fields.Item("Code").Value;

                    Lineas.Add(detalles);
                    rsCards.MoveNext();
                }
            }

            return Lineas;
        }

        private clases.cabecera_registro ObtenerCabecera(string code)
        {
            ventaRT.clases.cabecera_registro cabecera = new ventaRT.clases.cabecera_registro();
            String strSQL = String.Format("SELECT {0},{1},{2},{3},{4},{5},{6},{7},{8}  FROM {9} Where {10}='{11}'",
                                         Constantes.View.CAB_RVT.U_numOC,
                                         Constantes.View.CAB_RVT.U_fechaC,
                                         Constantes.View.CAB_RVT.U_fechaV,
                                         Constantes.View.CAB_RVT.U_estado,
                                         Constantes.View.CAB_RVT.U_idVend,
                                         Constantes.View.CAB_RVT.U_idAut,
                                         Constantes.View.CAB_RVT.U_idTR,
                                         Constantes.View.CAB_RVT.U_idTV,
                                         Constantes.View.CAB_RVT.U_comment,
                                         Constantes.View.CAB_RVT.CAB_RV,
                                         Constantes.View.CAB_RVT.Code,
                                         code);
            Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsCards.DoQuery(strSQL);
            if (rsCards.RecordCount != 0)
            {
                SAPbobsCOM.Fields fields = rsCards.Fields;
                cabecera.numOC = fields.Item(Constantes.View.CAB_RVTabla.U_numOC).Value.ToString();
                cabecera.fechaC = fields.Item(Constantes.View.CAB_RVTabla.U_fechaC).Value.ToString();
                cabecera.fechaV = fields.Item(Constantes.View.CAB_RVTabla.U_fechaV).Value.ToString();
                cabecera.idVend = fields.Item(Constantes.View.CAB_RVTabla.U_idVend).Value;
                cabecera.idAut = fields.Item(Constantes.View.CAB_RVTabla.U_idAut).Value.ToString();
                cabecera.idTV = fields.Item(Constantes.View.CAB_RVTabla.U_idTV).Value.ToString();
                cabecera.idTR = fields.Item(Constantes.View.CAB_RVTabla.U_idTR).Value.ToString();
                cabecera.estado = fields.Item(Constantes.View.CAB_RVTabla.U_estado).Value.ToString();
                cabecera.comment = fields.Item(Constantes.View.CAB_RVTabla.U_comment).Value.ToString();
            }

            return cabecera;

        }

        private bool insertarLineas(string DocNum, bool tipoUpdate = false)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            bool todoOk = true;

            int IDnEXT = obtenerUltimoID("Grid");
            if (tipoUpdate == true)
            {
                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    //IDnEXT++;
                    SAPbobsCOM.UserTable oUserTableDE;
                    oUserTableDE = B1.Company.UserTables.Item("DET_RV");

                    //TENGO QUE VER SI las lineas  el grid actual se diferencia con la anterior
                    //que pasa si eliminaron una linea, o modificar un codigo de una linea?
                    //se tiene que hacer una funcion preguntando si el articulo esta en la lista, si lo esta, se tiene que obtener el codigo que tiene
                    //sino se tiene se tiene que agregar una una lista que seran nuevas lineas pero
                    //que pasa si hay eenos lineas, entonces tambien debo buscar cual es el elmento eliminado y que elimine esas lineas en especifico
                    //esto demorara dos dias mas como minimo solo eso,  quedan los bugs, mas modo tipo cerrado, validar antes de guardar y por ultimo al ventana nueva
                    //que demorara como dos días mas  tambien , ent total el proyecto estaria  mejor de los casos el 29 para  el 2 de diciembre es el realista , ya que no trabajo ni el 30 ni el 1

                    //string CodeOb = getCodePerCodNUM(oGrid.DataTable.GetValue("code", i));
                    string CodeOb = oGrid.DataTable.GetValue("code", i);
                    oUserTableDE.GetByKey(CodeOb);
                    oUserTableDE.UserFields.Fields.Item("U_codArti").Value = oGrid.DataTable.GetValue("codArti", i);
                    oUserTableDE.UserFields.Fields.Item("U_codClie").Value = oGrid.DataTable.GetValue("codClie", i);
                    oUserTableDE.UserFields.Fields.Item("U_cant").Value = oGrid.DataTable.GetValue("cant", i);
                    oUserTableDE.UserFields.Fields.Item("U_estado").Value = oGrid.DataTable.GetValue("estado", i);
                    oUserTableDE.UserFields.Fields.Item("U_idTV").Value = oGrid.DataTable.GetValue("idTV", i);
                    oUserTableDE.UserFields.Fields.Item("U_numOC").Value = txt_numoc.Value;
  //                  if (ValidarDetalle(txt_numoc.Value.ToString(), oGrid.DataTable.GetValue("codArti", i),
  //                      oGrid.DataTable.GetValue("codClie", i),CodeOb))
  //                  {

                        int d = oUserTableDE.Update();
                        if (d != 0)
                        {
                            B1.Application.SetStatusBarMessage("Error, no se cerro el documento  " + DocNum + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            todoOk = false;
                            break;
                        }
                        else
                        {

                            B1.Application.SetStatusBarMessage("Exito en la actualización Detalle", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            todoOk = true;
                        }
                    //}
                    //else
                    //{

                    //    B1.Application.SetStatusBarMessage("Error, Datos Repetidos ",  SAPbouiCOM.BoMessageTime.bmt_Medium, false);                        
                    //    todoOk = false;
                    //    break;
                    //}


                }
            }
            else
            {
                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    IDnEXT++;
                    SAPbobsCOM.UserTable oUserTableDE;
                    oUserTableDE = B1.Company.UserTables.Item("DET_RV");
                    oUserTableDE.Code = IDnEXT.ToString();
                    oUserTableDE.Name = IDnEXT.ToString();
                    oUserTableDE.UserFields.Fields.Item("U_codArti").Value = oGrid.DataTable.GetValue("U_codArti", i);
                    oUserTableDE.UserFields.Fields.Item("U_codClie").Value = oGrid.DataTable.GetValue("U_codClie", i);
                    oUserTableDE.UserFields.Fields.Item("U_cant").Value = oGrid.DataTable.GetValue("U_cant", i);
                    oUserTableDE.UserFields.Fields.Item("U_estado").Value = oGrid.DataTable.GetValue("U_estado", i);
                    oUserTableDE.UserFields.Fields.Item("U_idTV").Value = oGrid.DataTable.GetValue("U_idTV", i);
                    oUserTableDE.UserFields.Fields.Item("U_numOC").Value = DocNum;
                    //if (ValidarDetalle(txt_numoc.Value.ToString(), oGrid.DataTable.GetValue("codArti", i),
                    //    oGrid.DataTable.GetValue("codClie", i), "")) 
                    //{
                        int d = oUserTableDE.Add();
                        if (d != 0)
                        {
                            B1.Application.SetStatusBarMessage("Error, solo se creo el documento cabecera de la oc:  " + DocNum + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            todoOk = false;
                            break;
                        }
                        else
                        {
                            B1.Application.SetStatusBarMessage("Exito en la inserción Detalle", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            todoOk = true;
                        }
                    //}
                    //else
                    //{
                    //    B1.Application.SetStatusBarMessage("Error,Datos Repetidos ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    //    todoOk = false;
                    //    break;
                    //}

                }
            }
            return todoOk;
        }

        private string getCodePerCodNUM(string p)
        {
            String strSQL = String.Format("SELECT {0} FROM {1} where {2}='{3}'",
                                    Constantes.Views.DET_REC_IMPT.Code,
                                    Constantes.Views.DET_REC_IMPT.DET_REC_IMP,
                                    Constantes.Views.DET_REC_IMPT.U_Cod_Articulo,
                                    p
                                    );

            Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsCards.DoQuery(strSQL);

            string Code = rsCards.Fields.Item("Code").Value.ToString();

            return Code;
        }

        private List<ventaRT.clases.detalle_registro> obtenerLineasParaGuardar(string DocNum)
        {
            List<ventaRT.clases.detalle_registro> Lineas = new List<ventaRT.clases.detalle_registro>();
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;

            //int IDnEXT = obtenerUltimoID("Grid");
            for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
            {

                ventaRT.clases.detalle_registro detalles = new ventaRT.clases.detalle_registro();
                detalles.code = oGrid.DataTable.GetValue("code", i);
                detalles.codArti = oGrid.DataTable.GetValue("codArti", i);
                detalles.codClie = oGrid.DataTable.GetValue("codClie", i);
                detalles.cant= oGrid.DataTable.GetValue("cant", i);
                detalles.estado = oGrid.DataTable.GetValue("estado", i);
                detalles.idTV = oGrid.DataTable.GetValue("idTV", i);
                detalles.numOC = DocNum;
                Lineas.Add(detalles);
            }

            return Lineas;

        }

        private bool ActualizarLineas(string DocNum)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
            bool todoOk = true;

            for (int i = 0; i < oGrid.Rows.Count; i++)
            {
                SAPbobsCOM.UserTable oUserTableDE;
                oUserTableDE = B1.Company.UserTables.Item("DET_RV");
                string code = oGrid.DataTable.GetValue("code", i);
                oUserTableDE.GetByKey(code);
                oUserTableDE.UserFields.Fields.Item("U_codArti").Value = oGrid.DataTable.GetValue("codArti", i);
                oUserTableDE.UserFields.Fields.Item("U_codClie").Value = oGrid.DataTable.GetValue("codClie", i);
                oUserTableDE.UserFields.Fields.Item("U_cant").Value = oGrid.DataTable.GetValue("cant", i);
                oUserTableDE.UserFields.Fields.Item("U_estado").Value = oGrid.DataTable.GetValue("estado", i);
                oUserTableDE.UserFields.Fields.Item("U_idTV").Value = oGrid.DataTable.GetValue("idTV", i);
                oUserTableDE.UserFields.Fields.Item("U_numOC").Value = DocNum;
                if (ValidarDetalle(DocNum, oGrid.DataTable.GetValue("codArti", i),
                    oGrid.DataTable.GetValue("codClie", i), code ))
                {
                    int d = oUserTableDE.Update();
                    if (d != 0)
                    {
                        B1.Application.SetStatusBarMessage("Error, no se guardo la actualizacion despues del guardado de articulo:  " + DocNum + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        todoOk = false;
                        break;
                    }
                    else
                    {
                        B1.Application.SetStatusBarMessage("Exito en la la actulizacion tabla Detalle Rec. Imp. ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        todoOk = true;
                    }
                }
                else
                {
                    B1.Application.SetStatusBarMessage("Error, Datos Repetidos  ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    todoOk = false;
                    break;
                }

            }
            return todoOk;
        }

        private bool validaciones(string tipo)
        {
            bool ISok = true;
            if (tipo == "Detalles")
            {

            }
            else
            {
                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                ISok = txt_numoc.Value.ToString() == "" ? false : true;
            }


            return ISok;
        }

        private void llenarGrid(Recordset rsCards2)
        {
            try
            {
                SAPbouiCOM.Form oForm = B1.Application.Forms.ActiveForm;
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.View.registro.grid).Specific;
                if (oGrid.DataTable.Rows.Count != 0)
                {
                    oGrid.DataTable.Rows.Clear();
                }

                SAPbobsCOM.Fields fields = rsCards2.Fields;

                //int lastRowIndex = DT_GRID.Rows.Count ;              
                rsCards2.MoveFirst();
                for (int i = 1; !rsCards2.EoF; i++)
                {
                    oGrid.DataTable.Rows.Add();

                    oGrid.DataTable.SetValue("codArti", oGrid.DataTable.Rows.Count - 1, fields.Item("U_codArti").Value.ToString());
                    oGrid.DataTable.SetValue("articulo", oGrid.DataTable.Rows.Count - 1, fields.Item("articulo").Value.ToString());
                    oGrid.DataTable.SetValue("codClie", oGrid.DataTable.Rows.Count - 1, fields.Item("U_codClie").Value.ToString());
                    oGrid.DataTable.SetValue("cliente", oGrid.DataTable.Rows.Count - 1, fields.Item("cliente").Value.ToString());
                    oGrid.DataTable.SetValue("cant", oGrid.DataTable.Rows.Count - 1, fields.Item("U_cant").Value.ToString());
                    oGrid.DataTable.SetValue("estado", oGrid.DataTable.Rows.Count - 1, fields.Item("U_estado").Value.ToString());
                    oGrid.DataTable.SetValue("idTV", oGrid.DataTable.Rows.Count - 1, fields.Item("U_idTV").Value.ToString());
                    oGrid.DataTable.SetValue("code", oGrid.DataTable.Rows.Count - 1, fields.Item("Code").Value.ToString());
                    oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 2, false);
                    oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 4, false);
                    oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 6, false);
                    oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 7, false);
                    oGrid.AutoResizeColumns();
                    rsCards2.MoveNext();

                }

            }
            catch (Exception ex)
            {

                throw;
            }

        }

        private void Insertar_Solicitud()
        {
            try { 
                SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;

                if (txt_numoc.Value.ToString().Length > 0)
                {
                    SAPbobsCOM.UserTable oUserTableCa;
                    oUserTableCa = B1.Company.UserTables.Item("CAB_RV");

                    int IDnEXT = obtenerUltimoID("CA");
                    IDnEXT = IDnEXT + 1;
                    oUserTableCa.Code = IDnEXT.ToString();
                    oUserTableCa.Name = IDnEXT.ToString();
                    oUserTableCa.UserFields.Fields.Item("U_numOC").Value = txt_numoc.Value.ToString();
                    DateTime fc = DateTime.Now.Date;
                    DateTime fv = fc.AddDays(10);
                    string fechaHoy = fc.ToString("yyyyMMdd");
                    oUserTableCa.UserFields.Fields.Item("U_fechaC").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fechaHoy);
                    string fechhoydt = fv.ToString("yyyyMMdd");
                    oUserTableCa.UserFields.Fields.Item("U_fechaV").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fechhoydt);
                     oUserTableCa.UserFields.Fields.Item("U_estado").Value = "N";
                    oUserTableCa.UserFields.Fields.Item("U_idTR").Value = "0";
                    oUserTableCa.UserFields.Fields.Item("U_idTV").Value = "0";
                    oUserTableCa.UserFields.Fields.Item("U_idAut").Value = "0";
                    oUserTableCa.UserFields.Fields.Item("U_comment").Value = txt_com.Value.ToString();
                    oUserTableCa.UserFields.Fields.Item("U_idVend").Value = obtenerIdVendedor();

                    int i = oUserTableCa.Add();

                    if (i != 0)
                    {
                        B1.Application.SetStatusBarMessage("Error" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);

                    }
                    else
                    {
                        //Guardando lineas productos
                        //aqui iran las validaciones
                        bool isOK = validacionesAntesDeGuardar();
                        if (isOK)
                        {
                                List<ventaRT.clases.detalle_registro> Lineas = new List<ventaRT.clases.detalle_registro>();
                                //verificar que traiga los datos act actualizados
                                Lineas = obtenerLineasParaGuardar(txt_numoc.Value);

                                ////aqui se guardara el articulo
                                //SAPbobsCOM.Items detalle = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                //if (Lineas.Count != 0)
                                //{
                                //    int fila = 0;

                                //    foreach (var item in Lineas)
                                //    {
                                //        //detalle.GetByKey(item.code);
                                //        detalle.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_codArti).Value = item.codArti;
                                //        detalle.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_codClie).Value = item.codClie;
                                //        detalle.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_cant).Value = item.cant;
                                //        detalle.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_numOC).Value = txt_numoc.Value.ToString();;
                                //        detalle.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_estado).Value = item.estado;
                                //        detalle.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_idTV).Value = item.idTV;

                                //        int d = detalle.Add();
                                //        if (d != 0)
                                //        {
                                //            B1.Application.SetStatusBarMessage("Error al guardar productos: " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                //        }
                                //        else
                                //        {

                                            validacionesDespuesDeGuardar(Lineas);
                                            B1.Application.SetStatusBarMessage("Solicitud guardada con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            
                            //              ActualizarCabecerayCerrar();

                             //               B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                              //              B1.Application.Forms.ActiveForm.Close();
                                    //    }
                                    //    fila++;
                                    //}
                                //}
                                //detalle.GetByKey(txt_numoc.Value);
                                //aqui se actualizara de nuevo la tabla con los datos despues de guardarse con exito
                            }
                            else
                            {
                                B1.Application.SetStatusBarMessage("Error, Por favor revisar los datos que faltan.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                            }
                        B1.Application.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    }
                }
                 
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error insertando solicitud: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
                      

  
        }

        private void Actualizar_Solicitud()
        {
            try
            {
                SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                SAPbouiCOM.EditText txt_aut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
                SAPbouiCOM.EditText txt_idtr = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtr).Specific;
                SAPbouiCOM.EditText txt_idtv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idtv).Specific;
                SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;

                if (txt_numoc.Value.ToString().Length > 0)
                {

                    //Guardando lineas productos
                    bool isOK = validacionesAntesDeGuardar();
                    if (isOK)
                    {
                         List<ventaRT.clases.detalle_registro> Lineas = new List<ventaRT.clases.detalle_registro>();
                         //verificar que traiga los datos act actualizados
                          Lineas = obtenerLineasParaGuardar(txt_numoc.Value);
                          //aqui se guardara el articulo
                          SAPbobsCOM.UserTable oUserTableDE = B1.Company.UserTables.Item("DET_RV");
                            if (Lineas.Count != 0)
                            {
                                int fila = 0;

                                foreach (var item in Lineas)
                                {
                                    if (oUserTableDE.GetByKey(item.code))
                                    {
                                        // existente par actualizar
                                        oUserTableDE.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_codArt).Value = item.codArti;
                                        oUserTableDE.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_codCli).Value = item.codClie;
                                        oUserTableDE.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_cant).Value =  item.cant.ToString();
                                        oUserTableDE.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_numOC).Value = txt_numoc.Value.ToString();
                                        oUserTableDE.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_estado).Value = item.estado;
                                        oUserTableDE.UserFields.Fields.Item(Constantes.View.DET_RVTabla.U_idTV).Value = item.idTV;
                                        if (ValidarDetalle(txt_numoc.Value.ToString(), item.codArti, item.codClie, item.code)) 
                                        {
                                            int d = oUserTableDE.Update();
                                            isOK = (d == 0); 
                                        }
                                        else
                                        {
                                            isOK = false;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        // no existente, adicionar
                                        isOK = AñadirLineaLineaUnaXUna(item);
                                    }
                                    
                                    fila++;
                                }
                                if (isOK)
                                {
                                    validacionesDespuesDeGuardar(Lineas);
                                    ActualizarCabecerayCerrar();
                                    B1.Application.SetStatusBarMessage("Solicitud guardada con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                }
                                else
                                {
                                    isOK = false;
                                    B1.Application.SetStatusBarMessage("Error al actualizar productos: " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                }
                            }
                            //detalle.GetByKey(txt_numoc.Value);
                            //aqui se actualizara de nuevo la tabla con los datos despues de guardarse con exito
                        }
                        else
                        {
                            isOK = false;
                            B1.Application.SetStatusBarMessage("Error, Por favor revisar los datos que faltan.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        }

                        if (isOK)
                        {
                            B1.Application.SetStatusBarMessage("Exito en la actualizacion", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        }
                    }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error insertando solicitud: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }



        }

        private bool ValidarDetalle(String numdoc,String codArti,String prov,String code)
        {
            //String strSQL = "";
            //try
            //{
            //     if (code == "")
            //{


            //    strSQL = String.Format("SELECT {4} " +
            //                                   " FROM {3} " +
            //                                     " Where {2}='{5}' and {0}='{6}' and {1}='{7}'",
            //                               Constantes.View.DET_RVT.U_codArti, //0
            //                               Constantes.View.DET_RVT.U_codClie, //1                                       
            //                               Constantes.View.DET_RVT.U_numOC,//2
            //                               Constantes.View.DET_RVT.DET_RV,//3
            //                              Constantes.View.DET_RVT.Code,//4
            //                               numdoc, //5
            //                               codArti,  //6
            //                               prov);  //7

            //}
            //else
            //{
            //    strSQL = String.Format("SELECT {4} " +
            //                                   " FROM {3} " +
            //                                     " Where {2}='{5}' and {0}='{6}' and {1}='{7}' and {4}<> '{8}' ",
            //                               Constantes.View.DET_RVT.U_codArti, //0
            //                               Constantes.View.DET_RVT.U_codClie, //1                                       
            //                               Constantes.View.DET_RVT.U_numOC,//2
            //                               Constantes.View.DET_RVT.DET_RV,//3
            //                              Constantes.View.DET_RVT.Code,//4
            //                               numdoc, //5
            //                               codArti,  //6
            //                               prov,   //7
            //                               code);  //8
            //}
  
            //Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            //rsCards.DoQuery(strSQL);
            //bool isOK = (rsCards.RecordCount == 0);
            //if (!isOK)
            //{
            //    String terror = "Error, Codigo Articulo (" + codArti + ") y Proveedor (" +
            //             prov + ") Repetidos en la Solicitud " + numdoc;
            //    B1.Application.MessageBox(terror, 1, "Ok", "", "");
            //}

            //return isOK;
            //}
            //catch(Exception ex)
            //{
            //    return false;
            //}
            return true;
                   
        }

        private void Buscar_Cargar_Solicitud()
        {
                SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                if (txt_numoc.Value.ToString() != "")
                {
                    String strSQL = String.Format("SELECT {0},{1},{2},{3},{4},{5},{6},{7},{8},{9} FROM {10} WHERE {0} = '{11}'",
                        Constantes.View.CAB_RVT.U_numOC,
                        Constantes.View.CAB_RVT.U_fechaC,
                        Constantes.View.CAB_RVT.U_fechaV,
                        Constantes.View.CAB_RVT.U_estado,
                        Constantes.View.CAB_RVT.U_idTR,
                        Constantes.View.CAB_RVT.U_idTV,
                        Constantes.View.CAB_RVT.U_estado,
                        Constantes.View.CAB_RVT.U_comment,
                        Constantes.View.CAB_RVT.U_idVend,
                        Constantes.View.CAB_RVT.U_idAut,
                        Constantes.View.CAB_RVT.CAB_RV,
                        txt_numoc.Value.ToString());
                    Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsCards.DoQuery(strSQL);
                    string U_numOC = rsCards.Fields.Item("U_numOC").Value.ToString();

                    if (U_numOC != "")
                    {

                        String strSQL2 = String.Format("SELECT {0}, T1.{9} articulo, {1}, T2.{11} cliente, {2},{3},{4}, {5}, {14}" +
                                                         " FROM {6} T0 INNER JOIN {12} T1 ON T0.{0} = T1.{8} INNER JOIN {13} T2 ON T0.{1} = T2.{10}" +
                                                           " Where {5}='{7}'",
                                                     Constantes.View.DET_RVT.U_codArt, //0
                                                     Constantes.View.DET_RVT.U_codCli, //1
                                                     Constantes.View.DET_RVT.U_cant, //2
                                                     Constantes.View.DET_RVT.U_estado, //3
                                                     Constantes.View.DET_RVT.U_idTV,//4
                                                     Constantes.View.DET_RVT.U_numOC,//5
                                                     Constantes.View.DET_RVT.DET_RV,//6
                                                     U_numOC, //7
                                                     Constantes.View.oitm.ItemCode,  //8
                                                     Constantes.View.oitm.ItemName,  //9
                                                     Constantes.View.ocrd.CardCode, //10
                                                     Constantes.View.ocrd.CardName, //11
                                                     Constantes.View.oitm.OITM,  //12
                                                     Constantes.View.ocrd.OCRD, //13
                                                     Constantes.View.DET_RVT.Code);  //14

                        Recordset rsCards2 = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsCards2.DoQuery(strSQL2);
                        LlenarPantalla(U_numOC);
                        llenarGrid(rsCards2);
                        B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    }
                }
        }

        private void llenarMatrix(Recordset rsCards2, bool primeravez = true)
        {


            SAPbouiCOM.Matrix oMatrix;
            oMatrix = (Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
            oMatrix.FlushToDataSource();
            oMatrix.Clear();
            SAPbouiCOM.UserDataSource num = B1.Application.Forms.ActiveForm.DataSources.UserDataSources.Add("num", BoDataType.dt_SHORT_NUMBER, 3);
            SAPbouiCOM.UserDataSource ItemCode = B1.Application.Forms.ActiveForm.DataSources.UserDataSources.Add("dsEndDt", BoDataType.dt_SHORT_TEXT, 20);
            SAPbouiCOM.UserDataSource Dscription = B1.Application.Forms.ActiveForm.DataSources.UserDataSources.Add("dsEndDt2", BoDataType.dt_LONG_TEXT, 100);
            SAPbouiCOM.UserDataSource Quantity = B1.Application.Forms.ActiveForm.DataSources.UserDataSources.Add("dsEndDt3", BoDataType.dt_LONG_NUMBER, 10);
            SAPbobsCOM.Fields fields = rsCards2.Fields;

            B1.Application.Forms.ActiveForm.Freeze(true);
            oMatrix.Columns.Item("#").DataBind.SetBound(true, "", "num");
            oMatrix.Columns.Item("Col_0").DataBind.SetBound(true, "", "dsEndDt");
            oMatrix.Columns.Item("Col_1").DataBind.SetBound(true, "", "dsEndDt2");
            oMatrix.Columns.Item("Col_2").DataBind.SetBound(true, "", "dsEndDt3");
            rsCards2.MoveFirst();
            for (int i = 1; !rsCards2.EoF; i++)
            {
                oMatrix.AutoResizeColumns();
                int fila = i;
                num.Value = fila.ToString();
                ItemCode.Value = fields.Item("ItemCode").Value.ToString();
                Dscription.Value = fields.Item("Dscription").Value.ToString();
                Quantity.Value = fields.Item("Quantity").Value.ToString();

                oMatrix.AddRow();

                rsCards2.MoveNext();
            }
            B1.Application.Forms.ActiveForm.Freeze(false);
        }

        private string obtenerIdVendedor()
        {
            try
            {
                string usrCurrent = B1.Company.UserName;
                String strSQL = String.Format("SELECT {0}  FROM {2} Where contains({1},'%{3}%')",
                          Constantes.View.ousr.uId,
                          Constantes.View.ousr.uName,
                          Constantes.View.ousr.OUSR,
                          usrCurrent);
                Recordset rsUsers = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsUsers.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsUsers.Fields;
                string User_Id = rsUsers.Fields.Item("USERID").Value.ToString();
                return User_Id;
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Vendedor", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }   

        private string obtenerVendedor()
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
                return usrCurrent + "-" + User_Name;
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error obteniendo Vendedor", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }

        private int obtenerUltimoID(string tipo)
        {
            int CodeNumCA = 0;
            int CodeNumDE = 0;
            if (tipo == "CA")
            {

                String strSQL = String.Format("SELECT  COUNT(*)  FROM {0}",
                                    Constantes.View.CAB_RVT.CAB_RV);

                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);

                string Code = rsCards.Fields.Item("COUNT(*)").Value.ToString();

                //probar cuando la tabla este vacia, osea el primero registro y no haya otro anterior
                if (Code != "")
                {
                    CodeNumCA = Convert.ToInt32(Code);

                }
                return CodeNumCA;
            }
            else
            {

                String strSQL = String.Format("SELECT  COUNT(*)  FROM {0}",
                                    Constantes.View.DET_RVT.DET_RV);

                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);

                string Code = rsCards.Fields.Item("COUNT(*)").Value.ToString();
                if (Code != "")
                {
                    CodeNumDE = Convert.ToInt32(Code);

                }

                return CodeNumDE;
            }




        }

    }
}
