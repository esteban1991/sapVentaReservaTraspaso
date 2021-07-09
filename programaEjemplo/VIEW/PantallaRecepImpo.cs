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
    class PantallaRecepImpo : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        private string ItemActiveMenu = "";
        private string formActual = "";
       

        public PantallaRecepImpo()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.RecepImpo.srf"], "HJ_RecepImpo" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {

    
            ThisSapApiForm.OnAfterItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterItemPressed);          
            this.B1.Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
         
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            ThisSapApiForm.OnAfterValidate += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterValidate);
            //ThisSapApiForm.OnAfterFormLoad += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_LoadAfter);
          
            
            cargarInfoInicial();
            //ThisSapApiForm.OnAfterValidate += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterValidate);
        }

        private void ThisSapApiForm_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                       if (FormUID == formActual)
                       {
                           //&&pVal.ItemChanged==true
                           if (pVal.BeforeAction == false && pVal.FormMode == 2 )
                           {
                                   SAPbouiCOM.EditText docNum = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
                                   SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                                   SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
                                   SAPbouiCOM.EditText cantCarto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
                                   SAPbouiCOM.EditText usrCre = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
                                   SAPbouiCOM.EditText fecha = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
                                   SAPbouiCOM.EditText prove = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
                                   SAPbouiCOM.EditText estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_est).Specific;
                                   SAPbouiCOM.Button btnBuscar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
                                   SAPbouiCOM.Button btn_save = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
                                   SAPbouiCOM.Button btn_oc = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_bucOC).Specific;
                                   SAPbouiCOM.Button btn_cre = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;

                                   btn_save.Item.Enabled = true;
                                   btn_cre.Item.Enabled = false;


             
               
                           }
      

                            }
       


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
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
                SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                switch (pVal.ItemUID)
                {

                    case ventaRT.Constantes.Views.recepImpo.grid:
                        if (pVal.ColUID == "ITEMCODE" && pVal.ItemChanged == true) 
                        {
                            oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 2, true);
                            string ValueChanged= oGrid.DataTable.GetValue("ITEMCODE", pVal.Row);
                            if (ValueChanged != "")
                            {
                                BuscarYGuardarValie(ValueChanged, pVal.Row);
                            }
                            
                        }
                        if (pVal.ColUID == "cart_Pall" && pVal.ItemChanged == true)
                        {
                            //agregar los enable  y disable, tambien cuando se muestren deben estar  disable,
                           // probar y verifar que no de error porque esta siendo  usado
                            oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 21, true);
                            oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 22, true);
                            int ValueChanged = oGrid.DataTable.GetValue("cart_Pall", pVal.Row);
                            string CasePack =  oGrid.DataTable.GetValue("REPack", pVal.Row);
                            int Resultado;
                            if (CasePack == "")
                            {
                                int CasePackint = 0;
                                Resultado = ValueChanged ;
                                txt.Active = true;
                                oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 21, false);
                                oGrid.CommonSetting.SetCellEditable(pVal.Row + 1 + 1, 22, false);

                            }
                            else
                            {

                                int CasePackint = Convert.ToInt32(oGrid.DataTable.GetValue("REPack", pVal.Row));
                                Resultado = ValueChanged * CasePackint;
                                oGrid.DataTable.SetValue("uniPall", pVal.Row, Resultado);
                                txt.Active = true;
                                oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 21, false);
                                oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 22, false);
                            }
                            
                        }
                        if ((pVal.ColUID == "PA.Base" || pVal.ColUID == "PA.Altura"||pVal.ColUID == "PA.Saldo")&& pVal.ItemChanged == true)
                        {
                            int resultado = 0;
                             int pbase = oGrid.DataTable.GetValue("PA.Base", pVal.Row);
                             int altura = oGrid.DataTable.GetValue("PA.Altura", pVal.Row);
                             int saldo = oGrid.DataTable.GetValue("PA.Saldo", pVal.Row);
                             if (pbase == 0 || altura ==0 || saldo == 0) 
                             {
                             }
                             else
                             {
                                 oGrid.CommonSetting.SetCellEditable(pVal.Row+1, 21, true);
                                 oGrid.CommonSetting.SetCellEditable(pVal.Row+1, 22, true);
                                 resultado = (pbase * altura )+ saldo;
                                 oGrid.DataTable.SetValue("cart_Pall", pVal.Row, resultado);
                                 string CasePack = oGrid.DataTable.GetValue("REPack", pVal.Row);                                
                                 if (CasePack != "")
                                 {
                                     int CasePackint = Convert.ToInt32(oGrid.DataTable.GetValue("REPack", pVal.Row));
                                     int unipall = resultado * CasePackint;

                                     oGrid.DataTable.SetValue("uniPall", pVal.Row, unipall);
                                     txt.Active = true;
                                     oGrid.CommonSetting.SetCellEditable(pVal.Row + 1, 21, false);
                                     oGrid.CommonSetting.SetCellEditable(pVal.Row+1, 22, false);
                                 }
                                 else
                                 {
                                     oGrid.DataTable.SetValue("uniPall", pVal.Row, resultado);
                                     txt.Active = true;
                                     oGrid.CommonSetting.SetCellEditable(pVal.Row, 21, false);
                                     oGrid.CommonSetting.SetCellEditable(pVal.Row, 22, false);
                                 }
                                
                             }
                        }
                        if (pVal.ColUID == "REPack" && pVal.ItemChanged == true)
                        {
                            string ValueChanged = oGrid.DataTable.GetValue("REPack", pVal.Row);
                            int cart_Pall = oGrid.DataTable.GetValue("cart_Pall", pVal.Row);
                            int Resultado;
                            if (cart_Pall == 0)
                            {

                                int ValueChangedint = Convert.ToInt32(oGrid.DataTable.GetValue("REPack", pVal.Row));
                                Resultado = ValueChangedint;
                                oGrid.DataTable.SetValue("uniPall", pVal.Row, Resultado);

                            }
                            else
                            {
                                int ValueChangedint = Convert.ToInt32(oGrid.DataTable.GetValue("REPack", pVal.Row));
                                Resultado = ValueChangedint * cart_Pall;
                                oGrid.DataTable.SetValue("uniPall", pVal.Row, Resultado);
                            }

                        }
                        break;
                }
                
            }
            catch (Exception ex)
            {
                
                throw;
            }
        }

        private void BuscarYGuardarValie(string ValueChanged, int row)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;

            String strSQL = String.Format("SELECT {0},{1}  FROM {2} Where contains({0},'%{3}%')",
                                          Constantes.Views.oitm.ItemCode,
                                          Constantes.Views.oitm.ItemName,
                                          Constantes.Views.oitm.OITM,                                      
                                          ValueChanged);
            Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsCards.DoQuery(strSQL);
            SAPbobsCOM.Fields fields = rsCards.Fields;

            string ItemCode = rsCards.Fields.Item("ItemCode").Value.ToString();
            string ItemName = rsCards.Fields.Item("ItemName").Value.ToString();
            if (ItemCode != "")
            {
                
                oGrid.DataTable.SetValue("DSCRIPTION", row, fields.Item("ItemName").Value.ToString());
                oGrid.DataTable.SetValue("ITEMCODE", row, fields.Item("ItemCode").Value.ToString());
                oGrid.CommonSetting.SetCellEditable(row + 1, 2, false);
            }
                  

        }

        private void cargarInfoInicial()
        {
            //if (CargaInicial == false)
            //{
                string usrCurrent = B1.Company.UserName;
                SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                SAPbouiCOM.EditText nomAut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
                SAPbouiCOM.Button btn_Save = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
                SAPbouiCOM.Button btn_Crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;

                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    nomAut.Value = usrCurrent;
                    btn_Save.Item.Enabled = false;
                    btn_Crear.Item.Enabled = true;

                }
                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    btn_Save.Item.Enabled = true;
                    btn_Crear.Item.Enabled = false;
                }
                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    btn_Save.Item.Enabled = false;
                    btn_Crear.Item.Enabled = true;


                }
                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //ok y en estado abierto falta agregarr
                    btn_Save.Item.Enabled = false;
                    btn_Crear.Item.Enabled = true;


                }
                formActual = B1.Application.Forms.ActiveForm.UniqueID;
                txt.Active = true;
                nomAut.Item.Enabled = false;
                //if(B1.Application.Forms.ActiveForm.Mode==)
                //CargaInicial = true;
            //}
            

            
        }

        DateTime agregoLinea;
        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {

            try
            {    SAPbouiCOM.Form oForm;
                BubbleEvent = false;
               
                if (pVal.BeforeAction == true)
                {
                   

                    SAPbouiCOM.Matrix oMatrix;
                    switch (pVal.MenuUID)
                    {
                    

                        case "1292":
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.Views.recepImpo.grid:
                                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
                                    DateTime inicio = DateTime.Now;
                                
                                    TimeSpan duracion = inicio-agregoLinea ;

                                    if (agregoLinea.Year==1 || (duracion.Seconds >=2 || duracion.Minutes>1))
                                    {
                                        oGrid.DataTable.Rows.Add();
                                        agregoLinea = DateTime.Now;
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 2, false);
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count , 21, false);
                                        oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 22, false);
                                        oGrid.AutoResizeColumns();
                                    }

                            
                                    BubbleEvent = false;
                                 
                                    break;
                            }

                            break;
                        case "1293":

                            switch (ItemActiveMenu)
                            {
                                    //ejemplo con una matrix 
                                case ventaRT.Constantes.Views.recepImpo.mtx:
                                    oForm = B1.Application.Forms.ActiveForm;
                                    oMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item(ventaRT.Constantes.Views.recepImpo.mtx).Specific);
                                    //SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@UDT");
                                    int nRow = (int)oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    oMatrix.FlushToDataSource();
                             
                                    oMatrix.LoadFromDataSource();
                                    BubbleEvent = false;

                                    break;

                                case ventaRT.Constantes.Views.recepImpo.grid:
                                    bool banderita = false;
                                    oForm = B1.Application.Forms.ActiveForm;
                                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
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
                        case"1294":
                            BubbleEvent = false;
                            break;
                        case "1282":
                            if (B1.Application.Forms.ActiveForm.UniqueID == formActual)
                            {
                                SAPbouiCOM.EditText docNum = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
                                SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                                SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
                                SAPbouiCOM.EditText cantCarto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
                                SAPbouiCOM.EditText usrCre = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
                                SAPbouiCOM.EditText fecha = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
                                SAPbouiCOM.EditText prove = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
                                SAPbouiCOM.EditText estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_est).Specific;
                                SAPbouiCOM.Button btnBuscar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
                                SAPbouiCOM.Button btn_save = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
                                SAPbouiCOM.Button btn_oc = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_bucOC).Specific;

                                btn_oc.Item.Enabled = true;
                                btnBuscar.Item.Enabled = true;
                                btnBuscar.Caption = "Crear";
                                btn_save.Item.Enabled = false;
                                docNum.Item.Enabled = true;

                                docNum.Value = "";
                                usrCre.Value = B1.Company.UserName; ;
                                docNum.Active = true;
                                docNum.Item.Enabled = false;
                                usrCre.Item.Enabled = false;
                                //docNum.BackColor

                                bcOC.Item.Enabled = true;
                                bcOC.Active = true;
                                comment.Item.Enabled = true;
                                cantCarto.Item.Enabled = true;
                                usrCre.Item.Enabled = false;
                                fecha.Item.Enabled = true;
                                prove.Item.Enabled = true;
                                estado.Item.Enabled = false;
                                B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                                BubbleEvent = false;
                            }

                            break;
                        case "1281":
                            if (B1.Application.Forms.ActiveForm.UniqueID == formActual)
                            {
                                SAPbouiCOM.EditText docNumFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
                                SAPbouiCOM.EditText bcOCFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                                SAPbouiCOM.EditText commentFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
                                SAPbouiCOM.EditText cantCartoFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
                                SAPbouiCOM.EditText usrCreFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
                                SAPbouiCOM.EditText fechaFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
                                SAPbouiCOM.EditText proveFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
                                SAPbouiCOM.EditText estadoFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_est).Specific;
                                SAPbouiCOM.Button btnBuscarFnd = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
                                SAPbouiCOM.Button btn_saveFnd = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
                                SAPbouiCOM.Button btn_ocFnd = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_bucOC).Specific;

                                btn_ocFnd.Item.Enabled = false;
                                btnBuscarFnd.Item.Enabled = true;
                                btnBuscarFnd.Caption = "Buscar";
                                btn_saveFnd.Item.Enabled = false;
                                docNumFnd.Item.Enabled = true;
                                docNumFnd.Value = "";
                                bcOCFnd.Value = "";
                                fechaFnd.Value = "";
                                usrCreFnd.Value = "";
                                proveFnd.Value = "";
                                estadoFnd.Value = "";
                                cantCartoFnd.Value = "0";

                                //docNum.BackColor
                                docNumFnd.Active = true;
                                bcOCFnd.Item.Enabled = false;
                                commentFnd.Item.Enabled = false;
                                cantCartoFnd.Item.Enabled = false;
                                usrCreFnd.Item.Enabled = false;
                                fechaFnd.Item.Enabled = false;
                                proveFnd.Item.Enabled = false;
                                estadoFnd.Item.Enabled = false;
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

                if (eventInfo.ItemUID == ventaRT.Constantes.Views.recepImpo.grid)
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

     
        private void ThisSapApiForm_OnAfterItemPressed(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                

                switch (pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:
                        switch (pVal.ItemUID)
                        {
                            case Constantes.Views.recepImpo.btn_bucOC:

                                SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                                if (txt.Value.ToString() != "")
                                {
                                    String strSQL = String.Format("SELECT {0},{1}  FROM {2} WHERE {3} = '{4}'",
                                        Constantes.Views.Oc_cabecera.DocEntry,   
                                        Constantes.Views.Oc_cabecera.CardName,                                    
                                        Constantes.Views.Oc_cabecera.OPOR,
                                        Constantes.Views.Oc_cabecera.DocNum,
                                        txt.Value.ToString());
                                    Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rsCards.DoQuery(strSQL);
                                    string DocEntry = rsCards.Fields.Item("DocEntry").Value.ToString();
                                    string CardName = rsCards.Fields.Item("CardName").Value.ToString();

                                    if (DocEntry != "")
                                    {
                                        String strSQL2 = String.Format("SELECT {0},{1},{2}  FROM {3} WHERE {4} = '{5}'",
                                       Constantes.Views.Oc_Detalle.ItemCode,
                                       Constantes.Views.Oc_Detalle.Dscription,
                                       Constantes.Views.Oc_Detalle.Quantity,
                                       Constantes.Views.Oc_Detalle.POR1,
                                       Constantes.Views.Oc_cabecera.DocEntry,
                                       DocEntry);
                                        Recordset rsCards2 = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rsCards2.DoQuery(strSQL2);

                                        SAPbouiCOM.EditText txtProv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
                                        txtProv.Value = CardName;
                                        txt.Active = true;
                                        txtProv.Item.Enabled=false;                                        
                                        //llenarMatrix(rsCards2);
                                        llenarGrid(rsCards2);
                                    }
                                  
                                   


                                }
                         
                                BubbleEvent = true;



                                break;

                            case Constantes.Views.recepImpo.btn_save:
                                SAPbouiCOM.EditText docNumSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
                                SAPbouiCOM.EditText bcOCSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                                SAPbouiCOM.EditText commentSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
                                SAPbouiCOM.EditText cantCartoSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
                                SAPbouiCOM.EditText usrCreSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
                                SAPbouiCOM.EditText usrAuto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_Aut).Specific;
                                SAPbouiCOM.EditText fechaSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
                                SAPbouiCOM.EditText proveSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
                                SAPbouiCOM.EditText estadoSave = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_est).Specific;
                                SAPbouiCOM.Button btnBuscarSave = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
                                SAPbouiCOM.Button btn_save = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
                                SAPbouiCOM.Button btn_ocSave = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_bucOC).Specific;
                                SAPbouiCOM.Button btn_creSave = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
                                if (estadoSave.Value == "A") 
                                {
                                    string usuarioAct =B1.Company.UserName;
 
                                    //BOD1, GOPE, manager
                                    if ((usuarioAct != usrCreSave.Value) || (usuarioAct == "manager" || usuarioAct == "GOPE" || usuarioAct == "BOD1" || usuarioAct == "DVLPM")) 
                                    {
                                        //aqui iran las validaciones
                                        bool isOK = validacionesAntesDeGuardar();
                                        if (isOK)
                                        {
                                            List<ventaRT.clases.detalle_recepIpo> Lineas = new List<ventaRT.clases.detalle_recepIpo>();
                                            //verificar que traiga los datos act actualizados
                                            Lineas = obtenerLineasParaGuardarenOitm(bcOCSave.Value);



                                            //aqui se guardara el articulo
                                            SAPbobsCOM.Items oItm = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                            if (Lineas.Count != 0)
                                            {
                                                int fila = 0;

                                                foreach (var item in Lineas)
                                                {
                                                    oItm.GetByKey(item.CodNum);
                                                    oItm.BarCodes.BarCode = item.Ean13.ToString();
                                                    oItm.BarCodes.UoMEntry = -1;
                                                    oItm.BarCode = item.Ean13.ToString();
                                                    oItm.PurchaseUnitWeight = item.prodPeso;
                                                    oItm.PurchaseUnitHeight = item.prodLargo;
                                                    oItm.PurchaseUnitWidth = item.prodAncho;
                                                    oItm.PurchaseUnitLength = item.prodLargo;
                                                    oItm.UserFields.Fields.Item("U_C_Inner_Pack").Value = item.InnerPack;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_DunBar).Value = item.Dun14;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Base).Value = item.Palletbase;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Altura).Value = item.altura;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Saldo).Value = item.saldo;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Peso).Value = item.cmPeso;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Largo).Value = item.cmLargo;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Ancho).Value = item.cmAncho;
                                                    oItm.UserFields.Fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Alto).Value = item.cmAlto;


                                                    //oItm.pur = item.pro;


                                                    int d = oItm.Update();
                                                    if (d != 0)
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error en el guardado: " + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                    }
                                                    else
                                                    {
                                                        B1.Application.SetStatusBarMessage("Exito en el guardado", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                                        //validacionesDespuesDeGuardar(Lineas);
                                                        //ActualizarCabecerayCerrar();
                                                        B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                                        B1.Application.Forms.ActiveForm.Close();


                                                    }


                                                    fila++;

                                                }

                                            }
                                            oItm.GetByKey(bcOCSave.Value);

                                            //oItm.BarCode

                                            //aqui se actualizara de nuevo la tabla con los datos despues de guardarse con exito

                                        }
                                        else
                                        {
                                            B1.Application.SetStatusBarMessage("Error, Por favor revisar los datos que faltan.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                        }
                                        //despues la insercion a los campos  dichos la oitm

                                    }else
                                    {
                                        B1.Application.SetStatusBarMessage("Error, usted no puede autorizar " , SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                        B1.Application.Forms.ActiveForm.Close();

                                    }
                                
                                }

                                break;

                            case Constantes.Views.recepImpo.btn_cre:
                                SAPbouiCOM.Form oForm = B1.Application.Forms.ActiveForm;
                                if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                    SAPbouiCOM.EditText code = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
                                    if (code.Value != "")
                                    {
                                        B1.Application.Forms.ActiveForm.Mode= SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                        LlenarPantalla(code.Value);

                                    }


                                }
                                else if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                                    SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
                                    SAPbouiCOM.EditText cantCarto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
                                    SAPbouiCOM.EditText usrCre = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
                                    SAPbouiCOM.EditText fecha = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
                                    SAPbouiCOM.EditText prove = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;

                                    int IDnEXT = obtenerUltimoID("Cabecera");
                                    IDnEXT = IDnEXT + 1;
                                    bool isValid = validaciones("Cabecera");
                                    if (isValid == true)
                                    {
                                        SAPbobsCOM.UserTable oUserTableCa;
                                        oUserTableCa = B1.Company.UserTables.Item("CAB_REC_IMP");
                                        oUserTableCa.Code = IDnEXT.ToString();
                                        oUserTableCa.Name = IDnEXT.ToString();
                                        oUserTableCa.UserFields.Fields.Item("U_Num_OC").Value = bcOC.Value.ToString();
                                        oUserTableCa.UserFields.Fields.Item("U_Estado").Value = "A";
                                        oUserTableCa.UserFields.Fields.Item("U_Nom_Creador").Value = usrCre.Value.ToString();
                                        string fechaHoy = fecha.Value;
                                        if (fechaHoy != "")
                                        {
                                            oUserTableCa.UserFields.Fields.Item("U_Fecha").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fechaHoy);
                                        }
                                        else
                                        {
                                            string fechhoydt = DateTime.Now.ToString("yyyyMMdd");

                                            oUserTableCa.UserFields.Fields.Item("U_Fecha").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fechhoydt);
                                        }



                                        oUserTableCa.UserFields.Fields.Item("U_Nom_Proveedor").Value = prove.Value.ToString();
                                        oUserTableCa.UserFields.Fields.Item("U_Total_Carton").Value = cantCarto.Value;
                                        oUserTableCa.UserFields.Fields.Item("U_Comentarios").Value = comment.Value.ToString();

                                        int i = oUserTableCa.Add();

                                        if (i != 0)
                                        {
                                            B1.Application.SetStatusBarMessage("Error" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                        }
                                        else
                                        {
                                            B1.Application.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            //oForm = B1.Application.Forms.Item("edm");
                                            bool TodoOk = insertarLineas( bcOC.Value.ToString());
                                            if (TodoOk == true)
                                            {
                                                B1.Application.SetStatusBarMessage("Creado con EXITO.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                                oForm.Close();

                                                //SAPbouiCOM.EditText numDoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ContabilizacionDeNominas.Constantes.Views.recepImpo.txt_numI).Specific;
                                                //SAPbouiCOM.EditText estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ContabilizacionDeNominas.Constantes.Views.recepImpo.txt_est).Specific;
                                                //numDoc.Item.Enabled = true;
                                                //numDoc.Value = IDnEXT.ToString();
                                                //estado.Item.Enabled = true;
                                                //estado.Value = "A";
                                                //numDoc.Item.Enabled = false;
                                               

                                                //B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                            }
                                            else
                                            {
                                                B1.Application.SetStatusBarMessage("Error, ha pasado algo con las línea, favor veriicar lo datos escritos.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            }

                                            //oForm.Close();



                                        }
                                    }
                                    else
                                    {
                                        B1.Application.SetStatusBarMessage("No se guardaron los datos por que no tienes los datos necesarios", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    }

                                }

                               
                          
                                BubbleEvent = true;
                                break;
                 


                        }
                        break;
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
           SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
           SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
           SAPbouiCOM.EditText cantCarto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
           SAPbouiCOM.EditText ustAuto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
           SAPbouiCOM.EditText fecha = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
           SAPbouiCOM.EditText prove = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
           SAPbouiCOM.EditText txtNumL = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
           oUserTableCa = B1.Company.UserTables.Item("CAB_REC_IMP");

           oUserTableCa.GetByKey(txtNumL.Value);       
           oUserTableCa.UserFields.Fields.Item("U_Estado").Value = "C";
           oUserTableCa.UserFields.Fields.Item("U_Nom_Autorizador").Value = B1.Company.UserName;
           string fechaHoy = fecha.Value;
           if (fechaHoy != "")
           {
               oUserTableCa.UserFields.Fields.Item("U_Fecha").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fechaHoy);
           }
           else
           {
               string fechhoydt = DateTime.Now.ToString("yyyyMMdd");
               oUserTableCa.UserFields.Fields.Item("U_Fecha").Value = SSIFramework.Utilidades.GenericFunctions.GetDate(fechhoydt);
           }

            oUserTableCa.UserFields.Fields.Item("U_Total_Carton").Value = cantCarto.Value;
            oUserTableCa.UserFields.Fields.Item("U_Comentarios").Value = comment.Value.ToString();
            int i = oUserTableCa.Update();

            if (i != 0)
            {
                B1.Application.SetStatusBarMessage("Error" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            else
            {
                B1.Application.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                //oForm = B1.Application.Forms.Item("edm");
            }
        }

        private void validacionesDespuesDeGuardar(List<clases.detalle_recepIpo> Lineas)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
            SAPbouiCOM.EditText docNum = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
       
            if (Lineas.Count != 0)
            {
                List<ventaRT.clases.detalle_recepIpo> DetalleBD = new List<ventaRT.clases.detalle_recepIpo>();
                List<ventaRT.clases.detalle_recepIpo> lineasaAgregar = new List<ventaRT.clases.detalle_recepIpo>();
                DetalleBD= ObtenerDetalle(docNum.Value);

                if (Lineas.Count == DetalleBD.Count)
                {

                    insertarLineas("",true);
                }
                    //sino no contiene las mismas lineas
                else
                {

                    bool noExiste = false;
                    foreach (var item in Lineas)
                    {
                        

                        foreach (var itemDB in DetalleBD)
                        {
                            if (item.CodNum == itemDB.CodNum)
                            {
                                noExiste = true;
                               
                                //actualizar el articulo
                                ActualizarLineaUnaXUna(item, itemDB.Code);
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
                            if (item.CodNum == itemGrid.CodNum)
                            {
                                ExisteAun = true;
                                break;
                            }
                            else { ExisteAun = false; }
                        }
                        if (ExisteAun==false)
                        {
                            //eliminamos el articulo
                            elimininarArticulo(item.Code);

                        }
                    }
                  
                    
                }
            }
          
         
        }

        private void elimininarArticulo(string code)
        {
            SAPbobsCOM.UserTable oUserTableDE;
            oUserTableDE = B1.Company.UserTables.Item("DET_REC_IMP");    
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

        private void AñadirLineaLineaUnaXUna(clases.detalle_recepIpo item)
        {
            SAPbobsCOM.UserTable oUserTableDE;
            oUserTableDE = B1.Company.UserTables.Item("DET_REC_IMP");
            int IDnEXT = obtenerUltimoID("Detalle");
            oUserTableDE.Code = IDnEXT.ToString();
            oUserTableDE.Name = IDnEXT.ToString();
            oUserTableDE.UserFields.Fields.Item("U_Cod_Articulo").Value = item.CodNum;
            oUserTableDE.UserFields.Fields.Item("U_Nom_Articulo").Value = item.DescCod;
            oUserTableDE.UserFields.Fields.Item("U_Cant_OC").Value = item.CantidaOc;
            oUserTableDE.UserFields.Fields.Item("U_Cant_Carton").Value = item.CantidaCartones;
            oUserTableDE.UserFields.Fields.Item("U_Case_Pack").Value = item.CasePack;
            oUserTableDE.UserFields.Fields.Item("U_Cant_Recibida").Value = item.CantidaRecibidas;
            oUserTableDE.UserFields.Fields.Item("U_EanBar").Value = item.Ean13;
            oUserTableDE.UserFields.Fields.Item("U_Peso").Value = item.prodPeso;
            oUserTableDE.UserFields.Fields.Item("U_Largo").Value = item.prodLargo;
            oUserTableDE.UserFields.Fields.Item("U_Ancho").Value = item.prodAncho;
            oUserTableDE.UserFields.Fields.Item("U_Alto").Value = item.prodAlto;
            oUserTableDE.UserFields.Fields.Item("U_Inner_Pack").Value = item.InnerPack;
            oUserTableDE.UserFields.Fields.Item("U_DunBar").Value = item.Dun14;
            oUserTableDE.UserFields.Fields.Item("U_C_Peso").Value = item.cmPeso;
            oUserTableDE.UserFields.Fields.Item("U_C_Largo").Value = item.cmLargo;
            oUserTableDE.UserFields.Fields.Item("U_C_Ancho").Value = item.cmAncho;
            oUserTableDE.UserFields.Fields.Item("U_C_Alto").Value = item.cmAlto;
            oUserTableDE.UserFields.Fields.Item("U_Base").Value = item.Palletbase;
            oUserTableDE.UserFields.Fields.Item("U_Altura").Value = item.altura;
            oUserTableDE.UserFields.Fields.Item("U_Saldo").Value = item.saldo;
            oUserTableDE.UserFields.Fields.Item("U_Num_OC").Value = item.DocNum;

            oUserTableDE.UserFields.Fields.Item("U_Carton_Pallet").Value = item.cantPallet;
            oUserTableDE.UserFields.Fields.Item("U_Unidad_Pallet").Value = item.UnitXPalleT;




            int d = oUserTableDE.Add();
            if (d != 0)
            {
                B1.Application.SetStatusBarMessage("Error, solo se creo el documento cabecera de la oc:  "+ B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
             
               
            }
            else
            {
                B1.Application.SetStatusBarMessage("Exito en la inserción Detalle", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
               



            }
        }

        private void ActualizarLineaUnaXUna(clases.detalle_recepIpo item,string code)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
            SAPbouiCOM.EditText bcOCFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
            bool todoOk = true;
           
                int IDnEXT = obtenerUltimoID("Grid");
                int cartonPallet = 0;
             
                   
                        IDnEXT++;
                        SAPbobsCOM.UserTable oUserTableDE;
                        oUserTableDE = B1.Company.UserTables.Item("DET_REC_IMP");

                        oUserTableDE.GetByKey(code);              
                        oUserTableDE.UserFields.Fields.Item("U_Cod_Articulo").Value = item.CodNum;
                        oUserTableDE.UserFields.Fields.Item("U_Nom_Articulo").Value = item.DescCod;
                        oUserTableDE.UserFields.Fields.Item("U_Cant_OC").Value = item.CantidaOc;
                        oUserTableDE.UserFields.Fields.Item("U_Cant_Carton").Value = item.CantidaCartones;
                        oUserTableDE.UserFields.Fields.Item("U_Case_Pack").Value =item.CasePack;
                        oUserTableDE.UserFields.Fields.Item("U_Cant_Recibida").Value = item.CantidaRecibidas;
                        oUserTableDE.UserFields.Fields.Item("U_EanBar").Value = item.Ean13;
                        oUserTableDE.UserFields.Fields.Item("U_Peso").Value = item.prodPeso;
                        oUserTableDE.UserFields.Fields.Item("U_Largo").Value = item.prodLargo;
                        oUserTableDE.UserFields.Fields.Item("U_Ancho").Value = item.prodAncho;
                        oUserTableDE.UserFields.Fields.Item("U_Alto").Value = item.prodAlto;
                        oUserTableDE.UserFields.Fields.Item("U_Inner_Pack").Value = item.InnerPack;
                        oUserTableDE.UserFields.Fields.Item("U_DunBar").Value =item.Dun14;
                        oUserTableDE.UserFields.Fields.Item("U_C_Peso").Value =item.cmPeso;
                        oUserTableDE.UserFields.Fields.Item("U_C_Largo").Value = item.cmLargo;
                        oUserTableDE.UserFields.Fields.Item("U_C_Ancho").Value = item.cmAncho;
                        oUserTableDE.UserFields.Fields.Item("U_C_Alto").Value = item.cmAlto;
                        oUserTableDE.UserFields.Fields.Item("U_Base").Value =item.Palletbase;
                        oUserTableDE.UserFields.Fields.Item("U_Altura").Value = item.altura;
                        oUserTableDE.UserFields.Fields.Item("U_Saldo").Value = item.saldo;
                        oUserTableDE.UserFields.Fields.Item("U_Num_OC").Value = item.DocNum;

                        oUserTableDE.UserFields.Fields.Item("U_Carton_Pallet").Value = item.cantPallet;
                        oUserTableDE.UserFields.Fields.Item("U_Unidad_Pallet").Value = item.UnitXPalleT;




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

                    
        }
       

        private bool validacionesAntesDeGuardar()
        {
            SAPbouiCOM.EditText docNum = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
            SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
            SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
            SAPbouiCOM.EditText cantCarto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
            SAPbouiCOM.EditText usrCre = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
            SAPbouiCOM.EditText fecha = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
            SAPbouiCOM.EditText prove = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
            SAPbouiCOM.EditText estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_est).Specific;
            SAPbouiCOM.Button btnBuscar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
            SAPbouiCOM.Button btn_save = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
            SAPbouiCOM.Button btn_oc = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_bucOC).Specific;
            bool todoOk = true;

            if (bcOC.Value == "" && docNum.Value == "" && cantCarto.Value=="0")
            {
                todoOk = false;
            }
            if (todoOk!=false)
            {
                todoOk = verificarLineasParaGuardar();
            }
          

            

            return todoOk;
        }

        private bool verificarLineasParaGuardar()
        {
            bool todoOk = true;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
            if (oGrid.DataTable.Rows.Count != 0)
            {
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    if (todoOk == false)
                    {
                        break;
                    }
                    for (int x = 0; x < oGrid.DataTable.Columns.Count; x++)
                    {

                        string valor = Convert.ToString(oGrid.DataTable.GetValue(x, i));
                        if (valor=="")
                        {
                            todoOk = false;
                            break;

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
            SAPbouiCOM.EditText docNum = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_numI).Specific;
            SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
            SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_come).Specific;
            SAPbouiCOM.EditText cantCarto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_canoes).Specific;
            SAPbouiCOM.EditText usrCre = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_cre).Specific;
            SAPbouiCOM.EditText usrAuto = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_Aut).Specific;
            SAPbouiCOM.EditText fecha = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_fecI).Specific;
            SAPbouiCOM.EditText prove = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_prov).Specific;
            SAPbouiCOM.EditText estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_est).Specific;
            SAPbouiCOM.Button btnBuscar = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cre).Specific;
            SAPbouiCOM.Button btn_save = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_save).Specific;
            SAPbouiCOM.Button btn_cancel = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_cancel).Specific;
            SAPbouiCOM.Button btn_oc = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.btn_bucOC).Specific;
        
            comment.Item.Enabled = true;
            //tendria que hacer un select de ambas tablas yy con eso llenar los campos que  se e usuaron y ver que quede bien el modo ok,
            //y lo mas  dificil que si modifican algo pase a modo update y si solo no tiene autorizador y que este abierta
            //comment.Value=
            ventaRT.clases.cabecera_recepIpo cabecera = new ventaRT.clases.cabecera_recepIpo();
            ventaRT.clases.detalle_recepIpo detalle = new ventaRT.clases.detalle_recepIpo();
            List<ventaRT.clases.detalle_recepIpo> lineas = new  List<ventaRT.clases.detalle_recepIpo>();
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
            cabecera = ObtenerCabecera(code);
            if (cabecera.DocNum != "")
            {
                lineas = ObtenerDetalle(cabecera.DocNum);

                estado.Item.Enabled = true;
                estado.Value = cabecera.estado;
                fecha.Item.Enabled = true;
                if (cabecera.fecha != "")
                {
                    DateTime dd = DateTime.Parse(cabecera.fecha);
                    //string fechaNormal = cabecera.fecha.Substring(0, 10);              
                    //fechaNormal=fechaNormal.Replace("/", "");
                    fecha.Value = SSIFramework.Utilidades.GenericFunctions.GetSapDate(dd);
                }
              
                comment.Item.Enabled = true;

            
                comment.Value = cabecera.comment;
                prove.Item.Enabled = true;
                prove.Value = cabecera.NomEmp;
                bcOC.Item.Enabled = true;
                bcOC.Value = cabecera.DocNum;
                usrCre.Item.Enabled = true;
                usrCre.Value = cabecera.creador;
                cantCarto.Item.Enabled = true;
                cantCarto.Value = cabecera.totalCarton.ToString();
                docNum.Item.Enabled = false;
                bcOC.Item.Enabled = false;
                prove.Item.Enabled = false;
                estado.Item.Enabled = false;
                usrCre.Item.Enabled = false;
               
                if (lineas.Count != 0)
                {
                    int fila = 0;
                    oGrid.DataTable.Rows.Clear();
                    foreach (var item in lineas)
                    {
                        oGrid.DataTable.Rows.Add();
                       
                                       
                        oGrid.DataTable.SetValue("ITEMCODE", fila, item.CodNum);
                        oGrid.DataTable.SetValue("DSCRIPTION", fila , item.DescCod);
                        oGrid.DataTable.SetValue("QUANTITY", fila, item.CantidaOc.ToString());
                        oGrid.DataTable.SetValue("RECantCa", fila, item.CantidaCartones.ToString());
                        oGrid.DataTable.SetValue("REPack", fila, item.CasePack);
                        oGrid.DataTable.SetValue("RECantReci", fila, item.CantidaRecibidas.ToString());
                        oGrid.DataTable.SetValue("Pro.EAN13", fila, item.Ean13.ToString());
                        oGrid.DataTable.SetValue("Pro.Peso", fila, item.prodPeso.ToString());
                        oGrid.DataTable.SetValue("Pro.Largo", fila, item.prodLargo.ToString());
                        oGrid.DataTable.SetValue("Pro.Ancho", fila, item.prodAncho.ToString());
                        oGrid.DataTable.SetValue("Pro.Alto", fila, item.prodAlto.ToString());
                        oGrid.DataTable.SetValue("CM.InPack", fila, item.InnerPack);
                        oGrid.DataTable.SetValue("CM.DUN14", fila, item.Dun14);
                        oGrid.DataTable.SetValue("CM.Peso", fila, item.cmPeso);
                        oGrid.DataTable.SetValue("CM.Largo", fila, item.cmLargo);
                        oGrid.DataTable.SetValue("CM.Ancho", fila, item.cmAncho);
                        oGrid.DataTable.SetValue("CM.Alto", fila, item.cmAlto);
                        oGrid.DataTable.SetValue("PA.Base", fila, item.Palletbase.ToString());
                        oGrid.DataTable.SetValue("PA.Altura", fila, item.altura.ToString());             
                        oGrid.DataTable.SetValue("cart_Pall", fila, item.cantPallet.ToString());
                        oGrid.DataTable.SetValue("uniPall", fila, item.UnitXPalleT.ToString());
                        oGrid.DataTable.SetValue("PA.Saldo", fila, item.saldo.ToString());

                        oGrid.AutoResizeColumns();
                        oGrid.CommonSetting.SetCellEditable(fila+1, 2, false);
                        oGrid.CommonSetting.SetCellEditable(fila+1, 21, false);
                        oGrid.CommonSetting.SetCellEditable(fila+1, 22, false);
                        fila++;
                        
                    }

                }


                //porbar con busqueda de un docu  cerrado
                if (cabecera.estado != "A")
                {
                    fecha.Item.Enabled = false;                 
                    comment.Item.Enabled = false;
                    usrAuto.Item.Enabled = true;
                    usrAuto.Value = cabecera.autorizador;
                    cantCarto.Item.Enabled = false;
                    oGrid.Item.Enabled = false;
                    //B1.Application.Forms.ActiveForm.;
                    //btn_cancel.Item.Click();
                    btn_save.Item.Enabled = false;
                    //pienso que se pued ecrear un boton invisible, acelro visivble, clickearlo, enable false para urauto ymodo invisible del boton falso
                  // usrAuto.Item.Enabled = false;
                  
                }
             




                B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
         
            //if ( usrCre.Value != B1.Company.UserName)
            //{}
            //aqui se agregara el tema de la funcion para que el creador no sea el mismo que el autorizador
          
        }

        private List<ventaRT.clases.detalle_recepIpo> ObtenerDetalle(string docnum)
        {


            List<ventaRT.clases.detalle_recepIpo> Lineas = new List<ventaRT.clases.detalle_recepIpo>();

            String strSQL = String.Format("SELECT {0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21} ,{22}  FROM {23} Where {24}='{25}'",
                                       Constantes.Views.DET_REC_IMPT.U_Cod_Articulo,
                                       Constantes.Views.DET_REC_IMPT.U_Nom_Articulo,
                                       Constantes.Views.DET_REC_IMPT.U_Cant_OC,
                                       Constantes.Views.DET_REC_IMPT.U_Cant_Carton,
                                       Constantes.Views.DET_REC_IMPT.U_Case_Pack,
                                       Constantes.Views.DET_REC_IMPT.U_Cant_Recibida,
                                       Constantes.Views.DET_REC_IMPT.U_EanBar,
                                       Constantes.Views.DET_REC_IMPT.U_Peso,
                                       Constantes.Views.DET_REC_IMPT.U_Largo,
                                       Constantes.Views.DET_REC_IMPT.U_Ancho,
                                       Constantes.Views.DET_REC_IMPT.U_Alto,
                                       Constantes.Views.DET_REC_IMPT.U_Inner_Pack,
                                       Constantes.Views.DET_REC_IMPT.U_DunBar,
                                       Constantes.Views.DET_REC_IMPT.U_C_Peso,
                                       Constantes.Views.DET_REC_IMPT.U_C_Largo,
                                       Constantes.Views.DET_REC_IMPT.U_C_Ancho,
                                       Constantes.Views.DET_REC_IMPT.U_C_Alto,
                                       Constantes.Views.DET_REC_IMPT.U_Base,
                                       Constantes.Views.DET_REC_IMPT.U_Altura,
                                       Constantes.Views.DET_REC_IMPT.U_Saldo,
                                       Constantes.Views.DET_REC_IMPT.U_Carton_Pallet,
                                       Constantes.Views.DET_REC_IMPT.U_Unidad_Pallet,
                                       Constantes.Views.DET_REC_IMPT.Code, 
                                       Constantes.Views.DET_REC_IMPT.DET_REC_IMP,
                                       Constantes.Views.DET_REC_IMPT.U_Num_OC,
                                       docnum);

            Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsCards.DoQuery(strSQL);

            if (rsCards.RecordCount != 0)
            {
                rsCards.MoveFirst();
                
                for (int i = 1; !rsCards.EoF; i++)
                {
                    ventaRT.clases.detalle_recepIpo detalles = new ventaRT.clases.detalle_recepIpo();
                    SAPbobsCOM.Fields fields = rsCards.Fields;
                    detalles.CodNum = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Cod_Articulo).Value.ToString();
                    detalles.DescCod = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Nom_Articulo).Value.ToString(); 
                    detalles.CantidaOc = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Cant_OC).Value;
                    detalles.CantidaCartones = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Cant_Carton).Value;
                    detalles.CasePack = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Case_Pack).Value.ToString();
                    detalles.CantidaRecibidas = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Cant_Recibida).Value;
                    detalles.Ean13 = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_EanBar).Value;
                    detalles.prodPeso = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Peso).Value;
                    detalles.prodLargo = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Largo).Value;
                    detalles.prodAncho = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Ancho).Value;
                    detalles.prodAlto = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Alto).Value;
                    detalles.InnerPack = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Inner_Pack).Value.ToString();
                    detalles.Dun14 = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_DunBar).Value.ToString();
                    detalles.cmPeso = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Peso).Value.ToString();
                    detalles.cmLargo = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Largo).Value.ToString();
                    detalles.cmAncho = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Ancho).Value.ToString();
                    detalles.cmAlto = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_C_Alto).Value.ToString();
                    detalles.Palletbase = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Base).Value;
                    detalles.altura = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Altura).Value;
                    detalles.saldo = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Saldo).Value;
                    detalles.cantPallet = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Carton_Pallet).Value;
                    detalles.UnitXPalleT = fields.Item(Constantes.Views.DET_REC_IMPTTabla.U_Unidad_Pallet).Value;
                    detalles.Code = fields.Item("Code").Value;

                    Lineas.Add(detalles);
                    rsCards.MoveNext();
                }
               
              
              
            }

            return Lineas;
        }

        private clases.cabecera_recepIpo ObtenerCabecera(string code)
        {
            ventaRT.clases.cabecera_recepIpo cabecera = new ventaRT.clases.cabecera_recepIpo();
            String strSQL = String.Format("SELECT {0},{1},{2},{3},{4},{5},{6},{7},{9}  FROM {8} Where {9}='{10}'",
                                         Constantes.Views.CAB_REC_IMPT.U_Num_OC,
                                         Constantes.Views.CAB_REC_IMPT.U_Nom_Proveedor,
                                         Constantes.Views.CAB_REC_IMPT.U_Nom_Creador,
                                         Constantes.Views.CAB_REC_IMPT.U_Total_Carton,
                                         Constantes.Views.CAB_REC_IMPT.U_Comentarios, 
                                         Constantes.Views.CAB_REC_IMPT.U_Nom_Autorizador,
                                         Constantes.Views.CAB_REC_IMPT.U_Fecha,
                                         Constantes.Views.CAB_REC_IMPT.U_Estado,
                                         Constantes.Views.CAB_REC_IMPT.CAB_REC_IMP,
                                         Constantes.Views.CAB_REC_IMPT.Code,
                                         code);
            Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsCards.DoQuery(strSQL);
            if (rsCards.RecordCount != 0)
            {
                SAPbobsCOM.Fields fields = rsCards.Fields;
                cabecera.DocNum = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Num_OC).Value.ToString();
                cabecera.NomEmp = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Nom_Proveedor).Value.ToString();
                cabecera.creador = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Nom_Creador).Value.ToString();
                cabecera.totalCarton =fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Total_Carton).Value;
                cabecera.comment = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Comentarios).Value.ToString();
                cabecera.autorizador = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Nom_Autorizador).Value.ToString();
                cabecera.fecha = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Fecha).Value.ToString();
                cabecera.estado = fields.Item(Constantes.Views.CAB_REC_IMPTTabla.U_Estado).Value.ToString();
            }
          
            return cabecera;
            
        }

        private bool insertarLineas(string DocNum,bool tipoUpdate=false)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
            SAPbouiCOM.EditText bcOCFnd = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
            bool todoOk = true;
           
                int IDnEXT = obtenerUltimoID("Grid");
                int cartonPallet = 0;
                if (tipoUpdate==true)
                {
                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {
                        IDnEXT++;
                        SAPbobsCOM.UserTable oUserTableDE;
                        oUserTableDE = B1.Company.UserTables.Item("DET_REC_IMP");
                        
                        //TENGO QUE VER SI las lineas  el grid actual se diferencia con la anterior
                        //que pasa si eliminaron una linea, o modificar un codigo de una linea?
                        //se tiene que hacer una funcion preguntando si el articulo esta en la lista, si lo esta, se tiene que obtener el codigo que tiene
                        //sino se tiene se tiene que agregar una una lista que seran nuevas lineas pero
                        //que pasa si hay eenos lineas, entonces tambien debo buscar cual es el elmento eliminado y que elimine esas lineas en especifico
                        //esto demorara dos dias mas como minimo solo eso,  quedan los bugs, mas modo tipo cerrado, validar antes de guardar y por ultimo al ventana nueva
                        //que demorara como dos días mas  tambien , ent total el proyecto estaria  mejor de los casos el 29 para  el 2 de diciembre es el realista , ya que no trabajo ni el 30 ni el 1

                        string CodeOb = getCodePerCodNUM(oGrid.DataTable.GetValue("ITEMCODE", i));

                        oUserTableDE.GetByKey(CodeOb);              
                        oUserTableDE.UserFields.Fields.Item("U_Cod_Articulo").Value = oGrid.DataTable.GetValue("ITEMCODE", i);
                        oUserTableDE.UserFields.Fields.Item("U_Nom_Articulo").Value = oGrid.DataTable.GetValue("DSCRIPTION", i);
                        oUserTableDE.UserFields.Fields.Item("U_Cant_OC").Value = oGrid.DataTable.GetValue("QUANTITY", i);
                        oUserTableDE.UserFields.Fields.Item("U_Cant_Carton").Value = oGrid.DataTable.GetValue("RECantCa", i);
                        oUserTableDE.UserFields.Fields.Item("U_Case_Pack").Value = oGrid.DataTable.GetValue("REPack", i);
                        oUserTableDE.UserFields.Fields.Item("U_Cant_Recibida").Value = oGrid.DataTable.GetValue("RECantReci", i);
                        oUserTableDE.UserFields.Fields.Item("U_EanBar").Value = oGrid.DataTable.GetValue("Pro.EAN13", i);
                        oUserTableDE.UserFields.Fields.Item("U_Peso").Value = oGrid.DataTable.GetValue("Pro.Peso", i);
                        oUserTableDE.UserFields.Fields.Item("U_Largo").Value = oGrid.DataTable.GetValue("Pro.Largo", i);
                        oUserTableDE.UserFields.Fields.Item("U_Ancho").Value = oGrid.DataTable.GetValue("Pro.Ancho", i);
                        oUserTableDE.UserFields.Fields.Item("U_Alto").Value = oGrid.DataTable.GetValue("Pro.Alto", i);
                        oUserTableDE.UserFields.Fields.Item("U_Inner_Pack").Value = oGrid.DataTable.GetValue("CM.InPack", i);
                        oUserTableDE.UserFields.Fields.Item("U_DunBar").Value = oGrid.DataTable.GetValue("CM.DUN14", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Peso").Value = oGrid.DataTable.GetValue("CM.Peso", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Largo").Value = oGrid.DataTable.GetValue("CM.Largo", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Ancho").Value = oGrid.DataTable.GetValue("CM.Ancho", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Alto").Value = oGrid.DataTable.GetValue("CM.Alto", i);
                        oUserTableDE.UserFields.Fields.Item("U_Base").Value = oGrid.DataTable.GetValue("PA.Base", i);
                        oUserTableDE.UserFields.Fields.Item("U_Altura").Value = oGrid.DataTable.GetValue("PA.Altura", i);
                        oUserTableDE.UserFields.Fields.Item("U_Saldo").Value = oGrid.DataTable.GetValue("PA.Saldo", i);
                        oUserTableDE.UserFields.Fields.Item("U_Num_OC").Value = bcOCFnd.Value;

                        oUserTableDE.UserFields.Fields.Item("U_Carton_Pallet").Value = oGrid.DataTable.GetValue("cart_Pall", i);
                        oUserTableDE.UserFields.Fields.Item("U_Unidad_Pallet").Value = oGrid.DataTable.GetValue("uniPall", i);




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

                    }
                }
                else
                {
                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {
                        IDnEXT++;
                        SAPbobsCOM.UserTable oUserTableDE;
                        oUserTableDE = B1.Company.UserTables.Item("DET_REC_IMP");
                        oUserTableDE.Code = IDnEXT.ToString();
                        oUserTableDE.Name = IDnEXT.ToString();
                        oUserTableDE.UserFields.Fields.Item("U_Cod_Articulo").Value = oGrid.DataTable.GetValue("ITEMCODE", i);
                        oUserTableDE.UserFields.Fields.Item("U_Nom_Articulo").Value = oGrid.DataTable.GetValue("DSCRIPTION", i);
                        oUserTableDE.UserFields.Fields.Item("U_Cant_OC").Value = oGrid.DataTable.GetValue("QUANTITY", i);
                        oUserTableDE.UserFields.Fields.Item("U_Cant_Carton").Value = oGrid.DataTable.GetValue("RECantCa", i);
                        oUserTableDE.UserFields.Fields.Item("U_Case_Pack").Value = oGrid.DataTable.GetValue("REPack", i);
                        oUserTableDE.UserFields.Fields.Item("U_Cant_Recibida").Value = oGrid.DataTable.GetValue("RECantReci", i);
                        oUserTableDE.UserFields.Fields.Item("U_EanBar").Value = oGrid.DataTable.GetValue("Pro.EAN13", i);
                        oUserTableDE.UserFields.Fields.Item("U_Peso").Value = oGrid.DataTable.GetValue("Pro.Peso", i);
                        oUserTableDE.UserFields.Fields.Item("U_Largo").Value = oGrid.DataTable.GetValue("Pro.Largo", i);
                        oUserTableDE.UserFields.Fields.Item("U_Ancho").Value = oGrid.DataTable.GetValue("Pro.Ancho", i);
                        oUserTableDE.UserFields.Fields.Item("U_Alto").Value = oGrid.DataTable.GetValue("Pro.Alto", i);
                        oUserTableDE.UserFields.Fields.Item("U_Inner_Pack").Value = oGrid.DataTable.GetValue("CM.InPack", i);
                        oUserTableDE.UserFields.Fields.Item("U_DunBar").Value = oGrid.DataTable.GetValue("CM.DUN14", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Peso").Value = oGrid.DataTable.GetValue("CM.Peso", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Largo").Value = oGrid.DataTable.GetValue("CM.Largo", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Ancho").Value = oGrid.DataTable.GetValue("CM.Ancho", i);
                        oUserTableDE.UserFields.Fields.Item("U_C_Alto").Value = oGrid.DataTable.GetValue("CM.Alto", i);
                        oUserTableDE.UserFields.Fields.Item("U_Base").Value = oGrid.DataTable.GetValue("PA.Base", i);
                        oUserTableDE.UserFields.Fields.Item("U_Altura").Value = oGrid.DataTable.GetValue("PA.Altura", i);
                        oUserTableDE.UserFields.Fields.Item("U_Saldo").Value = oGrid.DataTable.GetValue("PA.Saldo", i);
                        oUserTableDE.UserFields.Fields.Item("U_Num_OC").Value = DocNum;


                        oUserTableDE.UserFields.Fields.Item("U_Carton_Pallet").Value = oGrid.DataTable.GetValue("cart_Pall", i);
                        oUserTableDE.UserFields.Fields.Item("U_Unidad_Pallet").Value = oGrid.DataTable.GetValue("uniPall", i);




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

        private List<ventaRT.clases.detalle_recepIpo> obtenerLineasParaGuardarenOitm(string DocNum)
        {
            List<ventaRT.clases.detalle_recepIpo> Lineas = new List<ventaRT.clases.detalle_recepIpo>();
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;

            //int IDnEXT = obtenerUltimoID("Grid");
            int cartonPallet = 0;

            for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
            {

                ventaRT.clases.detalle_recepIpo detalles = new ventaRT.clases.detalle_recepIpo();
           
                detalles.CodNum = oGrid.DataTable.GetValue("ITEMCODE", i);
                detalles.DescCod = oGrid.DataTable.GetValue("DSCRIPTION", i);
                detalles.CantidaOc = oGrid.DataTable.GetValue("QUANTITY", i);
                detalles.CantidaCartones = oGrid.DataTable.GetValue("RECantCa", i);
                detalles.CasePack = oGrid.DataTable.GetValue("REPack", i);
                detalles.CantidaRecibidas = oGrid.DataTable.GetValue("RECantReci", i);
                detalles.Ean13 = oGrid.DataTable.GetValue("Pro.EAN13", i);
                detalles.prodPeso = oGrid.DataTable.GetValue("Pro.Peso", i);
                detalles.prodLargo = oGrid.DataTable.GetValue("Pro.Largo", i);
                detalles.prodAncho = oGrid.DataTable.GetValue("Pro.Ancho", i);
                detalles.prodAlto = oGrid.DataTable.GetValue("Pro.Alto", i);
                detalles.InnerPack = oGrid.DataTable.GetValue("CM.InPack", i);
                detalles.Dun14 = oGrid.DataTable.GetValue("CM.DUN14", i);
                detalles.cmPeso = oGrid.DataTable.GetValue("CM.Peso", i);
                detalles.cmLargo = oGrid.DataTable.GetValue("CM.Largo", i);
                detalles.cmAncho = oGrid.DataTable.GetValue("CM.Ancho", i);
                detalles.cmAlto = oGrid.DataTable.GetValue("CM.Alto", i);
                detalles.Palletbase = oGrid.DataTable.GetValue("PA.Base", i);
                detalles.altura = oGrid.DataTable.GetValue("PA.Altura", i);
                detalles.saldo = oGrid.DataTable.GetValue("PA.Saldo", i);
                detalles.cantPallet = oGrid.DataTable.GetValue("cart_Pall", i);
                detalles.UnitXPalleT = oGrid.DataTable.GetValue("uniPall", i);
                detalles.DocNum= DocNum;


                Lineas.Add(detalles);

            
            }

            return Lineas;

        }

        private bool ActualizarLineas( string DocNum)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
            bool todoOk = true;

            for (int i = 0; i < oGrid.Rows.Count; i++)
            {
                SAPbobsCOM.UserTable oUserTableDE;
                oUserTableDE = B1.Company.UserTables.Item("DET_REC_IMP");
                //oUserTableDE.GetByKey();
                oUserTableDE.UserFields.Fields.Item("U_Cod_Articulo").Value = oGrid.DataTable.GetValue("ITEMCODE", i);
                oUserTableDE.UserFields.Fields.Item("U_Nom_Articulo").Value = oGrid.DataTable.GetValue("DSCRIPTION", i);
                oUserTableDE.UserFields.Fields.Item("U_Cant_OC").Value = oGrid.DataTable.GetValue("QUANTITY", i);
                oUserTableDE.UserFields.Fields.Item("U_Cant_Carton").Value = oGrid.DataTable.GetValue("RECantCa", i);
                oUserTableDE.UserFields.Fields.Item("U_Case_Pack").Value = oGrid.DataTable.GetValue("REPack", i);
                oUserTableDE.UserFields.Fields.Item("U_Cant_Recibida").Value = oGrid.DataTable.GetValue("RECantReci", i);
                oUserTableDE.UserFields.Fields.Item("U_EanBar").Value = oGrid.DataTable.GetValue("Pro.EAN13", i);
                oUserTableDE.UserFields.Fields.Item("U_Peso").Value = oGrid.DataTable.GetValue("Pro.Peso", i);
                oUserTableDE.UserFields.Fields.Item("U_Largo").Value = oGrid.DataTable.GetValue("Pro.Largo", i);
                oUserTableDE.UserFields.Fields.Item("U_Ancho").Value = oGrid.DataTable.GetValue("Pro.Ancho", i);
                oUserTableDE.UserFields.Fields.Item("U_Alto").Value = oGrid.DataTable.GetValue("Pro.Alto", i);
                oUserTableDE.UserFields.Fields.Item("U_Inner_Pack").Value = oGrid.DataTable.GetValue("CM.InPack", i);
                oUserTableDE.UserFields.Fields.Item("U_DunBar").Value = oGrid.DataTable.GetValue("CM.DUN14", i);
                oUserTableDE.UserFields.Fields.Item("U_C_Peso").Value = oGrid.DataTable.GetValue("CM.Peso", i);
                oUserTableDE.UserFields.Fields.Item("U_C_Largo").Value = oGrid.DataTable.GetValue("CM.Largo", i);
                oUserTableDE.UserFields.Fields.Item("U_C_Ancho").Value = oGrid.DataTable.GetValue("CM.Ancho", i);
                oUserTableDE.UserFields.Fields.Item("U_C_Alto").Value = oGrid.DataTable.GetValue("CM.Alto", i);
                oUserTableDE.UserFields.Fields.Item("U_Base").Value = oGrid.DataTable.GetValue("PA.Base", i);
                oUserTableDE.UserFields.Fields.Item("U_Altura").Value = oGrid.DataTable.GetValue("PA.Altura", i);
                oUserTableDE.UserFields.Fields.Item("U_Saldo").Value = oGrid.DataTable.GetValue("PA.Saldo", i);
                oUserTableDE.UserFields.Fields.Item("U_Num_OC").Value = DocNum;
                oUserTableDE.UserFields.Fields.Item("U_Carton_Pallet").Value = oGrid.DataTable.GetValue("cart_Pall", i);
                oUserTableDE.UserFields.Fields.Item("U_Unidad_Pallet").Value = oGrid.DataTable.GetValue("uniPall", i);




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
                SAPbouiCOM.EditText bcOC = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.txt_busOC).Specific;
                ISok = bcOC.Value.ToString() == "" ? false : true;
            }
       

            return ISok;
        }

        private int obtenerUltimoID(string tipo)
        {
            int CodeNumCA = 0;
            int CodeNumDE = 0;
            if (tipo == "Cabecera")
            {
                ///SELECT  COUNT(*) from "SBO_DEVELOPMENT_HJ"."@CAB_REC_IMP" 
                String strSQL = String.Format("SELECT  COUNT(*)  FROM {0}",
                                    Constantes.Views.CAB_REC_IMPT.CAB_REC_IMP);

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
                                    Constantes.Views.DET_REC_IMPT.DET_REC_IMP);

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

        private void llenarGrid(Recordset rsCards2)
        {
            try
            {
                SAPbouiCOM.Form oForm = B1.Application.Forms.ActiveForm;
                //SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item(ContabilizacionDeNominas.Constantes.Views.recepImpo.DT_Grid);         
                //oMatrix = (Matrix)B1.Application.Forms.ActiveForm.Items.Item(ContabilizacionDeNominas.Constantes.Views.recepImpo.mtx).Specific;
                //SAPbouiCOM.Grid oGrid = oForm.Items.Item(ContabilizacionDeNominas.Constantes.Views.recepImpo.grid).Specific;
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(ventaRT.Constantes.Views.recepImpo.grid).Specific;
                if (oGrid.DataTable.Rows.Count!=0)
                {
                    oGrid.DataTable.Rows.Clear();
                }
               
                SAPbobsCOM.Fields fields = rsCards2.Fields;
               
                //int lastRowIndex = DT_GRID.Rows.Count ;              
                rsCards2.MoveFirst();
                for (int i = 1; !rsCards2.EoF; i++)
                {
                    oGrid.DataTable.Rows.Add();
                               
                    oGrid.DataTable.SetValue("ITEMCODE", oGrid.DataTable.Rows.Count - 1, fields.Item("ItemCode").Value.ToString());
                    oGrid.DataTable.SetValue("DSCRIPTION", oGrid.DataTable.Rows.Count - 1, fields.Item("Dscription").Value.ToString());
                    oGrid.DataTable.SetValue("QUANTITY", oGrid.DataTable.Rows.Count - 1, fields.Item("Quantity").Value.ToString());
                    oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 2, false);
                   oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count , 21, false);
                   oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, 22, false);
                   oGrid.AutoResizeColumns();
                   rsCards2.MoveNext();

                }

            }
            catch (Exception ex)
            {
                
                throw;
            }
           
        }

        private void llenarMatrix(Recordset rsCards2,bool primeravez=true)
        {

          
            SAPbouiCOM.Matrix oMatrix;
            oMatrix = (Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.recepImpo.mtx).Specific;
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
                int fila = i ;
                num.Value = fila.ToString();
                ItemCode.Value = fields.Item("ItemCode").Value.ToString();
                Dscription.Value = fields.Item("Dscription").Value.ToString();
                Quantity.Value = fields.Item("Quantity").Value.ToString();
                
                oMatrix.AddRow();
             
                rsCards2.MoveNext();
            }
            B1.Application.Forms.ActiveForm.Freeze(false);
        }
    }
}
