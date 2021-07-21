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
       

        public PantallaRegistro()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Registro.srf"], "ventaRT_Registro" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            formActual = "ventaRT_Registro" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();

            
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);

            
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);

 
            cargarInfoInicial();
        }

        private void cargarInfoInicial()
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;
            oCombo.Item.Visible = false;
            SForm = B1.Application.Forms.ActiveForm;
            SMatrix = SForm.Items.Item("mtx").Specific;

            SForm.EnableMenu("1290", true); SForm.EnableMenu("1289", true);
            SForm.EnableMenu("1288", true); SForm.EnableMenu("1291", true);

            SForm.EnableMenu("1282", true); SForm.EnableMenu("1283", true);
            SForm.EnableMenu("1281", true); 

            SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
            SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
            SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
            SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
            SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
            SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
            SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
            
            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
            
            if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                Insertar_Solicitud();
            }
            //if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            //{
            //    btn_crear.Caption = "Actualizar";
            //}
            //if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            //{
            //    Preparar_Find();
            //}
            //if (B1.Application.Forms.ActiveForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            //{
            //    btn_crear.Caption = "OK";
            //}

        }


        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
                if (pVal.BeforeAction)
                {
                    BubbleEvent = true;
                    switch (pVal.MenuUID)
                    {
                        case "1282":    // Crear      
                 
                            Insertar_Solicitud();
                            BubbleEvent = false;
                            break;
                        case "1281":    // Buscar                      
                            Preparar_Find();
                            BubbleEvent = false;
                            break;
                        case "1283":    // Eliminar                     
                            Eliminar_Solicitud();
                            BubbleEvent = false;
                            break;
                        case "1292":   //ADICIONAR LINEA
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.registro.mtx:
                                    SMatrix.AddRow(1, SMatrix.RowCount);
                                    SMatrix.ClearRowData(SMatrix.RowCount);
                                    //SMatrix.Columns.Item(5).Cells.Item(SMatrix.RowCount ).Specific.Value = "1";
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
                                    if (nRow < SMatrix.RowCount)
                                    {
                                            SMatrix.DeleteRow(nRow);
                                            SMatrix.FlushToDataSource();
                                            SMatrix.LoadFromDataSource();
                                    }

                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                        case "1290":    // Primero                      
                            Activar_Primero();
                            break;
                        case "1289":    // Ant                      
                            Activar_Anterior();
                            break;
                        case "1288":    // Sig                      
                            Activar_Posterior();
                            break;
                        case "1291":    // Ultimo                      
                            Activar_Ultimo();
                            break;
                    }
                    BubbleEvent = false;
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
            catch (Exception)
            {

                throw;
            }

        }


        private void ThisSapApiForm_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
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
                                                    string buscado = oCombox.Selected.Value == null ? " 0" :oCombox.Selected.Value.ToString();
                                                    if (buscado != "0")
                                                    {
                                                        Cargar_Solicitud(buscado, true);
                                                    }
                                                    BubbleEvent = false;
                                                    break;
                                                }
                                            case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                                                {
                                                    Guardar_Solicitud();
                                                    BubbleEvent = false;
                                                    break;
                                                }
                                            case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                                                {
                                                    Guardar_Solicitud();
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
                                SAPbouiCOM.DBDataSource oObHeaderDataSource = null, oObLineDataSource = null;
                                SAPbouiCOM.IChooseFromListEvent CFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;
                                oObHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
                                oObLineDataSource = SForm.DataSources.DBDataSources.Item("@DET_RSTV");
                                string CFL_Id = CFLEvent.ChooseFromListUID;
                                oCFL = SForm.ChooseFromLists.Item(CFL_Id);
                                if (pVal.FormTypeEx.Substring(0, 10) == "ventaRT_Re" && CFLEvent.SelectedObjects != null)
                                {
                                    if (pVal.ItemUID == "mtx" && pVal.ColUID == "codArt")
                                    {
                                        SMatrix.FlushToDataSource();
                                        oObLineDataSource.SetValue("U_CodArt", oObLineDataSource.Offset,
                                            CFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString());
                                        oObLineDataSource.SetValue("U_articulo", oObLineDataSource.Offset,
                                            CFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString());
                                        SMatrix.LoadFromDataSource();

                                        SMatrix.Columns.Item("codCli").Cells.Item(pVal.Row).Click();
                                        //SMatrix.SetCellFocus(pVal.Row+1, 3);
                                    }
                                    if (pVal.ItemUID == "mtx" && pVal.ColUID == "codCli")
                                    {
                                        SMatrix.FlushToDataSource();
                                        oObLineDataSource.SetValue("U_CodCli", oObLineDataSource.Offset,
                                            CFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString());
                                        oObLineDataSource.SetValue("U_cliente", oObLineDataSource.Offset,
                                            CFLEvent.SelectedObjects.GetValue("CardName", 0).ToString());
                                        SMatrix.LoadFromDataSource();
                                        SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Click();
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

                                    FillCombo(oCombo, SQLQuery);
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
                                                Guardar_Solicitud();
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

                                switch (pVal.ColUID)
                                {
                                    case "codArt":
                                        {
                                            string codart = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codArt").Cells.Item(pVal.Row).Specific).Value.ToString();
                                            if (codart == "")
                                            {
                                                B1.Application.SetStatusBarMessage("Error Codigo Articulo es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                                BubbleEvent = false;
                                            }
                                        }
                                        break;
                                    case "codCli":
                                        {
                                            string codcli = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("codCli").Cells.Item(pVal.Row).Specific).Value.ToString();
                                            if (codcli == "")
                                            {
                                                B1.Application.SetStatusBarMessage("Error Codigo Cliente es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Short, true);

                                                BubbleEvent = false;
                                            }
                                        }
                                        break;
                                    case "cant":
                                        {
                                            string cantidad = ((SAPbouiCOM.EditText)SMatrix.Columns.Item("cant").Cells.Item(pVal.Row).Specific).Value.ToString();
                                            if (cantidad == "0")
                                            { SMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific.Value = "1"; }
                                        }
                                        break;

                                }

                            }
   
                                
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }

        }




        private bool Insertar_Solicitud()
        {

            bool todoOk = true;
            string serror = "";            
            try {
                    SForm = B1.Application.Forms.ActiveForm;
                    SMatrix = SForm.Items.Item("mtx").Specific;

                    B1.Application.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                        int norecord = obtenerUltimoID("CA") + 1;
               
                        //Insertando nuevo record

                        SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
                        oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
                        oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size - 1;
                        oDbHeaderDataSource.Query();
                        oDbHeaderDataSource.InsertRecord(oDbHeaderDataSource.Size);
                        oDbHeaderDataSource.Offset = oDbHeaderDataSource.Size-1;

                        DateTime fc = DateTime.Now.Date;
                        DateTime fv = fc.AddDays(10);

                        

                        oDbHeaderDataSource.SetValue("U_numDoc", norecord, norecord.ToString());
                        oDbHeaderDataSource.SetValue("U_IdVend", norecord, obtenerVendedor());
                        oDbHeaderDataSource.SetValue("U_fechaC", norecord, fc.ToString("yyyyMMdd"));
                        oDbHeaderDataSource.SetValue("U_fechaV", norecord, fv.ToString("yyyyMMdd"));
                        oDbHeaderDataSource.SetValue("U_estado", norecord, "Nueva");
                        oDbHeaderDataSource.SetValue("U_comment", norecord, "");


                        SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                        SAPbouiCOM.EditText txt_idaut = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idaut).Specific;
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
                        txt_idvend.Value = obtenerVendedor(); 
                        txt_idaut.Value = "";
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

        private bool Preparar_Find()
        {
            bool todoOk = true; 
            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");

            if (oDbHeaderDataSource.Size == 1 && SForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                B1.Application.SetStatusBarMessage("No se puede activar Modo Busqueda porque no tiene registros... ", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;
            }
            else
            {
                try
                {
                    int borrado = 0;
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
                        SForm = B1.Application.Forms.ActiveForm;
                        SMatrix = SForm.Items.Item("mtx").Specific;

                        SAPbouiCOM.EditText txt_idvend = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_idvend).Specific;
                        SAPbouiCOM.EditText txt_numoc = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_numoc).Specific;
                        SAPbouiCOM.EditText txt_fechac = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechac).Specific;
                        SAPbouiCOM.EditText txt_fechav = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_fechav).Specific;
                        SAPbouiCOM.EditText txt_estado = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_estado).Specific;
                        SAPbouiCOM.EditText txt_com = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.txt_com).Specific;
                        SAPbouiCOM.Button btn_crear = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.btn_crear).Specific;
                        SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.registro.mtx).Specific;
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)B1.Application.Forms.ActiveForm.Items.Item("cbnd").Specific;

                        SForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        btn_crear.Caption = "Buscar";
                        oCombo.Item.Visible = true;
                        oCombo.Item.Enabled = true;
                        //mtx.Item.Enabled = false;
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

        private void Activar_Primero()
        {
            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
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
                    Cargar_Solicitud("0", false);
                    B1.Application.SetStatusBarMessage("Movimiento al Primero ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {

                }
            }
         }

        private void Activar_Anterior()
        {

            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
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

                        Cargar_Solicitud(oDbHeaderDataSource.Offset.ToString(), false);
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

        private void Activar_Posterior()
        {
            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
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
                    Cargar_Solicitud(oDbHeaderDataSource.Offset.ToString(), false);

                    B1.Application.SetStatusBarMessage("Movimiento al Siguiente ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {

                }
            }
  


        }

        private void Activar_Ultimo()
      {
            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
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
                    Cargar_Solicitud(oDbHeaderDataSource.Offset.ToString(), false);
                    B1.Application.SetStatusBarMessage("Movimiento al Ultimo ", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                }
                catch (Exception ex)
                {

                }
            }



      }

        private bool Eliminar_Solicitud()
        {
            bool todoOk = true;
            string serror = "";

            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
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

                        if (oDbHeaderDataSource.Offset == 0) { Activar_Posterior(); }
                        else { Activar_Anterior(); }
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

        private bool Guardar_Solicitud()
        {
            bool todoOk = true;
            string serror = "";
            int iRet;
            try
            {
                SAPbouiCOM.DBDataSource oDbHeaderDataSource = null, oDbLineDataSource = null;
                oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");

                oDbLineDataSource = SForm.DataSources.DBDataSources.Item("@DET_RSTV");
                SAPbobsCOM.UserTable UTDoc = B1.Company.UserTables.Item("CAB_RSTV");
                SAPbobsCOM.UserTable UTLines = B1.Company.UserTables.Item("DET_RSTV");
                //SForm.Freeze(true);
                string sCode = ""; string sName = "";
 
              

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

                            iRet = UTDoc.Update();
                            todoOk = (iRet==0);
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
     
                            iRet = UTDoc.Add();
                            todoOk = (iRet==0);
                        }
                    
                    
                        //Guardando con instrucciones SQL
                        // Buscar si existe ese codigo para update

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
                        //    SQLQuery = String.Format("INSERT INTO {0} ({7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}) "+
                        //    " VALUES('{1}','{2}','{3}','{4}','{5}','{6}','','','','{1}','{1}') ",
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
                        //                     Constantes.View.CAB_RVT.Name);
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
                    int norecord2 = obtenerUltimoID("DE") ;
                    SMatrix.FlushToDataSource();
                    for(int i=0; i <= oDbLineDataSource.Size-1; i++)
                    {

                        // Obteniendo texto de los campos de DbDataSource
                        string sCodeL = oDbLineDataSource.GetValue("Code", i);
                        string sNameL = oDbLineDataSource.GetValue("Name" ,i);
                        string scodart = oDbLineDataSource.GetValue("U_codArt",i);
                        string sart = oDbLineDataSource.GetValue("U_articulo",i);
                        string scodcli = oDbLineDataSource.GetValue("U_codCli",i);
                        string sccli = oDbLineDataSource.GetValue("U_cliente",i);
                        string scant = oDbLineDataSource.GetValue("U_cant",i);

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
                                B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                //UTLines = null;
                            }
                        }
                        //else
                        //{
                        //    if (i > 0)
                        //    {
                        //        SMatrix.DeleteRow(i);
                        //        SMatrix.FlushToDataSource();
                        //    }
                        //}
  
                    }
                    //Cargar_Lineas(sCode);
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

        private bool Cargar_Lineas(string noDoc)
        {
           bool todoOk = true;

           string serror = "";
           if (noDoc != "")
           {
               try
               {
                   SForm.Freeze(true);
                   String strSQL = String.Format("SELECT {0}, {1}, {2},{3},{4}, {5}, {6}, {8}" +
                                                       " FROM {9} Where {7}='{10}'",
                                                   Constantes.View.DET_RVT.U_codArt, //0
                                                   Constantes.View.DET_RVT.U_articulo, //1
                                                   Constantes.View.DET_RVT.U_codCli, //2
                                                   Constantes.View.DET_RVT.U_cliente, //3
                                                   Constantes.View.DET_RVT.U_cant, //4
                                                   Constantes.View.DET_RVT.U_estado, //5
                                                   Constantes.View.DET_RVT.U_idTV,//6
                                                   Constantes.View.DET_RVT.U_numOC,//7
                                                   Constantes.View.DET_RVT.Code,//8
                                                   Constantes.View.DET_RVT.DET_RV,//9
                                                   noDoc);  //10

                   Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                   rsCards.DoQuery(strSQL);
                   SMatrix.Clear();
                   SAPbobsCOM.Fields fields = rsCards.Fields;
                   rsCards.MoveFirst();
                   for (int i = 1; !rsCards.EoF; i++)
                   {
                       SMatrix.AddRow(1, 1);
                       SMatrix.Columns.Item(1).Cells.Item(i).Specific.Value = fields.Item("U_codArt").Value.ToString();
                       SMatrix.Columns.Item(2).Cells.Item(i).Specific.Value = fields.Item("U_articulo").Value.ToString();
                       SMatrix.Columns.Item(3).Cells.Item(i).Specific.Value = fields.Item("U_codCli").Value.ToString();
                       SMatrix.Columns.Item(4).Cells.Item(i).Specific.Value = fields.Item("U_cliente").Value.ToString();
                       SMatrix.Columns.Item(5).Cells.Item(i).Specific.Value = fields.Item("U_cant").Value.ToString();
                       SMatrix.Columns.Item(8).Cells.Item(i).Specific.Value = fields.Item("Code").Value.ToString();
                       
                       rsCards.MoveNext();
                   }
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

        private bool Cargar_Solicitud(string noDoc, bool posicion)
        {

            bool todoOk = true;

            string serror = "";

            SAPbouiCOM.DBDataSource oDbHeaderDataSource = null;
            oDbHeaderDataSource = SForm.DataSources.DBDataSources.Item("@CAB_RSTV");
            if (oDbHeaderDataSource.Size == 0)
            {
                return Insertar_Solicitud();
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

                        todoOk = Cargar_Lineas(txt_numoc.Value.ToString());
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
                return User_Code ;
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

                String strSQL = String.Format("SELECT  MAX({0}) AS nd  FROM {1}",
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

                String strSQL = String.Format("SELECT MAX({0}) AS nl FROM {1}",
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

        public void FillCombo(SAPbouiCOM.ComboBox oCombo, string SqlQuery)
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


        
    }
}
