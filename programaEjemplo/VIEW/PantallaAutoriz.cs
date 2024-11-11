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
    class PantallaAutoriz : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        private string ItemActiveMenu = "";

        private string formActual = "";
        SAPbouiCOM.Form UForm = null;
        SAPbouiCOM.Matrix UMatrix = null;
        SAPbouiCOM.DBDataSource oDbAutDataSource = null;

        List<string> lineasdel = new List<string>();
        int rowsel = 0;
        private string msgError = "";

        public PantallaAutoriz()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Autorizad.srf"], "AutRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            string errorMessage = "";
            formActual = "AutRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
       
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
            this.B1.Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(ThisSapApiForm_OnAfterRightClick);

            errorMessage =  cargar_info_inicial();
            if (!string.IsNullOrEmpty(errorMessage)){ HandleError(new Exception(errorMessage));}
        }


        // Metodos Override

        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            string errorMessage = "";
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
                        case "1292":   //ADICIONAR LINEA
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.autorizad.umtx:
                                    errorMessage = insertar_linea_autoriz();
                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                        case "1293":  //BORRAR LINEA
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.autorizad.umtx:
                                    errorMessage = borrar_linea_autoriz();
                                    if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                    }
                    //BubbleEvent = true;
                }
             }
            }

            catch (Exception ex)
            {
                msgError = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
                B1.Application.SetStatusBarMessage("Error: " + msgError, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                    if (eventInfo.BeforeAction && eventInfo.ItemUID == ventaRT.Constantes.View.autorizad.umtx)
                    {
                        rowsel = eventInfo.Row;
                    }
                }
            }

            catch (Exception ex)
            {
                msgError = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
                B1.Application.SetStatusBarMessage("Error: " + msgError, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                            case BoEventTypes.et_ITEM_PRESSED:
                                {
                                    switch (pVal.ItemUID)
                                    {
                                        case Constantes.View.autorizad.btn_Exit:
                                            {
                                                UForm.Close();
                                            }
                                            break;
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
                                        oCFL = UForm.ChooseFromLists.Item(CFL_Id);
                                        if (pVal.FormTypeEx.Substring(0, 5) == "AutRT" && CFLEvent.SelectedObjects != null)
                                        {
                                            if (pVal.ItemUID == "umtx" && pVal.ColUID == "idAut")
                                            {
                                                bool Ok = true;
                                                string usrsel = CFLEvent.SelectedObjects.GetValue("USER_CODE", 0).ToString();
                                                // Validar que no existan repetidos 
                                                bool is_unique = validar_usr_unico(usrsel, pVal.Row, out errorMessage);
                                                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                if (usrsel != "" && !is_unique)
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error: Autorizador Repetido", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                                    BubbleEvent = false;
                                                }
                                                if (Ok)
                                                {
                                                    int nRow = (int)UMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                                    nRow = nRow == -1 ? pVal.Row : nRow - 1;
                                                    UMatrix.FlushToDataSource();
                                                    oDbAutDataSource.SetValue("U_idAut", nRow - 1, usrsel);
                                                    oDbAutDataSource.SetValue("U_aut", nRow - 1, CFLEvent.SelectedObjects.GetValue("U_NAME", 0).ToString());
                                                    UMatrix.LoadFromDataSource();
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
                            case BoEventTypes.et_FORM_CLOSE:
                                {
                                    if (pVal.FormTypeEx.Substring(0, 5) == "AutRT")
                                    {
                                        errorMessage =  guardar_autoriz();
                                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                    }
                                }
                                break;

                            case BoEventTypes.et_VALIDATE:
                                {
                                    if (pVal.InnerEvent == false && pVal.ItemUID == "umtx")
                                    {
                                        string idaut = ((SAPbouiCOM.EditText)UMatrix.Columns.Item("idAut").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        switch (pVal.ColUID)
                                        {
                                            case "idAut":
                                                {
                                                    if (idaut == "")
                                                    {
                                                        B1.Application.SetStatusBarMessage("Error: Código Autorizador es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        bool is_unique = validar_usr_unico(idaut, pVal.Row, out errorMessage);
                                                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                                                        if (idaut != "" && !is_unique)
                                                        {
                                                            B1.Application.SetStatusBarMessage("Error: Autorizador Repetido", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                msgError = (B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message;
                B1.Application.SetStatusBarMessage("Error: " + msgError, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;

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
                for (int i = 1; i <= UMatrix.RowCount; i++)
                {
                    UMatrix.Columns.Item(0).Cells.Item(i).Specific.Value = i.ToString();
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Numerar Líneas: " + ((B1.Company.GetLastErrorCode() != 0) 
                    ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }

        private string cargar_lineas()
        {
            string errorMessage = "";
            try
            {
                UForm.Freeze(true);
                oDbAutDataSource.Query();
                UMatrix.LoadFromDataSource();
                UMatrix.AutoResizeColumns();
                SAPbouiCOM.Column oColumn = UMatrix.Columns.Item("idAut");
                oColumn.TitleObject.Sort(BoGridSortType.gst_Ascending);
                errorMessage = num_lineas();
                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
            }
            catch (Exception ex)
            {
                errorMessage = "Cargando Líneas: " +  
                    ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                UForm.Freeze(false);
            }
            return errorMessage;
        }

        private string cargar_info_inicial()
        {
            string errorMessage = "";
            try
            {
                UForm = B1.Application.Forms.ActiveForm;
                UMatrix = UForm.Items.Item("umtx").Specific;
                oDbAutDataSource = UForm.DataSources.DBDataSources.Item("@AUT_RSTV");
                formActual = B1.Application.Forms.ActiveForm.UniqueID;
                bool isAdmin = es_Admin(out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                UForm.EnableMenu("1292", isAdmin); //Activar Agregar Linea
                UForm.EnableMenu("1293", isAdmin); //Activar Borrar Linea

                errorMessage = cargar_lineas();
                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
            }
            catch (Exception ex)
            {
                errorMessage = "Cargando Datos: " + ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return errorMessage;
        }
        
        private string guardar_autoriz()
        {
            string errorMessage = "";
            int iRet = 0;
            try
            {
                UForm.Freeze(true);
                SAPbobsCOM.UserTable UTAut = B1.Company.UserTables.Item("AUT_RSTV");
                //Salvando autorizadores
                if (UMatrix != null)
                {
                    int norecord = obtener_ultimo_ID(out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        UForm.Freeze(false);
                        return errorMessage; 
                    }
  
                    UMatrix.FlushToDataSource();
                    for(int i=0; i <= oDbAutDataSource.Size-1; i++)
                    {
                        // Obteniendo texto de los campos de DbDataSource
                        string sCodeL = oDbAutDataSource.GetValue("Code", i);
                        string sNameL = oDbAutDataSource.GetValue("Name" ,i);
                        string scodaut = oDbAutDataSource.GetValue("U_idAut",i);
                        string saut = oDbAutDataSource.GetValue("U_aut",i);
                        string sactivo = oDbAutDataSource.GetValue("U_activo", i);
                        iRet = 0;
                        if (scodaut != "")
                        {
                            // Guardando en la UserTable
                            B1.Company.StartTransaction();
                            if (UTAut.GetByKey(sCodeL))
                            {
                                //UPDATE
                                UTAut.UserFields.Fields.Item("U_idAut").Value = scodaut;
                                UTAut.UserFields.Fields.Item("U_aut").Value = saut;
                                UTAut.UserFields.Fields.Item("U_activo").Value = sactivo;
                                iRet = UTAut.Update();
                            }
                            else
                            {
                                //INSERT
                                norecord = norecord + 1;
                                sCodeL = norecord.ToString();
                                UTAut.Code = sCodeL;
                                UTAut.Name = sCodeL;
                                UTAut.UserFields.Fields.Item("U_idAut").Value = scodaut;
                                UTAut.UserFields.Fields.Item("U_aut").Value = saut;
                                UTAut.UserFields.Fields.Item("U_activo").Value = sactivo;
                                iRet = UTAut.Add();
                            }
                            if (iRet != 0)
                            {
                                if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                                errorMessage = "Guardar Autorizadores" + 
                                    ((B1.Company.GetLastErrorCode() != 0) 
                                    ? B1.Company.GetLastErrorDescription() 
                                    : "");
                                return errorMessage;
                            }
                            if (B1.Company.InTransaction) { B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); }
                        }
                    }
                    UTAut = null;
                }
                errorMessage = eliminar_filas_borradas();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    UForm.Freeze(false);
                    return errorMessage;
                }
                B1.Application.SetStatusBarMessage("Datos de Autorizadores guardados con éxito...", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                UForm.Freeze(false);
            }
             catch (Exception ex)
            {
                if (B1.Company.InTransaction)
                {
                    B1.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                errorMessage = "Guardar Autorizadores:" + 
                    ((B1.Company.GetLastErrorCode() != 0) 
                    ? B1.Company.GetLastErrorDescription() 
                    : ex.Message);
            }
            finally {
                System.GC.Collect();
            }
            return errorMessage;
        }

        private string eliminar_filas_borradas()
        {
            string errorMessage = "";
            string SQLQuery = String.Empty;
            try
            {
                UMatrix.LoadFromDataSource();
                if (lineasdel !=null)
                {
                    for (int i = 0; i < lineasdel.Count ; i++)
                    {
                        SQLQuery = String.Format("DELETE FROM {1} WHERE {0} = '{2}' ",
                                        Constantes.View.AUT_RVT.Code,
                                        Constantes.View.AUT_RVT.AUT_RV,
                                        lineasdel[i]);
                        Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsCards.DoQuery(SQLQuery);
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Sincronizar filas eliminadas: " + 
                    ((B1.Company.GetLastErrorCode() != 0) 
                    ? B1.Company.GetLastErrorDescription() 
                    : ex.Message);
            }
            finally
            {
                lineasdel.Clear();
                System.GC.Collect();
            }
            return errorMessage;
        }

        private bool validar_usr_unico( string usr, int row, out string errorMessage)
        {
            errorMessage = "";
            bool todoOK = true;
            if(UMatrix.RowCount > 1)
            {
                try
                {
                    // Validar contra la misma matriz porque cuando es nuevo solo datos en linea, 
                    // No fisicos en la BD
                    int creg = 0;
                    for (int i = 1; i <= UMatrix.RowCount && creg < 1; i++)
                    {
                        if ((i != row) &&
                            (UMatrix.Columns.Item(1).Cells.Item(i).Specific).Value.ToString() == usr)
                        {
                            creg++;
                        }
                    }
                    todoOK = (creg < 1);
                }
                catch (Exception ex)
                {
                    todoOK = false;
                    errorMessage = "Error obteniendo Autorizaciones: " +
                        ((B1.Company.GetLastErrorCode() != 0)
                        ? B1.Company.GetLastErrorDescription()
                        : ex.Message);
                  }
            }
            return todoOK;
        }

        private int  obtener_ultimo_ID(out string errorMessage)
        {
            errorMessage = "";
            int CodeNum = 0;
            try
            {
                
                String strSQL = String.Format("SELECT TOP 1 CAST(T0.{0} AS INT) AS nd FROM {1} T0 ORDER BY CAST(T0.{0} AS INT) DESC",
                                        Constantes.View.AUT_RVT.Code,
                                        Constantes.View.AUT_RVT.AUT_RV);
                Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsCards.DoQuery(strSQL);
                string Code = rsCards.Fields.Item("nd").Value.ToString();
                if (Code != "")
                {
                    CodeNum = Convert.ToInt32(Code);
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Error Obtener nuevo ID: " +
                    ((B1.Company.GetLastErrorCode() != 0)
                    ? B1.Company.GetLastErrorDescription()
                    : ex.Message);
              }
            return CodeNum;
        }

        private bool tiene_Autorizadas(string autor, out string errorMessage)
        {
            errorMessage = "";
            try
            {
                string usrCurrent = B1.Company.UserName;
                String strSQL = String.Format("SELECT COUNT(*) FROM {1} Where {0}='{2}'",
                          Constantes.View.CAB_RVT.U_idAut,
                          Constantes.View.CAB_RVT.CAB_RV,
                          autor);
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
                    int existe = Int32.Parse(rsUsers.Fields.Item("COUNT(*)").Value.ToString());
                    return existe > 0;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Error obteniendo Autorizaciones: " +  
                    ((B1.Company.GetLastErrorCode() != 0) 
                    ? B1.Company.GetLastErrorDescription() 
                    : ex.Message);
             }
            return false;
        }

        private string borrar_linea_autoriz()
        {
            string errorMessage = "";
            try
            {
                UForm.Freeze(true);
                if (rowsel > 0)
                {
                    UMatrix.GetLineData(rowsel);
                    //  Verificando si tiene autorizadas
                    string autor = UMatrix.Columns.Item(1).Cells.Item(rowsel).Specific.Value.ToString();
                    bool tieneSolAutorizadas = tiene_Autorizadas(autor, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) { return errorMessage; }
                    if (tieneSolAutorizadas)
                    {
                        B1.Application.SetStatusBarMessage("Ese Autorizador tiene solicitudes autorizadas, por tanto, solo se Desactiva", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        UMatrix.Columns.Item(3).Cells.Item(rowsel).Specific.Checked = false;
                        UMatrix.FlushToDataSource();
                        UMatrix.LoadFromDataSource();
                    }
                    else
                    {
                        string lindel =  UMatrix.Columns.Item(4).Cells.Item(rowsel).Specific.Value.ToString();
                        lineasdel.Add(lindel);
                        UMatrix.DeleteRow(rowsel);
                        UMatrix.FlushToDataSource();
                        UMatrix.LoadFromDataSource();
                        errorMessage = num_lineas();
                        if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Borrar Línea: " +
                    ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                UForm.Freeze(false);
            }
            return errorMessage;
        }

        private string insertar_linea_autoriz()
        {
            string errorMessage = "";
            try
            {
                UMatrix.AddRow(1, UMatrix.RowCount);
                UMatrix.ClearRowData(UMatrix.RowCount);
                UMatrix.FlushToDataSource();
                UMatrix.LoadFromDataSource();
                errorMessage = num_lineas();
                if (!string.IsNullOrEmpty(errorMessage)) { HandleError(new Exception(errorMessage)); }
                UMatrix.Columns.Item(3).Cells.Item(UMatrix.RowCount).Specific.Checked = true;
                UMatrix.Columns.Item(1).Cells.Item(UMatrix.RowCount).Click(BoCellClickType.ct_Double);

            }
            catch (Exception ex)
            {
                errorMessage = "Adicionar Línea: " +
                    ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            finally
            {
                UForm.Freeze(false);
            }
            return errorMessage;
        }

        private bool es_Admin(out string errorMessage)
        {
            errorMessage = "";
            try
            {
                string usrCurrent = B1.Company.UserName;

                String strSQL = String.Format("SELECT COUNT(*) FROM {1} Where {0}='{3}' AND {2} = 'N'",
                          Constantes.View.ousr.uCode,  //0
                          Constantes.View.ousr.OUSR,       //1
                          Constantes.View.ousr.uLocked,     //2
                          usrCurrent);                     //3
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
                    int existe = Int32.Parse(rsUsers.Fields.Item("COUNT(*)").Value.ToString());
                    return existe > 0;
                }
            }
            catch (Exception ex)
            {
                errorMessage =  "Verificar Administrador: " +  ((B1.Company.GetLastErrorCode() != 0) ? B1.Company.GetLastErrorDescription() : ex.Message);
            }
            return false;
        }

    }
}
