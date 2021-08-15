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
       

        public PantallaAutoriz()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.Autorizad.srf"], "AutRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            formActual = "AutRT" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
       
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
                        case "1292":   //ADICIONAR LINEA
                            switch (ItemActiveMenu)
                            {
                                case ventaRT.Constantes.View.autorizad.umtx:
                                    insertar_linea_autoriz();
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                        case "1293":  //BORRAR LINEA
                            switch (ItemActiveMenu)
                            {
                                //ejemplo con una matrix 
                                case ventaRT.Constantes.View.autorizad.umtx:
                                    borrar_linea_autoriz();
                                    BubbleEvent = false;
                                    break;
                            }
                            break;
                    }
                    BubbleEvent = true;
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
                    if (eventInfo.BeforeAction && eventInfo.ItemUID == ventaRT.Constantes.View.autorizad.umtx)
                    {
                        UForm.EnableMenu("1292", true); //Activar Agregar Linea
                        UForm.EnableMenu("1293", true); //Activar Borrar Linea 
                        rowsel = eventInfo.Row;
                    }
                    else
                    {
                        UForm.EnableMenu("1292", false); //Desctivar Agregar Linea
                        UForm.EnableMenu("1293", false); //Desactivar Borrar Linea 
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
                                        case Constantes.View.autorizad.btn_Exit:
                                            {
                                                UForm.Close();
                                            }
                                            break;
                                    }
                                    break;

                                }

                            //case BoEventTypes.et_VALIDATE:
                            //    {
                            //        if (pVal.InnerEvent == false && pVal.ItemUID == "umtx" && pVal.ColUID == "activo")
                            //        {
                            //            string idAut = ((SAPbouiCOM.EditText)UMatrix.Columns.Item("idAut").Cells.Item(pVal.Row).Specific).Value.ToString();

                            //            if (idAut != "" && pVal.Row == UMatrix.RowCount)
                            //            {
                            //                UMatrix.AddRow(1, pVal.Row);
                            //                UMatrix.ClearRowData(UMatrix.RowCount);
                            //                UMatrix.FlushToDataSource();
                            //                UMatrix.LoadFromDataSource();
                            //            }
                            //        }
                            //    }
                            //    break;

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
                                                if (usrsel != "" && !validar_usr_unico(usrsel, pVal.Row))
                                                {
                                                    Ok = false;
                                                    B1.Application.SetStatusBarMessage("Error Usuario Repetido", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
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
                                        guardar_autoriz();
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
                                                        B1.Application.SetStatusBarMessage("Error Codigo Autorizador es Obligatorio", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                                        BubbleEvent = false;
                                                    }
                                                    else
                                                    {
                                                        if (idaut != "" && !validar_usr_unico(idaut, pVal.Row))
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
                B1.Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }

        }

         
        // Metodos No Override

        private void cargar_info_inicial()
        {
            UForm = B1.Application.Forms.ActiveForm;
            UMatrix = UForm.Items.Item("umtx").Specific;
            oDbAutDataSource = UForm.DataSources.DBDataSources.Item("@AUT_RSTV");
            formActual = B1.Application.Forms.ActiveForm.UniqueID;
            cargar_lineas();
        }

        private bool guardar_autoriz()
        {
            bool todoOk = true;
            string serror = "";
            string sCode = ""; string sName = "";
            int iRet;
            UForm.Freeze(true);
            try
            {
                SAPbobsCOM.UserTable UTAut = B1.Company.UserTables.Item("AUT_RSTV");
                //Salvando autorizadores
                if (UMatrix != null)
                {
                    int norecord = obtener_ultimo_ID() ;
  
                    UMatrix.FlushToDataSource();
                    for(int i=0; i <= oDbAutDataSource.Size-1; i++)
                    {

                        // Obteniendo texto de los campos de DbDataSource
                        string sCodeL = oDbAutDataSource.GetValue("Code", i);
                        string sNameL = oDbAutDataSource.GetValue("Name" ,i);
                        string scodaut = oDbAutDataSource.GetValue("U_idAut",i);
                        string saut = oDbAutDataSource.GetValue("U_aut",i);
                        string sactivo = oDbAutDataSource.GetValue("U_activo", i);

                        if (scodaut != "")
                        {
                            try
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
                                    todoOk = (iRet == 0);
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
                    UTAut = null;
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

                System.GC.Collect();
            }
            if (todoOk)
            {
                todoOk = eliminar_filas_borradas();
            }
            if (todoOk){
               B1.Application.SetStatusBarMessage("Datos guardados exitosamente", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            else {
                B1.Application.SetStatusBarMessage("Error guardando datos: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            UForm.Freeze(false);
            return todoOk;
        }

        private bool eliminar_filas_borradas()
        {
            bool todoOk = true;
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

        private bool cargar_lineas()
        {
           bool todoOk = true;

           string serror = "";
               try
               {
                   UForm.Freeze(true);
                   oDbAutDataSource.Query();
                   UMatrix.LoadFromDataSource();
                   UMatrix.AutoResizeColumns();
                   SAPbouiCOM.Column oColumn = UMatrix.Columns.Item("idAut");
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
                   B1.Application.SetStatusBarMessage("Datos cargados con exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
               }
               else
               {
                   B1.Application.SetStatusBarMessage("Error cargando datos: " + serror, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
               }

               UForm.Freeze(false);
               return todoOk;
        }

        private bool validar_usr_unico( string usr, int row)
        {
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
                    B1.Application.SetStatusBarMessage("Error validando datos repetidos" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    todoOK = false;
                    throw;
                }
            }

            return todoOK;
        }

        private int  obtener_ultimo_ID()
        {
            int CodeNum = 0;
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
            return CodeNum;
        }

        private bool tiene_Autorizadas(string autor)
        {
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
                B1.Application.SetStatusBarMessage("Error obteniendo Autorizaciones", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw;
            }
        }

        private void borrar_linea_autoriz()
        {
            UForm.Freeze(true);
            //int nRow = (int)UMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            //nRow = nRow == -1 ? UMatrix.RowCount : nRow ;
            if (rowsel > 0)
            {
                UMatrix.GetLineData(rowsel);
                //  Verificando si tiene autorizadas
                string autor = oDbAutDataSource.GetValue("U_idAut", rowsel - 1);
                if (tiene_Autorizadas(autor))
                {

                    B1.Application.SetStatusBarMessage("Ese Autorizador tiene autorizaciones, por tanto, solo se Desactiva", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    UMatrix.Columns.Item(3).Cells.Item(rowsel).Specific.Checked = false;
                    UMatrix.FlushToDataSource();
                    UMatrix.LoadFromDataSource();
                }
                else
                {
                    string lindel = oDbAutDataSource.GetValue("code", rowsel - 1);
                    lineasdel.Add(lindel);
                    UMatrix.DeleteRow(rowsel);
                    UMatrix.FlushToDataSource();
                    UMatrix.LoadFromDataSource();
                }
            }
            UForm.Freeze(false);
        }

        private void insertar_linea_autoriz()
        {
            UMatrix.AddRow(1, UMatrix.RowCount);
            UMatrix.ClearRowData(UMatrix.RowCount);
            UMatrix.FlushToDataSource();
            UMatrix.LoadFromDataSource();
            UMatrix.Columns.Item(3).Cells.Item(UMatrix.RowCount).Specific.Checked = true;
            UMatrix.Columns.Item(1).Cells.Item(UMatrix.RowCount).Click(BoCellClickType.ct_Double);
        }
        

    }
}
