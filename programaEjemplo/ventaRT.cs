using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SSIFramework;
using SSIFramework.Plugins;
using SSIFramework.Utilidades;
using SAPbobsCOM;
using ventaRT.Constantes.View;

namespace ventaRT
{

    public class addonGeneral : IPlugin
    {

        SAPbouiCOM.Form SForm = null;
        SAPbouiCOM.Matrix SMatrix = null;
        SAPbouiCOM.Form UForm = null;
        SAPbouiCOM.Matrix UMatrix = null;

        private SSIFramework.SSIConnector B1;

        public string Guid
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public string Nombre
        {
            get
            {
                return "ventaRT";
            }
        }

        public string Version
        {
            get
            {
                return "0.1";
            }
        }

        public void CambioCompañia()
        {
            throw new NotImplementedException();
        }

        public void CambioIdioma()
        {
            throw new NotImplementedException();
        }

        public void CrearMenu()
        {
            SSIConnector B1 = SSIConnector.GetSSIConnector();
            try
            {

                if (!B1.Application.Menus.Exists(Constantes.Views.Menu.MenuVentaReserva))
                    GenericFunctions.addMenu(Constantes.Views.Menu.MenuVentaReserva,
                     Constantes.Views.Menu.MenuVentaReserva_Desc, "43520",
                      BoMenuType.mt_POPUP, null);

                if (!B1.Application.Menus.Exists(Constantes.Views.Menu.MENU_submenu_registro_solicitud))
                    GenericFunctions.addMenu(Constantes.Views.Menu.MENU_submenu_registro_solicitud,
                    Constantes.Views.Menu.MENU_submenu_registro_solicitud_Desc,
                    Constantes.Views.Menu.MenuVentaReserva,
                    BoMenuType.mt_STRING, null);

                if (!B1.Application.Menus.Exists(Constantes.Views.Menu.MENU_submenu_control_aprobaciones))
                    GenericFunctions.addMenu(Constantes.Views.Menu.MENU_submenu_control_aprobaciones,
                    Constantes.Views.Menu.MENU_submenu_control_aprobaciones_Desc,
                    Constantes.Views.Menu.MenuVentaReserva,
                    BoMenuType.mt_STRING, null);

                //if (!B1.Application.Menus.Exists(Constantes.Views.Menu.MENU_submenu_control_anulaciones))
                //    GenericFunctions.addMenu(Constantes.Views.Menu.MENU_submenu_control_anulaciones,
                //    Constantes.Views.Menu.MENU_submenu_control_anulaciones_Desc,
                //    Constantes.Views.Menu.MenuVentaReserva,
                //    BoMenuType.mt_STRING, null);

                if (!B1.Application.Menus.Exists(Constantes.Views.Menu.MENU_submenu_autorizadores))
                    GenericFunctions.addMenu(Constantes.Views.Menu.MENU_submenu_autorizadores,
                    Constantes.Views.Menu.MENU_submenu_autorizadores_Desc,
                    Constantes.Views.Menu.MenuVentaReserva,
                    BoMenuType.mt_STRING, null);
            }
            catch (Exception ex)
            {
                B1.Application.MessageBox("Error : " + ex.Message);
            }
        }

        public void Finalizar()
        {
            throw new NotImplementedException();
        }

        public void Instalar()
        {
            DataBase.UserFields.CrearEstructura();


        }

        public bool PreInstalar()
        {
            throw new NotImplementedException();
        }

        public void Run()
        {
            // *******************************************************************
            //  Use SSIFramework.SSIConnector object to establish connection
            //  with the SAP Business One application and return an
            //  initialized appliction object
            // *******************************************************************
            B1 = SSIConnector.GetSSIConnector();
            GenericFunctions.GetResourceForm();
            Global.MenuEvent += new Global.MenuEventHandler(Global_MenuEvent);
            B1.Application.SetStatusBarMessage("Addon Iniciado correctamente...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
        }

        void Global_MenuEvent(ref MenuEvent pVal, ref bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case Constantes.Views.Menu.MENU_submenu_registro_solicitud:
                            {
                                B1.Application.SetStatusBarMessage("Abriendo menu...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                new VIEW.PantallaRegistro(null);
                                Configurar_Pantalla_Registro();
                            }

                            break;
                        case Constantes.Views.Menu.MENU_submenu_control_aprobaciones:
                            {
                                if(es_Autorizador())
                                {
                                    B1.Application.SetStatusBarMessage("Abriendo menu...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    new VIEW.PantallaAprobac();

                                }
                                else
                                {
                                    //B1.Application.SetStatusBarMessage("Ud. no está registrado como Autorizador, por tanto, no puede acceder a esta opción", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                    int resp;
                                    resp=B1.Application.MessageBox("Ud. no está registrado como Autorizador, por tanto, no puede acceder a esta opción");
                                }
                            }

                            break;
                        case Constantes.Views.Menu.MENU_submenu_autorizadores:
                            {
                                B1.Application.SetStatusBarMessage("Abriendo menu...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                new VIEW.PantallaAutoriz();
                                Configurar_Pantalla_Autoriz();
                            }

                            break;
                    }
                }


            }
            catch (Exception ex)
            {
                B1.Application.MessageBox("Error : " + ex.Message);
                bubbleEvent = false;
            }
        }

        void Global_ItemEvent(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            SAPbouiCOM.Form oForm;

            // *************************************************************************
            //  BubbleEvent sets the behavior of SAP Business One.
            //  False means that the application will not continue processing this event.
            // *************************************************************************
            bubbleEvent = true;

            try
            {
                if (!pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {

                        //case BoEventTypes.et_FORM_LOAD:

                        //    if (pVal.Action_Success)
                        //    {

                        //        switch (pVal.FormTypeEx)
                        //        {

                        //        }
                        //    }
                        //    break;
                        case BoEventTypes.et_CHOOSE_FROM_LIST:

  
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error al llegar al modulo principal");
                bubbleEvent = false;
                throw ex;


            }

        }

        
        private void AddChooseFromListToEditTextBox(SAPbouiCOM.Form XForm, string ObjectType,
            string CFLUID, SAPbobsCOM.BoYesNoEnum Condition, string ConAlias = "" ,
            string conVal = "", string oper = ""   )
        {
            if (XForm!= null)
            {
                try
                {
                    SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                    SAPbouiCOM.ChooseFromList oCFL = null;
                    SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                    SAPbouiCOM.Conditions oCons = null;
                    SAPbouiCOM.Condition oCon = null;
                    oCFLs = XForm.ChooseFromLists;

                    oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                        (B1.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams));
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = ObjectType;
                    oCFLCreationParams.UniqueID = CFLUID;
                    oCFL = oCFLs.Add(oCFLCreationParams);

                    if (Condition == BoYesNoEnum.tYES)
                    {
                        oCons = oCFL.GetConditions();
                        oCon = oCons.Add();
                        oCon.Alias = ConAlias;
                        oCon.Operation = oper == "=" ? SAPbouiCOM.BoConditionOperation.co_EQUAL :
                            oper == ">" ? SAPbouiCOM.BoConditionOperation.co_GRATER_THAN :
                            oper == "<" ? SAPbouiCOM.BoConditionOperation.co_LESS_THAN :
                             SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = conVal;
                        oCFL.SetConditions(oCons);
                    }
                }
                catch (Exception ex)
                {
                    B1.Application.MessageBox("Error : " + ex.Message);
                }
            }

        }

        private void AddCFLArtOnHandinCD(string ObjectType,string CFLUID)
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;
                oCFLs = SForm.ChooseFromLists;

                oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                    (B1.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = ObjectType;
                oCFLCreationParams.UniqueID = CFLUID;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCons = oCFL.GetConditions();
                //Ejecutar Query que devuelve los que tienen existencia en la Bodega CD
                String strSQL = String.Format("SELECT {1} FROM {3} " +
                         " WHERE {2}='{4}' AND {0} > 0 ",
                               Constantes.View.oitw.OnHand,
                               Constantes.View.oitw.ItemCode,
                               Constantes.View.oitw.WhsCode,
                               Constantes.View.oitw.OITW,
                               "CD");
                Recordset rsResult = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsResult.DoQuery(strSQL);
                SAPbobsCOM.Fields fields = rsResult.Fields;
                rsResult.MoveFirst();
                if (!rsResult.EoF)
                {
                    do {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = rsResult.Fields.Item("ItemCode").Value.ToString();
                        rsResult.MoveNext();
                        if (!rsResult.EoF)
                        {
                            oCon.Relationship = BoConditionRelationship.cr_OR;
                        }
                    } while (!rsResult.EoF);

                }
                oCFL.SetConditions(oCons);

            }
            catch (Exception ex)
            {
                B1.Application.MessageBox("Error : " + ex.Message);
            }
        }

        private void Configurar_Pantalla_Registro()
        {

            SForm = B1.Application.Forms.ActiveForm;
            SMatrix =SForm.Items.Item("mtx" ).Specific;

            SAPbouiCOM.Column _Col = (SAPbouiCOM.Column)SMatrix.Columns.Item("codArt");
            SAPbouiCOM.Column _Col1 = (SAPbouiCOM.Column)SMatrix.Columns.Item("articulo");
            AddCFLArtOnHandinCD("4", "CFL1");                                                           // 4 - oitm

            SAPbouiCOM.Column _Col2= (SAPbouiCOM.Column)SMatrix.Columns.Item("codCli");
            SAPbouiCOM.Column _Col3 = (SAPbouiCOM.Column)SMatrix.Columns.Item("cliente");
            AddChooseFromListToEditTextBox(SForm, "2", "CFL2", BoYesNoEnum.tYES, "CardType", "C", "="); // 2 - ocrd

            _Col.ChooseFromListUID = "CFL1";
            _Col.ChooseFromListAlias = "ItemCode";
            _Col1.Editable = false;

            _Col2.ChooseFromListUID = "CFL2";
            _Col2.ChooseFromListAlias = "CardCode";
            _Col3.Editable = false;    
        }

        private void Configurar_Pantalla_Autoriz()
        {

            UForm = B1.Application.Forms.ActiveForm;
            UMatrix = UForm.Items.Item("umtx").Specific;

            SAPbouiCOM.Column _Col4 = (SAPbouiCOM.Column)UMatrix.Columns.Item("idAut");
            SAPbouiCOM.Column _Col5 = (SAPbouiCOM.Column)UMatrix.Columns.Item("aut");
            AddChooseFromListToEditTextBox(UForm, "12", "CFL3", BoYesNoEnum.tNO);                       // 12 - ousr

            _Col4.ChooseFromListUID = "CFL3";
            _Col4.ChooseFromListAlias = "USER_CODE";
            _Col5.Editable = false;
        }

        private bool es_Autorizador()
        {
            try
            {
                string usrCurrent = B1.Company.UserName;
                String strSQL = String.Format("SELECT COUNT(*) FROM {1} Where {0}='{3}' AND {2} = 'Y'",
                          Constantes.View.AUT_RVT.U_idAut,
                          Constantes.View.AUT_RVT.AUT_RV,
                          Constantes.View.AUT_RVT.U_activo,
                          usrCurrent);
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
                B1.Application.SetStatusBarMessage("Error verificando Autorizador", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
        }
    
        private bool existe_Form(string ftype)
        {
            bool existe = false;
            try
            {
                for (int i = 0; i < B1.Application.Forms.Count && !existe; i++)
                {
                    existe = B1.Application.Forms.Item(i).TypeEx == ftype;
                    if (existe) { B1.Application.Forms.Item(i).Select(); }
                }
            }
            catch (Exception ex)
            {
                B1.Application.SetStatusBarMessage("Error verificando formulario", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                throw ex;
            }
            return existe;
        }
    }
}
