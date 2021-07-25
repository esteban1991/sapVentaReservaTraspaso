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

                if (!B1.Application.Menus.Exists(Constantes.Views.Menu.MENU_submenu_control_anulaciones))
                    GenericFunctions.addMenu(Constantes.Views.Menu.MENU_submenu_control_anulaciones,
                    Constantes.Views.Menu.MENU_submenu_control_anulaciones_Desc,
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
                                new VIEW.PantallaRegistro();
                                Configurar_Pantalla_Registro();
                            }

                            break;
                        case Constantes.Views.Menu.MENU_submenu_control_aprobaciones:
                            {
                                B1.Application.SetStatusBarMessage("Abriendo menu...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                new VIEW.PantallaAprobac();
                                //Configurar_Pantalla_Registro();
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

        
        private void AddChooseFromListToEditTextBox(string ObjectType,
            string CFLUID, SAPbobsCOM.BoYesNoEnum Condition, string ConAlias = "" ,
            string conVal = "", string oper = ""  )
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
                
                if (Condition == BoYesNoEnum.tYES)
                {
                    oCons = oCFL.GetConditions();
                    oCon = oCons.Add();
                    oCon.Alias = ConAlias;
                    oCon.Operation = oper == "="? SAPbouiCOM.BoConditionOperation.co_EQUAL: 
                        oper == ">"? SAPbouiCOM.BoConditionOperation.co_GRATER_THAN: 
                        oper == "<"? SAPbouiCOM.BoConditionOperation.co_LESS_THAN :
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

        private void Configurar_Pantalla_Registro()
        {

            SForm = B1.Application.Forms.ActiveForm;
            SMatrix =SForm.Items.Item("mtx" ).Specific;
            //SForm.AutoManaged = true;
            //SForm.Items.Item("txt_numoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,0,SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            //SForm.Items.Item("txt_numoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
             //   2,SAPbouiCOM.BoModeVisualBehavior.mvb_False);
             // SForm.DataBrowser.BrowseBy = "txt_numoc" ;
              //SForm.Items.Item("txt_numoc").Click();

            //SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)SForm.Items.Item("txt_numoc").Specific;
            //SAPbouiCOM.EditText oEditText2 = (SAPbouiCOM.EditText)SForm.Items.Item("txt_com").Specific;
            //SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)SForm.Items.Item("cbnumoc").Specific;
            //LoadDefaultValue("@CAB_RT", ref oComboBox, ref oEditText, ref oEditText2);

            SAPbouiCOM.Column _Col = (SAPbouiCOM.Column)SMatrix.Columns.Item("codArt");
            SAPbouiCOM.Column _Col1 = (SAPbouiCOM.Column)SMatrix.Columns.Item("articulo");
            //AddChooseFromListToEditTextBox("4", "CFL1", BoYesNoEnum.tNO);
            AddChooseFromListToEditTextBox("4", "CFL1", BoYesNoEnum.tYES, "onHand", "0", ">");

            SAPbouiCOM.Column _Col2= (SAPbouiCOM.Column)SMatrix.Columns.Item("codCli");
            SAPbouiCOM.Column _Col3 = (SAPbouiCOM.Column)SMatrix.Columns.Item("cliente");
            AddChooseFromListToEditTextBox("2", "CFL2", BoYesNoEnum.tYES,"CardType","C" ,"=" );

            _Col.ChooseFromListUID = "CFL1";
            _Col.ChooseFromListAlias = "ItemCode";
            _Col1.Editable = false;

            _Col2.ChooseFromListUID = "CFL2";
            _Col2.ChooseFromListAlias = "CardCode";
            _Col3.Editable = false;    
        }
    }
}
