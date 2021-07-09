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

        //ejemplos de constantes donde se utilizar la key del menu
        //public const string Interlocutors = "134";
        //public const string Articles = "150";
        //public const string Journal = "392";
        public const string CdeCostes = "810";
        public const string CdeCostesUf = "-810";
        //public const string ATP = "154";

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
                            B1.Application.SetStatusBarMessage("Abriendo menu...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            new VIEW.PantallaRegistro();
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

                        case BoEventTypes.et_FORM_LOAD:

                            if (pVal.Action_Success)
                            {

                                switch (pVal.FormTypeEx)
                                {




                                    case CdeCostes:
                                        String strSQL = String.Format("SELECT {0},{1}  FROM {2} WHERE {3} = '2'",
                                             Constantes.Views.ColGridLog.PrCode,
                                             Constantes.Views.ColGridLog.Nombre,
                                             Constantes.Views.ColGridLog.Tabla,
                                             Constantes.Views.ColGridLog.DimCode);
                                        Recordset rsCards = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rsCards.DoQuery(strSQL);


                                        try
                                        {
                                            oForm = B1.Application.Forms.Item(pVal.FormUID);

                                            SAPbouiCOM.Item Oitem2 = (SAPbouiCOM.Item)oForm.Items.Add("SSI_DPTOT", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                                            Oitem2.FromPane = 0;
                                            Oitem2.Left = oForm.Items.Item("540002008").Left;
                                            Oitem2.Top = oForm.Items.Item("540002008").Top + oForm.Items.Item("540002008").Height + 5;
                                            Oitem2.Width = 80;
                                            Oitem2.Height = 16;
                                            Oitem2.LinkTo = "540002008";
                                            SAPbouiCOM.StaticText DptpDesc = (SAPbouiCOM.StaticText)Oitem2.Specific;
                                            DptpDesc.Caption = "Departamento ";


                                            SAPbouiCOM.Item Oitem = (SAPbouiCOM.Item)oForm.Items.Add("SSI_DPTOS2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                                            Oitem.FromPane = 0;
                                            Oitem.Left = oForm.Items.Item("540002010").Left;
                                            Oitem.Top = oForm.Items.Item("540002010").Top + oForm.Items.Item("540002010").Height + 5;
                                            Oitem.Width = 198;
                                            Oitem.Height = 16;
                                            Oitem.DisplayDesc = true;
                                            Oitem.Description = "Departamento";
                                            Oitem.LinkTo = "SSI_DPTOT";
                                            Oitem.AffectsFormMode = true;
                                            SAPbouiCOM.ComboBox oComboBox = Oitem.Specific;
                                            oComboBox.DataBind.SetBound(true, "OPRC", "U_SSI_DPTOS");

                                            SSIFramework.Utilidades.GenericFunctions.fillComboBySQL(ref oComboBox, strSQL, "PrcCode", "PrcName", true);
                                            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                                        }

                                        catch (Exception ex) { throw ex; }
                                        break;

                                    case CdeCostesUf:
                                        try
                                        {
                                            //oForm = B1.Application.Forms.Item(pVal.FormUID);
                                            //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("U_SSI_DPTOS").Specific;
                                            //SAPbouiCOM.Item Oitem2;
                                            //Oitem2 = oForm.Items.Item("U_SSI_DPTOS");
                                            //Oitem2.Enabled = false;
                                            //oEdit.Item.FromPane=1;
                                            //Oitem2.Visible = false;
                                            ((SAPbouiCOM.EditText)B1.Application.Forms.GetForm("-810", pVal.FormTypeCount).Items.Item("U_SSI_DPTOS").Specific).Item.Visible = false;

                                        }
                                        catch (Exception ex) { throw ex; }
                                        break;

                                }
                            }
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
    }
}
