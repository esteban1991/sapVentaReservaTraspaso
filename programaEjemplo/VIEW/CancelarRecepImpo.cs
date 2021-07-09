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
using ventaRT.Constantes.Views;

namespace ventaRT.VIEW
{
    class CancelarRecepImpo : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        public CancelarRecepImpo()
            : base(GenericFunctions.ResourcesForms["ContabilizacionDeNominas.Forms.CancelRecepImpo.srf"], "HJ_RecepICan" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            ThisSapApiForm.OnAfterItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterItemPressed);          
        }

        private void ThisSapApiForm_OnAfterItemPressed(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
             BubbleEvent = true;
             try
             {
                 SAPbouiCOM.EditText num = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.camcelRecepImpo.num).Specific;
                 //SAPbouiCOM.EditText comment = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.camcelRecepImpo.txt_numD).Specific;
                 SAPbouiCOM.EditText txtComen = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.camcelRecepImpo.txt_com).Specific;
                 //SAPbouiCOM.EditText txtnid = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.camcelRecepImpo.num).Specific;
               
                 switch (pVal.EventType)
                 {
                     case BoEventTypes.et_ITEM_PRESSED:
                         switch (pVal.ItemUID)
                         {
                             case Constantes.View.camcelRecepImpo.btn_can:
                                 if (txtComen.Value == "")
                                 {
                                     B1.Application.SetStatusBarMessage("Error se necesita un comentario" , SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                 }
                                 else if (num.Value == "")
                                 {
                                     B1.Application.SetStatusBarMessage("Error se debe poner un numero mayor a cero", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                 }
                                 else
                                 {
                                     int ithReturnValue;
                                     ithReturnValue = B1.Application.MessageBox("Seguro que quieres cancelar el documento ?", 1, "Continuar", "Cancelar", "");

                                     if (ithReturnValue == 1)
                                     {
                                         BubbleEvent = true;
                                         CancelarDocumento();
                                     }
                                     else
                                     {
                                         BubbleEvent = false;
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
             }
        }

        private void CancelarDocumento()
        {

            SAPbouiCOM.EditText txtNumeroDocu = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.camcelRecepImpo.num).Specific;
            SAPbobsCOM.UserTable oUserTableDE;
            oUserTableDE = B1.Company.UserTables.Item("CAB_REC_IMP");
            oUserTableDE.GetByKey(txtNumeroDocu.Value);
            oUserTableDE.UserFields.Fields.Item("U_Estado").Value = "X";
            oUserTableDE.UserFields.Fields.Item("U_Nom_Creador").Value = B1.Company.UserName;
            int i = oUserTableDE.Update();

            if (i != 0)
            {
                B1.Application.SetStatusBarMessage("Error" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }else
            {
                B1.Application.SetStatusBarMessage("Exito" , SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
        }
    }
}
