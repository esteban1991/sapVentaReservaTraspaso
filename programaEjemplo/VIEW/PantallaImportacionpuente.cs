using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SAPbobsCOM;
using SSIFramework;
using SSIFramework.UI.UIApi;
using SSIFramework.Utilidades;
using System.Threading;


namespace ventaRT.VIEW
{
    class PantallaImportacionpuente : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();

        public PantallaImportacionpuente() : base(GenericFunctions.ResourcesForms["ContabilizacionDeNominas.Forms.CargarPuente.srf"], "SSI_Imp_ExcelPuente" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString())
        {
            

            ThisSapApiForm.OnAfterItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterItemPressed);

           // CrearGridpuente();
        }

        public void CrearGridpuente()
        {
            throw new NotImplementedException();
        }

        private void ThisSapApiForm_OnAfterItemPressed(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            switch (pVal.EventType)
            {
                case BoEventTypes.et_ITEM_PRESSED:
                    switch (pVal.ItemUID)
                    {
                        case Constantes.Views.PantallaImpoPuente.ButtonImportacion:


                            List<String> ListDocEntry = null;
                            bool ErroresEncontrados = false;
                            BubbleEvent = true;
                            String stFichero = ((SAPbouiCOM.EditText)this.ThisSapApiForm.Form.Items.Item(Constantes.Views.PantallaImpoPuente.EditTxtFichero).Specific).Value;
                            Importacion.Imp_Nominas oImpNomin = new Importacion.Imp_Nominas();
                            ErroresEncontrados = oImpNomin.GestionarImportacionPuente(stFichero, this, ref ListDocEntry);
                            break;
                        case Constantes.Views.PantallaImpoPuente.ButtonBuscarFichero:
                            SAPbouiCOM.Form oForm = B1.Application.Forms.Item(FormUID);
                            String sFichero;
                            SSIFramework.Utilidades.OpenFileDialogThread Dialog = new OpenFileDialogThread("Selección de archivo Excel");

                            Thread oThread = new Thread(new ThreadStart(Dialog.ShowDialog));
                            oThread.SetApartmentState(ApartmentState.STA);
                            oThread.Start();
                            oThread.Join();
                            if (Dialog.DialogResult == System.Windows.Forms.DialogResult.OK)
                            {
                                sFichero = Dialog.SelectedFile;
                                ((SAPbouiCOM.EditText)oForm.Items.Item(Constantes.Views.PantallaImpoNominas.EditTxtFichero).Specific).String = sFichero;
                            }
                            break;


                    }
                    break;
            }
        }
    }
}
