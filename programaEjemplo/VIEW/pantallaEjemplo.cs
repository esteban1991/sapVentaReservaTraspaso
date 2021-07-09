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

namespace ventaRT.VIEW
{
    class pantallaEjemplo : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();

        public pantallaEjemplo()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.pantallaEjemplo.srf"], "pantEjemp" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            //los dos  eventos mas usados
            this.B1.Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(ThisSapApiForm_MenuEvent);
            this.B1.Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_ItemEvent);
        }

        private void ThisSapApiForm_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            //sirve para que continue con los eventos en sap, si es en false lo contrario, los detiene despues de pasar por aqui
            BubbleEvent = true;
        }

        private void ThisSapApiForm_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

       
    }
}
