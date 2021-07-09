using ventaRT.Constantes.Views;
using SAPbobsCOM;
using SAPbouiCOM;
using SSIFramework;
using SSIFramework.DI.Attributes;
using SSIFramework.Utilidades;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ventaRT.VIEW
{
    class PantallaPrueba1 : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();
        private string nSN = ""; //para obtener nombre del Socio de Negocio
        private string mSN = ""; //para obtener moneda del Socio de Negocio

        public PantallaPrueba1()
            : base(GenericFunctions.ResourcesForms["ventaRT.Forms.prueba1.srf"], "P1" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) 
        {
            ThisSapApiForm.OnAfterLostFocus += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterLostFocus);
            ThisSapApiForm.OnAfterItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterItemPressed);           
        }

        private void ThisSapApiForm_OnAfterLostFocus(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
             BubbleEvent = true;
             try
             {
                 SAPbouiCOM.EditText txtid = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.prueba1.txt_id).Specific;
             
                 switch (pVal.EventType)
                 {
                     case BoEventTypes.et_LOST_FOCUS:
                         switch (pVal.ItemUID)
                         {
                             case Constantes.View.prueba1.txt_id:
                                 if (txtid.Value == "")
                                 {
                                     B1.Application.SetStatusBarMessage("Error se necesita un id de proveedor" , SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                 }
                                 else
                                 {
                                      BubbleEvent = true;
                                      Importar_Datos_SN();
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


        private void Importar_Datos_SN()
        {

            try
            {
                SAPbouiCOM.EditText txtid= (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.prueba1.txt_id).Specific;
                //SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                Recordset oRecordset = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                String consultaSN = string.Format("SELECT {0},{1} FROM {2} Where contains({3},'%{4}%')",
                                                    ventaRT.Constantes.View.SNConsulta.fName,
                                                    ventaRT.Constantes.View.SNConsulta.fMoneda,
                                                    ventaRT.Constantes.View.SNConsulta.OCRD,
                                                    ventaRT.Constantes.View.SNConsulta.fCode,
                                                    txtid.Value
                );
                oRecordset.DoQuery(consultaSN);
                nSN = oRecordset.Fields.Item("CardName").Value.ToString();  //obteniendo nombre del Socio Neg
                mSN = oRecordset.Fields.Item("Currency").Value.ToString();  //obteniendo moneda del Socio Neg

                if (nSN != null)
                {
                    // Actualizando campos en el formulario a partir de la consulta realizada
                    SAPbouiCOM.EditText txtnombreProv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.prueba1.txt_nom).Specific;
                    SAPbouiCOM.EditText txtmonedaProv = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(Constantes.View.prueba1.txt_mon).Specific;
                    txtnombreProv.Value = nSN;
                    txtmonedaProv.Value = mSN;
                    B1.Application.SetStatusBarMessage("Exito", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                }
            }
            catch (Exception ex) { throw ex; }
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
                            case Constantes.View.prueba1.btn_reset:
                                {
  
                                    SAPbouiCOM.EditText txtNom = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_nom).Specific;
                                    txtNom.Value = " ";
                                    SAPbouiCOM.EditText txtMon = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_mon).Specific;
                                    txtMon.Value = " ";
                                    SAPbouiCOM.EditText txtId = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_id).Specific;
                                    txtId.Value = " ";
                                    txtId.Active = true;                               
                                    BubbleEvent = true;

                                }
                                break;
                            case Constantes.View.prueba1.btn_save:
                                {
                                    SAPbouiCOM.EditText txtId = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_id).Specific;
                                    SAPbouiCOM.EditText txtNom = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_nom).Specific;
                                    SAPbouiCOM.EditText txtMon = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_mon).Specific;
                                    if (txtId.Value.ToString().Length > 0)
                                    {
                                        SAPbobsCOM.UserTable oUserTableCa;
                                        oUserTableCa = B1.Company.UserTables.Item("Prueba1");
                                        int IDnEXT = obtenerUltimoID("Prueba1");
                                        IDnEXT = IDnEXT + 1;
                                        oUserTableCa.Code = IDnEXT.ToString();
                                        oUserTableCa.Name = IDnEXT.ToString();


                                        // validando si exisita ese codigo

                                        String strSQL = String.Format("SELECT  COUNT(*)  FROM  {0}  Where contains({1},'%{2}%')",
                                                    Constantes.View.P1Consulta.Prueba1, Constantes.View.P1Consulta.fidprov, txtId.Value.ToString());
                                        Recordset rsReg = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rsReg.DoQuery(strSQL);
                                        string Code = rsReg.Fields.Item("COUNT(*)").Value.ToString();
                                        if (Convert.ToInt32(Code) > 0)
                                        {
                                            B1.Application.SetStatusBarMessage("Error Adicionando en Prueba 1, ya existe ese Proveedor", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                        }
                                        else
                                        {
                                            B1.Application.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);


                                            oUserTableCa.UserFields.Fields.Item("U_idprov").Value = txtId.Value.ToString();
                                            int lt = txtNom.Value.ToString().Length;
                                            oUserTableCa.UserFields.Fields.Item("U_nomprov").Value = txtNom.Value.ToString().Substring(0, lt == 1 ? 1 : lt < 10 ? lt - 1 : 9);
                                            lt = txtMon.Value.ToString().Length;
                                            oUserTableCa.UserFields.Fields.Item("U_moneda").Value = txtMon.Value.ToString().Substring(0, lt == 1 ? 1 : lt < 10 ? lt - 1 : 9);



                                            int i = oUserTableCa.Add();

                                            if (i != 0)
                                            {
                                                B1.Application.SetStatusBarMessage("Error" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);

                                            }
                                            else
                                            {
                                                B1.Application.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            }
                                        } 
                      
                                    }
                                     else {
                                            B1.Application.SetStatusBarMessage("Error: Id no valido" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true );

                                     }  
                    
                     

                                }
               
                                                break;
                

                            case Constantes.View.prueba1.btn_sap:
                                {
                                    SAPbouiCOM.EditText txtId = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_id).Specific;
                                    SAPbouiCOM.EditText txtNom = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_nom).Specific;
                                    SAPbouiCOM.EditText txtMon = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.View.prueba1.txt_mon).Specific;

                                  

                                    SAPbobsCOM.BusinessPartners oItem;
                                    oItem = B1.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                                    if (oItem.GetByKey(txtId.Value.ToString()) == true)
                                    {
                                        oItem.CardName = txtNom.Value.ToString ();
                                        oItem.Currency = txtMon.Value.ToString ();
                                    }

                                    int j = oItem.Update();
                                    if (j != 0)
                                    {
                                        B1.Application.SetStatusBarMessage("Error actualizando SAP Socio Negocio" + B1.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);

                                    }
                                    else
                                    {
                                        B1.Application.SetStatusBarMessage("Exito actualizando socio Negocio", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    }


                                    
                                    txtNom.Value = " ";
                                    txtMon.Value = " ";
                                    txtId.Value = " ";

                           

                                }

                                break;

                            case Constantes.View.prueba1.btn_exit:
                                {
                                    SAPbouiCOM.Form oForm = B1.Application.Forms.ActiveForm;
                                    oForm.Close();
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


        private int obtenerUltimoID(string tipo)
        {
            int CodeNum = 0;


                String strSQL = String.Format("SELECT  COUNT(*)  FROM  {0}" ,
                  Constantes.View.P1Consulta.Prueba1);
                    
    

                Recordset rsReg = (Recordset)B1.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsReg.DoQuery(strSQL);

                string Code = rsReg.Fields.Item("COUNT(*)").Value.ToString();

                //probar cuando la tabla este vacia, osea el primero registro y no haya otro anterior
                if (Code != "")
                {
                    CodeNum = Convert.ToInt32(Code);

                }
                return CodeNum;


        }

        
        }

 
           
           
        
}
