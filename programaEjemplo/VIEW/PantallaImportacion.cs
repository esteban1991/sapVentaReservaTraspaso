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
    class PantallaImportacion : SSIFramework.UI.UIApi.UserForm
    {
        private readonly SSIConnector B1 = SSIConnector.GetSSIConnector();


        public PantallaImportacion() : base(GenericFunctions.ResourcesForms["ContabilizacionDeNominas.Forms.CargarNominas.srf"], "SSI_Imp_ExcelNom" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString()) {



            ThisSapApiForm.OnAfterItemPressed += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterItemPressed);
            ThisSapApiForm.OnAfterValidate += new _IApplicationEvents_ItemEventEventHandler(ThisSapApiForm_OnAfterValidate);
            //desabilito el boton importar hasta que tenga una ruta y fecha al menos
            SAPbouiCOM.Button btn = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(Constantes.Views.PantallaImpoNominas.ButtonImportacion).Specific;
            btn.Item.Enabled = false;
            CrearGrid();
        }

        private void ThisSapApiForm_OnAfterValidate(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {

            SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)B1.Application.Forms.ActiveForm.Items.Item(ventaRT.Constantes.Views.PantallaImpoNominas.TxtFechaContab).Specific;
            object p2 = new Importacion.Imp_Nominas();
      
            try { 
              //cargar el combo cuando este el txt de fecha de conta tenga un dato
             if (txt.Value.ToString() != "") { 
            DateTime dt2 = DateTime.Parse(txt.String);
            string s3 = dt2.ToString(ventaRT.Constantes.Views.formatos.formatoFechaAsiento);
            DateTime dtnew2 = DateTime.Parse(s3);
           //casting a la clase imp_nominas, para poder ejecutar el metodo y pasarle el parametro
            ((Importacion.Imp_Nominas)p2).cargarCombo(dtnew2);

            }

            }
            catch (Exception ex){ throw ex; }

            // habilitar boton si hay fecha y un archivo
            try {
                string stFichero = ((SAPbouiCOM.EditText)this.ThisSapApiForm.Form.Items.Item(Constantes.Views.PantallaImpoNominas.EditTxtFichero).Specific).Value;
                string fechaVacia= ((SAPbouiCOM.EditText)this.ThisSapApiForm.Form.Items.Item(Constantes.Views.PantallaImpoNominas.TxtFechaContab).Specific).Value;

                if (stFichero != "" && fechaVacia!="")
                {
                    SAPbouiCOM.Button btn = (SAPbouiCOM.Button)B1.Application.Forms.ActiveForm.Items.Item(Constantes.Views.PantallaImpoNominas.ButtonImportacion).Specific;
                    btn.Item.Enabled = true;
                }

               
            }
            catch (Exception ex) {throw ex; }
            BubbleEvent = true;
        } 


        #region ThisSapApiForm_OnAfterItemPressed
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
                            case Constantes.Views.PantallaImpoNominas.ButtonImportacion:


                                List<String> ListDocEntry = null;
                                bool ErroresEncontrados = false;
                                BubbleEvent = true;
                                String stFichero = ((SAPbouiCOM.EditText)this.ThisSapApiForm.Form.Items.Item(Constantes.Views.PantallaImpoNominas.EditTxtFichero).Specific).Value;
                           
                                Importacion.Imp_Nominas oImpNomin = new Importacion.Imp_Nominas();
                                
                                ErroresEncontrados = oImpNomin.GestionarImportacionNominas(stFichero, this, ref ListDocEntry);
                                
          
                                break;
                            case Constantes.Views.PantallaImpoNominas.ButtonBuscarFichero:
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
            catch (Exception ex)
            { throw ex; }
            finally { GC.Collect(); }
        }
        #endregion

        public void CrearGrid()
        {
            try
            {



              //String sSQL = String.Format("SELECT CAST('' as nvarchar({4})) as '{5}', CAST('' as nvarchar({0}) ) as '{1}',CAST('' as nvarchar({2})) as '{3}'",
               // Constantes.Views.ColGridLog.ColMensajeSize,
               //Constantes.Views.ColGridLog.ColMensaje,
               //Constantes.Views.ColGridLog.ColTipoErrorSize,
               //Constantes.Views.ColGridLog.ColTipoError,
               //Constantes.Views.ColGridLog.ColLineaSize,
               //Constantes.Views.ColGridLog.ColLinea);
               // this.ThisSapApiForm.Form.DataSources.DataTables.Item(Constantes.Views.PantallaImpoNominas.DataTableLog).ExecuteQuery(sSQL);
               // SAPbouiCOM.Grid oGrid = this.ThisSapApiForm.Form.Items.Item(Constantes.Views.PantallaImpoNominas.GridLog).Specific;

               // //((EditTextColumn)oGrid.Columns.Item(Constantes.Views.ColGridLog.ColDatoCreado)).LinkedObjectType = "13";
               // //oGrid.Columns.Item(Constantes.Views.ColGridLog.ColObjectType).Visible = false;

               // oGrid.CollapseLevel = 1;
               // oGrid.AutoResizeColumns();



                //String sSQL = String.Format("SELECT {0},{1} FROM {3}",
                //Constantes.Views.ColGridLog.PrCode,
                //Constantes.Views.ColGridLog.Nombre,
                //  Constantes.Views.ColGridLog.Tabla);


                //this.ThisSapApiForm.Form.DataSources.DataTables.Item(Constantes.Views.PantallaImpoNominas.DataTableLog).ExecuteQuery(sSQL);
                //SAPbouiCOM.Grid oGrid = this.ThisSapApiForm.Form.Items.Item(Constantes.Views.PantallaImpoNominas.GridLog).Specific;

                ////oGrid.CollapseLevel = 1;
                //oGrid.AutoResizeColumns();
            }
            catch (Exception ex){ throw ex; }
        }


        public void AnotarEnGridError(String sMensajeError, String sTipoError, string asiento)
        {
            try
            {
                //SELECT CAST('' as nvarchar(250) ) as 'Mensaje',CAST('' as nvarchar(10)) as 'ObjectType',CAST('' as nvarchar(30)) as 'Dato Creado'
                ThisSapApiForm.Item(Constantes.Views.PantallaImpoNominas.GridLog).Enabled = false;
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)ThisSapApiForm.Item(Constantes.Views.PantallaImpoNominas.GridLog).Specific;
                //if (oGrid.DataTable.GetValue("Error", oGrid.DataTable.Rows.Count - 1) != "")
                //{
                    oGrid.DataTable.Rows.Add();
                //}

                
                //oGrid.DataTable.SetValue(Constantes.Views.ColGridLog.ColLinea, oGrid.DataTable.Rows.Count - 1, "Línea: " + Linea);
                oGrid.DataTable.SetValue("Error", oGrid.DataTable.Rows.Count - 1, sTipoError);
                oGrid.DataTable.SetValue(Constantes.Views.ColGridLog.ColMensaje, oGrid.DataTable.Rows.Count - 1, sMensajeError);
                oGrid.DataTable.SetValue("Fecha", oGrid.DataTable.Rows.Count - 1, DateTime.Now);

                switch (asiento) {
                    case Constantes.Views.ASIENTOS.Barcelona:
                        asiento = "Barcelona";
                        break;
                    case Constantes.Views.ASIENTOS.Becarios:
                        asiento = "Becarios";
                        break;
                    case Constantes.Views.ASIENTOS.Madrid:
                        asiento = "Madrid";
                        break;
                    case Constantes.Views.ASIENTOS.Andalucia:
                        asiento = "Andalucía";
                        break;
                    case Constantes.Views.ASIENTOS.Canarias:
                        asiento = "Canarias";
                        break;

                }
                oGrid.DataTable.SetValue("Asiento", oGrid.DataTable.Rows.Count - 1, asiento);

                //oGrid.DataTable.SetValue(Constantes.Views.ColGridLog.ColObjectType, oGrid.DataTable.Rows.Count - 1, sObjecteType);
                //oGrid.DataTable.SetValue(Constantes.Views.ColGridLog.ColDatoCreado, oGrid.DataTable.Rows.Count - 1, sDatoCreado);

                oGrid.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        }
    }

