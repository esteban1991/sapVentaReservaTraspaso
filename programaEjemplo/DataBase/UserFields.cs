using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

using SSIFramework;
using SSIFramework.Utilidades;

namespace ventaRT.DataBase
{
    class UserFields
    {
        public static object Framework { get; private set; }

        public static void CrearEstructura()
        {
            try {
            SSIConnector.GetSSIConnector().Application.MetadataAutoRefresh = false;
                //CrearEstructuraImportacion();

            //crear tabla prueba 1
            //CrearTablasSAP();

            //CrearCamposUsuarioTablasSAP();

            CreateUserTables();

            SSIConnector.GetSSIConnector().Application.MetadataAutoRefresh = true;
            }
            catch (Exception ex) { throw ex; }
        }

        // Gisela : creando nueva tabla MD prueba1 para proveedores
        public static void CreateUserTables()
        {


            try
            {

                //if (!GenericFunctions.ExistUserField("DET_RV", "codClie"))
                //{
                //GC.Collect();
                //GenericFunctions.DelUserField("DET_RV", "codClie");
                //GC.Collect();

                GC.Collect();
                GenericFunctions.AddUserField("DET_RV", "codArti", "Articulo", HelpBaseType.Tipo.Regular,50, "");
                GC.Collect();


                //GC.Collect();
                //GenericFunctions.AddUserField("DET_RV", "codCliee", "Cliente", HelpBaseType.Tipo.Regular, 20, "");
                //GC.Collect();

                //}


                if (!GenericFunctions.ExistUserTable("CAB_RV"))
                {

                    GC.Collect();
                    GenericFunctions.AddUserTable("CAB_RV", "Cab Res Venta", BoUTBTableType.bott_NoObject); // MasterDataTyp 
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "numOC", "no orden compra", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "idVend", "Id Vendedor", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "fechaC", "Fecha de Creacion", HelpBaseType.Tipo.Date, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "fechaV", "Fecha de Vencimiento", HelpBaseType.Tipo.Date, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "estado", "Estado de la Solicitud", HelpBaseType.Tipo.Regular, 5, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "idTV", "Id Transferencia Virtual", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "idTR", "Id Transferencia Real o DocEntry", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "idAut", "Id Autorizador", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RV", "comment", "Comentarios", HelpBaseType.Tipo.Text, 150, "");
                    GC.Collect();
                }


 

                if (!GenericFunctions.ExistUserTable("DET_RV"))
                {
                    GC.Collect();
                    GenericFunctions.AddUserTable("DET_RV", "DETALLE VENTA RESERVA", BoUTBTableType.bott_NoObject); 
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RV", "numOC", "no orden compra", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();
                    
                    
                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RV", "codArti", "Cod Articulo", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RV", "codClie", "Cliente", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RV", "cant", "Cantidad", HelpBaseType.Tipo.Quantity, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RV", "estado", "Estado de la Linea" , HelpBaseType.Tipo.Regular, 5, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RV", "idTV", "Id Transferencia Virtual", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();


                }

            }
            catch (Exception EX) { throw EX; }



        }


        public static void CrearCamposUsuarioTablasSAP()
        {
          
            //preguntar si va este campo ya que ya existiran los dptos creado como centros de costes
            try { 
            if (!GenericFunctions.ExistUserField("OPRC", "SSI_DPTOS"))
            {
                GC.Collect();
                GenericFunctions.AddUserField("OPRC", "SSI_DPTOS", null, HelpBaseType.Tipo.Regular, 20, "");
                GC.Collect();
            }
            }
            catch(Exception EX) { throw EX; }

            //cuenta que tendra el addon para el haber, tendre que tomarlo desde aca en  principio seria la 4659999
            try
            {
                if (!GenericFunctions.ExistUserField("@OADM", "SSI_CtaPu"))
                {
                    GC.Collect();
                    GenericFunctions.AddUserField("@OADM", "SSI_CtaPu", "Cuenta Puente para addon", HelpBaseType.Tipo.Regular, 13, "");
                    GC.Collect();
                }


                if (!GenericFunctions.ExistUserField("@OADM", "SSICtaDesc"))
                {
                    GC.Collect();
                    GenericFunctions.AddUserField("@OADM", "SSICtaDesc", "Cuenta Descuadres Varios", HelpBaseType.Tipo.Regular, 13, "");
                    GC.Collect();
                }
            }
            catch (Exception EX) { throw EX; }

        }

    }
}
