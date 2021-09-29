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

            CreateUserTables();

            SSIConnector.GetSSIConnector().Application.MetadataAutoRefresh = true;
            }
            catch (Exception ex) { throw ex; }
        }

        public static void CreateUserTables()
        {


            try
            {

                //if (!GenericFunctions.ExistUserTable("SER_RSTV"))
                //{

                //    GC.Collect();
                //    GenericFunctions.AddUserTable("SER_RSTV", "Serie Doc Res Stock Traspaso", BoUTBTableType.bott_NoObject);
                //    GC.Collect();

                //    GC.Collect();
                //    GenericFunctions.AddUserField("SER_RSTV", "cabprox", "Doc Prox", HelpBaseType.Tipo.Regular, 25, "");
                //    GC.Collect();

                //}

                if (!GenericFunctions.ExistUserTable("AUT_RSTV"))
                {

                    GC.Collect();
                    GenericFunctions.AddUserTable("AUT_RSTV", "Aprobadores Res Stock Traspaso", BoUTBTableType.bott_NoObject);
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("AUT_RSTV", "idAut", "Id Autorizador", HelpBaseType.Tipo.Regular, 25, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("AUT_RSTV", "aut", "Autorizador", HelpBaseType.Tipo.Regular, 155, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("AUT_RSTV", "activo", "Activo", HelpBaseType.Tipo.Regular, 3, "");
                    GC.Collect();

                }
 
                if (!GenericFunctions.ExistUserTable("CAB_RSTV"))
                {

                    GC.Collect();
                    GenericFunctions.AddUserTable("CAB_RSTV", "Cabecera Res Stock Traspaso", BoUTBTableType.bott_NoObject);  
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "numDoc", "no OC", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "idVend", "Id Vendedor", HelpBaseType.Tipo.Regular, 25, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "fechaC", "Fecha de Creacion", HelpBaseType.Tipo.Date, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "fechaV", "Fecha de Vencimiento", HelpBaseType.Tipo.Date, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "estado", "Estado de la Solicitud", HelpBaseType.Tipo.Regular, 15, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "idTV", "Id Transf Virtual", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "idTR", "Id Transf Real o DocEntry", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "idAut", "Id Autorizador", HelpBaseType.Tipo.Regular, 25, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "comment", "Comentarios", HelpBaseType.Tipo.Text, 150, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "logs", "Logs del Sistema", HelpBaseType.Tipo.Text, 300, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "diasv", "Dias Vigentes", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "vend", "Vendedor", HelpBaseType.Tipo.Regular, 155, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("CAB_RSTV", "aut", "Autorizador", HelpBaseType.Tipo.Regular, 155, "");
                    GC.Collect();

                }




                if (!GenericFunctions.ExistUserTable("DET_RSTV"))
                {
                    GC.Collect();
                    GenericFunctions.AddUserTable("DET_RSTV", "Detalle Res Stock Traspaso", BoUTBTableType.bott_NoObject);
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "numOC", "no orden compra", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();


                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "codArt", "Cod Articulo", HelpBaseType.Tipo.Regular, 50, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "codCli", "Cliente", HelpBaseType.Tipo.Regular, 20, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "cant", "Cantidad", HelpBaseType.Tipo.Quantity, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "estado", "Estado de la Linea", HelpBaseType.Tipo.Regular, 5, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "idTV", "Id Transf Virtual", HelpBaseType.Tipo.Regular, 10, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "articulo", "Articulo", HelpBaseType.Tipo.Regular, 100, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "cliente", "Cliente", HelpBaseType.Tipo.Regular, 100, "");
                    GC.Collect();

                    GC.Collect();
                    GenericFunctions.AddUserField("DET_RSTV", "onHand", "Stock", HelpBaseType.Tipo.Quantity, 10, "");
                    GC.Collect();




                }





            }
            catch (Exception EX) { throw EX; }



        }


    }
}
