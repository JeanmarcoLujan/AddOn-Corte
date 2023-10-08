using AddOnCorte.Clases;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Comunes
{
    class Funciones
    {

        internal static void AutoResizeColumnsMatrix(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            try
            {

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;
                //oMatrix.AddRow();
                oMatrix.FlushToDataSource();

                //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                //else bAddNewRow = false;
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.LoadFromDataSource();


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
                //oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_CORRID");
                //oDBDataSource.SetValue("U_MGS_CL_TITU", 0, "Hola 0");
                //oMatrix.AddRow();
                //oMatrix.FlushToDataSource();

                //oMatrix = oForm.Items.Item("ID_de_Tu_Matriz").Specific
                


                for (int i = 0; i < 21; i++)
                {
                    
                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_CORRID");
                    var sdsd = oDBDataSource.Size - 1;
                    oDBDataSource.Offset = oDBDataSource.Size - 1;
                    string descripcio = "";
                    if (i == 0)
                        descripcio = "Largo (ft)";
                    else if (i == 1)
                        descripcio = "Resfile";
                    else if (i == 17)
                        descripcio = "Subtotal";
                    else if (i == 18)
                        descripcio = "Total";
                    else if (i == 19)
                        descripcio = "ChkS Ancho dif";
                    else if (i == 20)
                        descripcio = "Sum M2 x corr";
                    else
                        descripcio = "Slit#" + (i - 1);

                    oDBDataSource.SetValue("U_MGS_CL_TITU", oDBDataSource.Size - 1, descripcio);
                    for (int j = 1; j <= 10; j++)
                    {
                        string columnId = "U_MGS_CL_C" + j;
                        oDBDataSource.SetValue(columnId, oDBDataSource.Size - 1, "0");
                    }
                    
                    oMatrix.AddRow();
                    oMatrix.FlushToDataSource();


                }
              

                for (int colIndex = 1; colIndex <= oMatrix.Columns.Count-1; colIndex++)
                {
                    oMatrix.CommonSetting.SetCellEditable(18, colIndex, false);
                    oMatrix.CommonSetting.SetCellEditable(19, colIndex, false);
                    oMatrix.CommonSetting.SetCellEditable(20, colIndex, false);
                    oMatrix.CommonSetting.SetCellEditable(21, colIndex, false);
                }

                //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                //else bAddNewRow = false;
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.LoadFromDataSource();

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;
                //oMatrix.AddRow();

                //oMatrix.FlushToDataSource();

                //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                //else bAddNewRow = false;
                //oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                //oMatrix.LoadFromDataSource();

                

               


            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        internal static void AutoResizeColumnsMatrixInicial(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            try
            {

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;
           
                oMatrix.FlushToDataSource();

                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.LoadFromDataSource();


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
                
                oMatrix.FlushToDataSource();
                //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                //else bAddNewRow = false;
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.LoadFromDataSource();

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;
                
                oMatrix.FlushToDataSource();


                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.LoadFromDataSource();


            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }


        

        
    }
}
