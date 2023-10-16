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

        public static void CbxSolCorte(ref SAPbouiCOM.Form oForm, SAPbouiCOM.Application SBOApp, SAPbouiCOM.ComboBox oComboBox)
        {

            try
            {
                InsertValidValues(ref oComboBox
                                                , Comunes.Consultas.GetCortesCmb()
                                                , "Value", "Desc", bSetValueDefault: true);
            }
            catch
            {

            }

        }

        internal static void InsertValidValues(ref SAPbouiCOM.ComboBox oCombo, String Query, String fieldValue, String fieldDesc, string fielValueIfNull = "-1", string fielDescIfNull = "--", bool bSetValueDefault = false)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            SAPbouiCOM.ValidValues oValidValues = null;
            string sDescription = string.Empty;
            try
            {
                oValidValues = oCombo.ValidValues;
                oRecordSet = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(Query);
                while (oCombo.ValidValues.Count > 0)
                {
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        sDescription = oRecordSet.Fields.Item(fieldDesc).Value.ToString().Trim();
                        if (sDescription.Length > 100) sDescription = sDescription.Substring(0, 99);

                        oValidValues.Add(oRecordSet.Fields.Item(fieldValue).Value.ToString().Trim(), sDescription);
                        oRecordSet.MoveNext();
                    }
                    try
                    {
                        if (bSetValueDefault)
                        {
                            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                    }
                    catch (Exception) { }
                }
                else
                {
                    oValidValues.Add(fielValueIfNull, fielDescIfNull);
                }
            }
            catch (Exception ex)
            {

                Globales.oApp.StatusBar.SetText(Properties.Resources.NombreAddon + " Error: FuncionesComunes.cs > DisplayErrorMessages(): "
                    + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                //LiberarObjetoGenerico(oRecordSet);
                oRecordSet = null;
            }
        }



    }
}
