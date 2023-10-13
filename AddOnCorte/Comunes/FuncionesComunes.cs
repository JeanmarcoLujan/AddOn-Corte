using AddOnCorte.Clases;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Comunes
{
    class FuncionesComunes
    {
        #region Metodos

        public static void LiberarObjetoGenerico(Object objeto)
        {
            try
            {
                if (objeto != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(AddOnCorte.Properties.Resources.NombreAddon + " Error Liberando Objeto: " + ex.Message);
            }
        }

        /// <summary>
        /// Método para mostrar los errores a través de la barra de mansajes de SAP B1 en la sociedad conectada.
        /// </summary>
        /// <param name="messageError">Descripción del mensaje de error.</param>
        /// <param name="oMethodBase">Objeto Reflection con el detalle de la clase donde se generó el error.</param>
        public static void DisplayErrorMessages(string messageError, System.Reflection.MethodBase oMethodBase)
        {
            string className = string.Empty;
            string methodName = string.Empty;

            try
            {
                className = oMethodBase.DeclaringType.Name.ToString();
                methodName = oMethodBase.Name.ToString();

                Globales.oApp.StatusBar.SetText(Properties.Resources.NombreAddon +
                    " Error: " + className + ".cs > " + methodName + "(): " + messageError,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            catch (Exception ex)
            {
                Globales.oApp.StatusBar.SetText(Properties.Resources.NombreAddon + " Error: FuncionesComunes.cs > DisplayErrorMessages(): "
                    + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                LiberarObjetoGenerico(oRecordSet);
                oRecordSet = null;
            }
        }


        public static int ValidateNumberInt(string inputNumber)
        {
            int number;
            if (int.TryParse(inputNumber, out number))
                return number;
            else
                Globales.oApp.MessageBox("El valor especificado debe ser un número entero");
            return 0;
        }

        public static void UpdateUDO(string udoId, string docEntryOf)
        {

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = Globales.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("MGS_CL_COCABE");
                oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
                oGeneralParams.SetProperty("DocEntry", udoId);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("U_MGS_CL_OFEV", docEntryOf);
                oGeneralService.Update(oGeneralData);
                //return true;
            }
            catch (Exception ex)
            {
                //Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                //return false;
            }
        }

        public static void UpdateUDOAgendarSolicitud(string udoId, string fecha, string equipo, string serie)
        {

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = Globales.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("MGS_CL_COCABE");
                oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
                oGeneralParams.SetProperty("DocEntry", udoId);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("U_MGS_CL_AFECOR", fecha);
                oGeneralData.SetProperty("U_MGS_CL_AEQUIP", equipo);
                oGeneralData.SetProperty("U_MGS_CL_ASERIA", serie);
                oGeneralData.SetProperty("U_MGS_CL_ESTD", "G");
                oGeneralService.Update(oGeneralData);
                //return true;
            }
            catch (Exception ex)
            {
                //Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                //return false;
            }
        }




        #endregion
    }
}
