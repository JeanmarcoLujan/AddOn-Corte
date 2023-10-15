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

        public static void UpdateUDO(string udoId, string docEntryOf, string fieldUpdate)
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
                oGeneralData.SetProperty(fieldUpdate, docEntryOf);
                if(fieldUpdate.ToString().Equals("U_MGS_CL_SOLTRA"))
                    oGeneralData.SetProperty("U_MGS_CL_ESTD", "S");
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


        public static Solicitud GetUDOSolicitudAgendada(string udoId)
        {

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;

            try
            {
                oCompanyService = Globales.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("MGS_CL_COCABE");
                oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
                oGeneralParams.SetProperty("DocEntry", udoId);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                Solicitud solicitud = new Solicitud();

                if (oGeneralData != null)
                {
                    // Access the UDO record data here
                    solicitud.DocEntry = oGeneralData.GetProperty("DocEntry").ToString();
                    solicitud.MGS_CL_FECHA = DateTime.Parse( oGeneralData.GetProperty("U_MGS_CL_FECHA").ToString());
                    solicitud.MGS_CL_CLIE = oGeneralData.GetProperty("U_MGS_CL_CLIE").ToString();
                    solicitud.MGS_CL_CLID = oGeneralData.GetProperty("U_MGS_CL_CLID").ToString();
                    solicitud.MGS_CL_ARTC = oGeneralData.GetProperty("U_MGS_CL_ARTC").ToString();
                    solicitud.MGS_CL_ARTD = oGeneralData.GetProperty("U_MGS_CL_ARTD").ToString();
                    solicitud.MGS_CL_PRAR = oGeneralData.GetProperty("U_MGS_CL_PRAR").ToString();
                    solicitud.MGS_CL_OFVE = oGeneralData.GetProperty("U_MGS_CL_OFVE").ToString();
                    solicitud.MGS_CL_OFVI = oGeneralData.GetProperty("U_MGS_CL_OFVI").ToString();
                    solicitud.MGS_CL_EFCO = oGeneralData.GetProperty("U_MGS_CL_EFCO").ToString();
                    solicitud.MGS_CL_MTANC = oGeneralData.GetProperty("U_MGS_CL_MTANC").ToString();
                    solicitud.MGS_CL_MTLAR = oGeneralData.GetProperty("U_MGS_CL_MTLAR").ToString();
                    solicitud.MGS_CL_MTCANT = oGeneralData.GetProperty("U_MGS_CL_MTCANT").ToString();
                    solicitud.MGS_CL_OFEV = oGeneralData.GetProperty("U_MGS_CL_OFEV").ToString();
                    solicitud.MGS_CL_AFECOR = DateTime.Parse( oGeneralData.GetProperty("U_MGS_CL_AFECOR").ToString());

                    oChildren = oGeneralData.Child("MGS_CL_COMAST");
                    var sss = oChildren.Count;
                    List<SolicitudDetalle> detalleSolicitud = new List<SolicitudDetalle>();
                    foreach (SAPbobsCOM.GeneralData detailRecord in oChildren)
                    {
                        SolicitudDetalle det = new SolicitudDetalle();
                        det.MGS_CL_LOTE = detailRecord.GetProperty("U_MGS_CL_LOTE").ToString();
                        det.MGS_CL_ANCM = detailRecord.GetProperty("U_MGS_CL_ANCM").ToString();
                        det.MGS_CL_MLAR = detailRecord.GetProperty("U_MGS_CL_MLAR").ToString();
                        det.MGS_CL_MNBO = detailRecord.GetProperty("U_MGS_CL_MNBO").ToString();
                        det.MGS_CL_MCAN = detailRecord.GetProperty("U_MGS_CL_MCAN").ToString();
                        detalleSolicitud.Add(det);
                        
                        // Acceder a los campos de detalle
                    }

                    solicitud.Detalle = detalleSolicitud;

                    //oChildren = oGeneralData.Child("MGS_CL_COMAST");
                    //var sldj = oChildren.Count;
                    //foreach (SAPbobsCOM.GeneralData detailRecord in oChildren)
                    //{
                    //    Console.WriteLine("Detalle - Nombre del UDO: " + detailRecord.GetProperty("LineId"));
                    //    // Acceder a los campos de detalle
                    //}

                    //oChildren = oGeneralData.Child("MGS_CL_CORESU");
                    //var asdasd = oChildren.Count;

                    //foreach (SAPbobsCOM.GeneralData detailRecord in oChildren)
                    //{
                    //    Console.WriteLine("Detalle - Nombre del UDO: " + detailRecord.GetProperty("LineId"));
                    //    // Acceder a los campos de detalle
                    //}
                    //return true;
                }

                return solicitud;
            }
            catch (Exception ex)
            {
                //Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                return null;
            }
        }

        #endregion
    }
}
