using AddOnCorte.Clases;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

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

        public static void UpdateUDORecibo(string udoId, string docEntryOf, string fieldUpdate)
        {

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = Globales.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("MGS_CL_RCOCABE");
                oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
                oGeneralParams.SetProperty("DocEntry", udoId);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty(fieldUpdate, docEntryOf);
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
                    solicitud.MGS_CL_RESFILE = oGeneralData.GetProperty("U_MGS_CL_RESFILE").ToString() == "N"?false: true;
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
                        det.MGS_CL_ALMC = detailRecord.GetProperty("U_MGS_CL_ALMC").ToString();
                        det.MGS_CL_FECADM = detailRecord.GetProperty("U_MGS_CL_FECADM").ToString();
                        det.MGS_CL_FIFO = detailRecord.GetProperty("U_MGS_CL_FIFO").ToString();
                        detalleSolicitud.Add(det);
                        
                        // Acceder a los campos de detalle
                    }

                    solicitud.Detalle = detalleSolicitud;

                    oChildren = oGeneralData.Child("MGS_CL_CORRID");
                    var sldj = oChildren.Count;
                    List<CorridasDetalle> detalleCorrida = new List<CorridasDetalle>();
                    foreach (SAPbobsCOM.GeneralData detailRecord in oChildren)
                    {
                        //Console.WriteLine("Detalle - Nombre del UDO: " + detailRecord.GetProperty("LineId"));
                        // Acceder a los campos de detalle
                        CorridasDetalle det = new CorridasDetalle();
                        det.MGS_CL_TITU = detailRecord.GetProperty("U_MGS_CL_TITU").ToString();
                        det.MGS_CL_C1 = detailRecord.GetProperty("U_MGS_CL_C1").ToString();
                        det.MGS_CL_C2 = detailRecord.GetProperty("U_MGS_CL_C2").ToString();
                        det.MGS_CL_C3 = detailRecord.GetProperty("U_MGS_CL_C3").ToString();
                        det.MGS_CL_C4 = detailRecord.GetProperty("U_MGS_CL_C4").ToString();
                        det.MGS_CL_C5 = detailRecord.GetProperty("U_MGS_CL_C5").ToString();
                        det.MGS_CL_C6 = detailRecord.GetProperty("U_MGS_CL_C6").ToString();
                        det.MGS_CL_C7 = detailRecord.GetProperty("U_MGS_CL_C7").ToString();
                        det.MGS_CL_C8 = detailRecord.GetProperty("U_MGS_CL_C8").ToString();
                        det.MGS_CL_C9 = detailRecord.GetProperty("U_MGS_CL_C9").ToString();
                        det.MGS_CL_C10 = detailRecord.GetProperty("U_MGS_CL_C10").ToString();
                        detalleCorrida.Add(det);
                    }

                    solicitud.DetalleCorridas = detalleCorrida;

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

        public static (bool, string) GetJsonValue(string xmlResult, string root)
        {
            XDocument xdoc = XDocument.Parse(xmlResult);

            // Encontrar el elemento NTNT y obtener su XML interno
            XElement ntntElement = xdoc.Descendants(root).FirstOrDefault();
            string ntntXml = "";
            string jsonString = "";
            bool isList = true;
            if (ntntElement != null)
            {
                // Convertir el elemento NTNT a cadena de XML
                ntntXml = ntntElement.ToString();
                int rowCount = ntntElement.Elements("row").Count();

                if (rowCount > 1)
                {
                    string xmlWithoutNilAttributes = RemoveNilAttributes(ntntXml);

                    jsonString = ConvertXmlToJson(xmlWithoutNilAttributes);
                    isList = true;
                }
                else
                {
                    //XDocument xdocRow = XDocument.Parse(ntntXml);
                    //XElement ntntElementRow = xdocRow.Descendants("Row").FirstOrDefault();

                    XElement ntntElementRow = ntntElement.Element("row");

                    if (ntntElementRow != null)
                    {
                        string xmlWithoutNilAttributes = RemoveNilAttributes(ntntElementRow.ToString());

                        jsonString = ConvertXmlToJson(xmlWithoutNilAttributes);
                        isList = false;
                    }

                }

            }



            return (isList, jsonString);
        }

        static string RemoveNilAttributes(string xml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);

            // Eliminar atributos nil="true"
            var nodesWithNilAttributes = xmlDoc.SelectNodes("//*[@nil='true']");
            foreach (XmlNode node in nodesWithNilAttributes)
            {
                var nilAttribute = node.Attributes["nil"];
                if (nilAttribute != null)
                {
                    node.Attributes.Remove(nilAttribute);
                }
            }

            return xmlDoc.OuterXml;
        }

        static string ConvertXmlToJson(string xml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);

            // Convertir XmlDocument a JSON
            string jsonString = JsonConvert.SerializeXmlNode(xmlDoc, Newtonsoft.Json.Formatting.Indented);

            return jsonString;
        }

        private static bool RegisterLotesUDO2_(SAPbobsCOM.Documents docu)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbobsCOM.Recordset oRecordSet = null;
            SAPbouiCOM.DBDataSource dBDataSource = null;
            SAPbobsCOM.Documents oSalidas = null;
            try
            {
                //oForm = Globales.oApp.Forms.Item(formUID);
                oRecordSet = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                oSalidas = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                dBDataSource = oForm.DataSources.DBDataSources.Item("");

                string docEntry =  dBDataSource.GetValue("DocEntry", 0);
                string docNum = dBDataSource.GetValue("DocNum", 0);
                string docType = dBDataSource.GetValue("DocType", 0);
                string objType = dBDataSource.GetValue("ObjType", 0);
                string canceled = dBDataSource.GetValue("CANCELED", 0);

                if (oSalidas.GetByKey(int.Parse(docEntry))) // Replace 123 with the DocEntry of an existing document if you want to update it
                {
                    string xmlResult = oSalidas.GetAsXML();

                    var resultado1 = FuncionesComunes.GetJsonValue(xmlResult, "BTNT");
                    var resultado2 = FuncionesComunes.GetJsonValue(xmlResult, "IGE19");

                    int valueFinalQty = canceled == "N" ? 1 : -1;

                    BO lote = new BO();
                    BTNT bTNT = new BTNT();
                    if (resultado1.Item1)
                        lote = JsonConvert.DeserializeObject<BO>(resultado1.Item2);
                    else
                    {
                        DOC docs = JsonConvert.DeserializeObject<DOC>(resultado1.Item2);
                        List<Row> rows = new List<Row>();
                        rows.Add(docs.row);
                        bTNT.row = rows;
                        lote.BTNT = bTNT;
                    }


                    BO ub = new BO();
                    IGE19 iGE19 = new IGE19();
                    if (resultado2.Item1)
                        ub = JsonConvert.DeserializeObject<BO>(resultado2.Item2);
                    else
                    {
                        DOC docs = JsonConvert.DeserializeObject<DOC>(resultado2.Item2);
                        List<Row> rows = new List<Row>();
                        rows.Add(docs.row);
                        iGE19.row = rows;
                        ub.IGE19 = iGE19;
                    }



                    int contador = 0;
                    List<int> linesP = new List<int>();



                    for (int i = 0; i < lote.BTNT.row.Count; i++) //lote.BOM.BO.BTNT.row.Count
                    {
                        var loteRegister = lote.BTNT.row[i]; //lote.BOM.BO.BTNT.row[i];

                        if (!linesP.Contains(loteRegister.DocLineNum))
                            contador = 0;
                        else
                            contador++;


                        oSalidas.Lines.SetCurrentLine(loteRegister.DocLineNum);

                        Lote lt = new Lote();
                        lt.MGS_CL_ARTICL = loteRegister.ItemCode;
                        lt.MGS_CL_IDLOTE = loteRegister.DistNumber;
                        lt.MGS_CL_CANTID = double.Parse(loteRegister.Quantity) * -1 * valueFinalQty;
                        //lt.MGS_CL_CANTID = oSalidas.Lines.Quantity;
                        lt.MGS_CL_LINNUM = loteRegister.DocLineNum.ToString();
                        lt.MGS_CL_CODWHS = loteRegister.WhsCode;


                        lt.MGS_CL_ANCHO = double.Parse(oSalidas.Lines.UserFields.Fields.Item("U_MGS_CL_ANCHO").Value.ToString());
                        lt.MGS_CL_LARGO = double.Parse(oSalidas.Lines.UserFields.Fields.Item("U_MGS_CL_LARGO").Value.ToString());
                        // lt.MGS_CL_NUMBOB = oSalidas.Lines.UserFields.Fields.Item("U_MGS_CL_CANBOB").Value*-1;

                        lt.MGS_CL_NUMBOB = Math.Round(lt.MGS_CL_CANTID * 10000 / (lt.MGS_CL_ANCHO * lt.MGS_CL_LARGO * 2.54 * 2.54 * 12), 1);

                        linesP.Add(loteRegister.DocLineNum);
                        List<Ubicacion> ubs = new List<Ubicacion>();

                        // var ubicacionesC = lote.BOM.BO.IGE19.row.Where(a => a.LineNum == loteRegister.DocLineNum).OrderBy(a => a.SnBMDAbs).ToList();
                        // var asdasd = lote.BOM.BO.DLN19;
                        if (ub != null)
                        {
                            var ubicacionesC = ub.IGE19.row.Where(a => a.LineNum == loteRegister.DocLineNum).OrderBy(a => a.SnBMDAbs).ToList();
                            var ubis = ubicacionesC.Where(b => b.SnBMDAbs == contador).ToList();

                            foreach (var u in ubis)
                            {
                                Ubicacion ubicacion = new Ubicacion();
                                ubicacion.MGS_CL_UBICAC = u.BinAbs.ToString();
                                ubicacion.MGS_CL_ARTICL = lt.MGS_CL_ARTICL;
                                ubicacion.MGS_CL_ANCHO = lt.MGS_CL_ANCHO;
                                ubicacion.MGS_CL_LARGO = lt.MGS_CL_LARGO;
                                ubicacion.MGS_CL_IDLOTE = lt.MGS_CL_IDLOTE;
                                ubicacion.MGS_CL_CANTID = double.Parse(u.Quantity.ToString()) * -1 * valueFinalQty;
                                ubicacion.MGS_CL_NUMBOB = Math.Round(ubicacion.MGS_CL_CANTID * 10000 / (ubicacion.MGS_CL_ANCHO * ubicacion.MGS_CL_LARGO * 2.54 * 2.54 * 12), 1);
                                ubs.Add(ubicacion);
                            }
                        }



                        GuardarUDO(lt, null, docEntry, oSalidas.DocNum.ToString(), docType, ubs, objType);


                    }

                }

                return true;

            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static bool RegisterLotesUDO3_(SAPbobsCOM.Documents docu, string tipoDetail)
        {
            //SAPbouiCOM.Form oForm = null;
            SAPbobsCOM.Recordset oRecordSet = null;
            SAPbouiCOM.DBDataSource dBDataSource = null;
            //   SAPbobsCOM.Documents oSalidas = null;
            try
            {
                //oForm = Globales.oApp.Forms.Item(formUID);
                oRecordSet = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                //oSalidas = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                //dBDataSource = oForm.DataSources.DBDataSources.Item(table);

                string docEntry = docu.DocEntry.ToString(); // dBDataSource.GetValue("DocEntry", 0);
                string docNum = docu.DocNum.ToString(); // dBDataSource.GetValue("DocNum", 0);
                string docType = docu.DocType == SAPbobsCOM.BoDocumentTypes.dDocument_Items ? "I" : "S";// .ToString(); // dBDataSource.GetValue("DocType", 0);
                string objType = docu.DocObjectCode == SAPbobsCOM.BoObjectTypes.oInventoryGenExit?"60": docu.DocObjectCode == SAPbobsCOM.BoObjectTypes.oInventoryGenEntry?"59":"15";

                string canceled = docu.Cancelled == SAPbobsCOM.BoYesNoEnum.tNO ? "N" : "Y"; // .  ToString(); // dBDataSource.GetValue("CANCELED", 0);

                string xmlResult = docu.GetAsXML();

                var resultado1 = FuncionesComunes.GetJsonValue(xmlResult, "BTNT");
                var resultado2 = FuncionesComunes.GetJsonValue(xmlResult, tipoDetail);

                int valueFinalQty = canceled == "N" ? 1 : -1;

                BO lote = new BO();
                BTNT bTNT = new BTNT();
                if (resultado1.Item1)
                    lote = JsonConvert.DeserializeObject<BO>(resultado1.Item2);
                else
                {
                    DOC docs = JsonConvert.DeserializeObject<DOC>(resultado1.Item2);
                    List<Row> rows = new List<Row>();
                    rows.Add(docs.row);
                    bTNT.row = rows;
                    lote.BTNT = bTNT;
                }


                BO ub = new BO();
                IGE19 iGE19 = new IGE19();
                IGN19 iGN19 = new IGN19();
                DLN19 dLN19 = new DLN19();

                if (resultado2.Item1)
                    ub = JsonConvert.DeserializeObject<BO>(resultado2.Item2);
                else
                {
                    DOC docs = JsonConvert.DeserializeObject<DOC>(resultado2.Item2);
                    List<Row> rows = new List<Row>();
                    rows.Add(docs.row);

                    if (tipoDetail == "IGN19")
                    {
                        iGN19.row = rows;
                        ub.IGN19 = iGN19;
                    }
                    else if (tipoDetail == "IGE19")
                    {
                        iGE19.row = rows;
                        ub.IGE19 = iGE19;
                    }
                    else
                    {
                        dLN19.row = rows;
                        ub.DLN19 = dLN19;
                    }

                    

                }



                int contador = 0;
                List<int> linesP = new List<int>();



                for (int i = 0; i < lote.BTNT.row.Count; i++) //lote.BOM.BO.BTNT.row.Count
                {
                    var loteRegister = lote.BTNT.row[i]; //lote.BOM.BO.BTNT.row[i];

                    if (!linesP.Contains(loteRegister.DocLineNum))
                        contador = 0;
                    else
                        contador++;


                    docu.Lines.SetCurrentLine(loteRegister.DocLineNum);

                    Lote lt = new Lote();
                    lt.MGS_CL_ARTICL = loteRegister.ItemCode;
                    lt.MGS_CL_IDLOTE = loteRegister.DistNumber;
                    lt.MGS_CL_CANTID = double.Parse(loteRegister.Quantity) * valueFinalQty * (objType != "59"?-1:1);
                    //lt.MGS_CL_CANTID = oSalidas.Lines.Quantity;
                    lt.MGS_CL_LINNUM = loteRegister.DocLineNum.ToString();
                    lt.MGS_CL_CODWHS = loteRegister.WhsCode;


                    lt.MGS_CL_ANCHO = double.Parse(docu.Lines.UserFields.Fields.Item("U_MGS_CL_ANCHO").Value.ToString());
                    lt.MGS_CL_LARGO = double.Parse(docu.Lines.UserFields.Fields.Item("U_MGS_CL_LARGO").Value.ToString());
                    // lt.MGS_CL_NUMBOB = oSalidas.Lines.UserFields.Fields.Item("U_MGS_CL_CANBOB").Value*-1;

                    lt.MGS_CL_NUMBOB = Math.Round(lt.MGS_CL_CANTID * 10000 / (lt.MGS_CL_ANCHO * lt.MGS_CL_LARGO * 2.54 * 2.54 * 12), 1);

                    linesP.Add(loteRegister.DocLineNum);
                    List<Ubicacion> ubs = new List<Ubicacion>();

                    if (ub != null)
                    {
                        List<Row> ubicacionesC = new List<Row>();

                        if(tipoDetail == "IGN19")
                            ubicacionesC = ub.IGN19.row.Where(a => a.LineNum == loteRegister.DocLineNum).OrderBy(a => a.SnBMDAbs).ToList();
                        else if(tipoDetail == "IGE19")
                            ubicacionesC = ub.IGE19.row.Where(a => a.LineNum == loteRegister.DocLineNum).OrderBy(a => a.SnBMDAbs).ToList();
                        else
                            ubicacionesC = ub.DLN19.row.Where(a => a.LineNum == loteRegister.DocLineNum).OrderBy(a => a.SnBMDAbs).ToList();


                        var ubis = ubicacionesC.Where(b => b.SnBMDAbs == contador).ToList();

                        foreach (var u in ubis)
                        {
                            Ubicacion ubicacion = new Ubicacion();
                            ubicacion.MGS_CL_UBICAC = u.BinAbs.ToString();
                            ubicacion.MGS_CL_ARTICL = lt.MGS_CL_ARTICL;
                            ubicacion.MGS_CL_ANCHO = lt.MGS_CL_ANCHO;
                            ubicacion.MGS_CL_LARGO = lt.MGS_CL_LARGO;
                            ubicacion.MGS_CL_IDLOTE = lt.MGS_CL_IDLOTE;
                            ubicacion.MGS_CL_CANTID = double.Parse(u.Quantity.ToString())  * valueFinalQty* (objType != "59" ? -1 : 1);
                            ubicacion.MGS_CL_NUMBOB = Math.Round(ubicacion.MGS_CL_CANTID * 10000 / (ubicacion.MGS_CL_ANCHO * ubicacion.MGS_CL_LARGO * 2.54 * 2.54 * 12), 1);
                            ubs.Add(ubicacion);
                        }
                    }



                    GuardarUDO(lt, null, docEntry, docu.DocNum.ToString(), docType, ubs, objType);


                }

                return true;

            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static bool GuardarUDO( Lote lt, Detalle dt, string docEntry, string docNum, string docType, List<Ubicacion> ub, string objType)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string s_Message = "";
            string s_CardCode = "";
            string s_TotalPago = "";

            try
            {
                oCompanyService = Globales.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("MGS_CL_LOTBOC");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));


                oGeneralData.SetProperty("U_MGS_CL_FECCON", DateTime.Now.ToShortDateString());
                oGeneralData.SetProperty("U_MGS_CL_ARTICL", lt.MGS_CL_ARTICL);
                oGeneralData.SetProperty("U_MGS_CL_ANCHO", lt.MGS_CL_ANCHO);
                oGeneralData.SetProperty("U_MGS_CL_LARGO", lt.MGS_CL_LARGO);
                oGeneralData.SetProperty("U_MGS_CL_IDLOTE", lt.MGS_CL_IDLOTE);
                oGeneralData.SetProperty("U_MGS_CL_CANTID", lt.MGS_CL_CANTID);
                oGeneralData.SetProperty("U_MGS_CL_NUMBOB", lt.MGS_CL_NUMBOB);
                oGeneralData.SetProperty("U_MGS_CL_CODWHS", lt.MGS_CL_CODWHS);
                oGeneralData.SetProperty("U_MGS_CL_LINNUM", lt.MGS_CL_LINNUM);
                oGeneralData.SetProperty("U_MGS_CL_DOCTYP", docType);
                oGeneralData.SetProperty("U_MGS_CL_DOCENT", docEntry);
                oGeneralData.SetProperty("U_MGS_CL_DOCNUM", docNum);
                oGeneralData.SetProperty("U_MGS_CL_OBJTYP", objType);
                oGeneralData.SetProperty("U_MGS_CL_REUSER", Globales.oCompany.UserName);


                oChildren = oGeneralData.Child("MGS_CL_LOTBOU");

                foreach (Ubicacion b in ub)
                {
                    oChild = oChildren.Add();
                    oChild.SetProperty("U_MGS_CL_UBICAC", b.MGS_CL_UBICAC);
                    oChild.SetProperty("U_MGS_CL_ARTICL", b.MGS_CL_ARTICL);
                    oChild.SetProperty("U_MGS_CL_ANCHO", b.MGS_CL_ANCHO);
                    oChild.SetProperty("U_MGS_CL_LARGO", b.MGS_CL_LARGO);
                    oChild.SetProperty("U_MGS_CL_IDLOTE", b.MGS_CL_IDLOTE);
                    oChild.SetProperty("U_MGS_CL_CANTID", b.MGS_CL_CANTID);
                    oChild.SetProperty("U_MGS_CL_NUMBOB", b.MGS_CL_NUMBOB);

                }


                oGeneralParams = oGeneralService.Add(oGeneralData);
                //s_Message = "Se ha guardado la informacion";
                //Conexion.Conexion_SBO.m_SBO_Appl.MessageBox(s_Message, 1, "Aceptar", "", "");
                return true;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                return false;
            }
        }

        #endregion
    }
}
