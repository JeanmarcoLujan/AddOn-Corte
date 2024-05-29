using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Forms
{
    public class Frm_149
    {
        public const string csFormType = "149";
        public static string corteCode = "";
        public void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try 
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                       // case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                            DubujarBoton(FormUID, pVal.FormTypeEx);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.FormTypeEx == "149" || pVal.FormTypeEx == "139" || pVal.FormTypeEx == "140" || pVal.FormTypeEx == "720" || pVal.FormTypeEx == "721" || pVal.FormTypeEx == "1250000940")
                            {
                                if (pVal.ItemUID == "btnProceso")
                                {
                                    OpenSolcitudCorte(FormUID, pVal.FormTypeEx);
                                    //Clases.Globales.oApp.MessageBox("Mostrar la solcitud de corte ");
                                }

                            }
                            break;
                            //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            //    switch (pVal.ItemUID)
                            //    {
                            //        case "1":
                            //            NumeroFiscalCambio(FormUID);
                            //            BubbleEvent = ValidaInsert(FormUID);
                            //            break;
                            //        default:
                            //            break;
                            //    }
                            //    break;
                    }
                }
                else
                {

                    //switch (pVal.EventType)
                    //{
                    //    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                    //        if (pVal.ActionSuccess)
                    //        {
                    //            BubbleEvent = NumeroFiscal(FormUID);
                    //        }

                    //        break;
                    //    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                    //        var sdfsdf = pVal.ActionSuccess;
                    //        var sdfsdfsss = pVal.ItemUID;
                    //        if (!pVal.ActionSuccess)
                    //        {
                    //            if (pVal.ItemUID.Equals("4"))
                    //            {
                    //                BubbleEvent = NumeroFiscal(FormUID);
                    //            }
                    //        }
                    //        else
                    //        {
                    //            if (pVal.ItemUID.Equals("8"))
                    //                HabDeshabSerie(FormUID, true);
                    //            else if (pVal.ItemUID.Equals("16") || pVal.ItemUID.Equals("4"))
                    //                HabDeshabSerie(FormUID, false);
                    //        }
                    //        break;
                    //    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                    //        if (pVal.ActionSuccess)
                    //        {
                    //            if (pVal.ItemUID.Equals("88"))
                    //            {
                    //                BubbleEvent = NumeroFiscal(FormUID);
                    //            }
                    //        }
                    //        break;
                    //}
                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                //ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
            }
        }

        public void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, string ItemUID = "", int row = -1, bool proceso = false)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = null;
            try
            {
                var sdfsdf = pVal.MenuUID;
                switch (pVal.MenuUID)
                {
                    case "1284":
                        if (proceso)
                        {
                            oForm = Clases.Globales.oApp.Forms.ActiveForm; //  Conexion.Conexion_SBO.m_SBO_Appl.Forms.ActiveForm;

                            var asdasd = oForm.TypeEx;

                            if (oForm.TypeEx == "149" || oForm.TypeEx == "139")
                            {
                                // BubbleEvent = FacturacionNX.LiberarDocument(oForm, 13);
                                BubbleEvent =  CancelSolcitudCorte(oForm.UDFFormUID, oForm.TypeEx); //GetInfo(oForm.UDFFormUID, "", bFormDataADDevent: true, csFormTable);
                            }

                        }

                        break;
                }
            }
            catch (Exception ex)
            {
                //ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
            }
        }


        public void m_SBO_Appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                switch (BusinessObjectInfo.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        //case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK:

                        if (!BusinessObjectInfo.BeforeAction)
                        {
                            if (BusinessObjectInfo.ActionSuccess)
                            {
                                if (corteCode != "")
                                {
                                    Comunes.FuncionesComunes.UpdateUDO(corteCode, "C", "U_MGS_CL_ESTD");
                                    corteCode = "";
                                }
                                    //BubbleEvent = Comunes.FuncionesComunes.UpdateUDO(corteCode, "C", "U_MGS_CL_ESTD"); // RegisterLotesUDO1_(BusinessObjectInfo.FormUID, BusinessObjectInfo.FormTypeEx, bFormDataADDevent: true, csFormTable, false);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
               // FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                BubbleEvent = false;
            }
        }


        private void DubujarBoton(string s_FormUID, string formTypeEx)
        {
            SAPbouiCOM.Form oForm = null;


            try
            {
                string captuon_btn = "";
                switch (formTypeEx)
                {
                    case "720":
                    case "721":
                    case "140":
                        captuon_btn = "Recibo producción";
                        break;
                    default:
                        captuon_btn = "Solicitud corte";
                        break;
                }


                oForm = Clases.Globales.oApp.Forms.Item(s_FormUID);

                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Button oButton = null;

                SAPbouiCOM.Item oItem1 = null;
                SAPbouiCOM.Button oButton1 = null;
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbouiCOM.ComboBox oCbo = null;

                oItem = oForm.Items.Add("btnProceso", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = captuon_btn;

                if (formTypeEx == "720" || formTypeEx == "721")
                {

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                    var asdas = oMatrix.Item.Top;

                    oCbo = (SAPbouiCOM.ComboBox)oForm.Items.Item("3").Specific;
                    var asdas1 = oCbo.Item.Top;

                    oItem.Top = asdas1 + (asdas - asdas1 -5)/2;
                    oItem.Left = (oItem.Width);
                    oItem.Width = 150;



                }
                else
                {
                    oItem.Top = 120;
                    oItem.Left = (oItem.Width);
                    oItem.Width = 150;
                }
                

                switch (oForm.Mode)
                {

                    case SAPbouiCOM.BoFormMode.fm_EDIT_MODE:
                    case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                    case SAPbouiCOM.BoFormMode.fm_OK_MODE:
                    case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                        oItem.Enabled = true;
                        break;
                    
                    case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                    case SAPbouiCOM.BoFormMode.fm_PRINT_MODE:
                    case SAPbouiCOM.BoFormMode.fm_VIEW_MODE:
                        oItem.Enabled = false;
                        break;
                }

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private void OpenSolcitudCorte(string s_FormUID, string formEx)
        {

            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.DBDataSource dBDataSource = null;
            try
            {
                oForm = Clases.Globales.oApp.Forms.Item(s_FormUID);
                string table = "";
                bool solicitud = true;
                switch (formEx)
                {
                    case "149":
                        table = "OQUT";
                        solicitud = true;
                        break;
                    case "140":
                        table = "ODLN";
                        solicitud = false;
                        break;
                    case "139":
                        table = "ORDR";
                        solicitud = true;
                        break;
                    case "720":
                        table = "OIGE";
                        solicitud = false;
                        break;
                    case "721":
                        table = "OIGN";
                        solicitud = false;
                        break;
                    case "1250000940":
                        table = "OWTQ";
                        solicitud = true;
                        break;
                }

                dBDataSource = oForm.DataSources.DBDataSources.Item(table);

                if (solicitud)
                {
                    string corteCode = dBDataSource.GetValue("U_MGS_CL_SOLCOR", 0);
                    if (corteCode != "")
                    {
                        Form1 activeForm = new Form1(int.Parse(corteCode));
                        activeForm.Show();
                    }
                    else
                    {
                        Clases.Globales.oApp.MessageBox("El documento no tiene solicitud de corte asociada.");
                    }
                }
                else
                {
                    string reciboProd = dBDataSource.GetValue("U_MGS_CL_REFREC", 0);
                    if (reciboProd != "")
                    {
                        Form3 activeForm = new Form3(int.Parse(reciboProd));
                        activeForm.Show();
                    }
                    else
                    {
                        Clases.Globales.oApp.MessageBox("El documento no tiene recibo de producción asociado.");
                    }
                }

                

                

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }
        private bool CancelSolcitudCorte(string s_FormUID, string formEx)
        {

            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.DBDataSource dBDataSource = null;
            bool rpta = true;
            try
            {
                oForm = Clases.Globales.oApp.Forms.Item(s_FormUID);
                string table = "";
                bool solicitud = true;
                switch (formEx)
                {
                    case "149":
                        table = "OQUT";
                        solicitud = true;
                        break;
                    case "139":
                        table = "ORDR";
                        solicitud = true;
                        break;
                }

                dBDataSource = oForm.DataSources.DBDataSources.Item(table);

                if (solicitud)
                {
                    corteCode = dBDataSource.GetValue("U_MGS_CL_SOLCOR", 0);
                    if (corteCode != "")
                    {
                        //Comunes.FuncionesComunes.UpdateUDO(corteCode, "C", "U_MGS_CL_ESTD");
                        rpta = true;
                        
                        //Form1 activeForm = new Form1(int.Parse(corteCode));
                        //activeForm.Show();
                    }
                    
                }





            }
            catch (Exception ex)
            {
                rpta = false;
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

            return rpta;
        }
    }
}
