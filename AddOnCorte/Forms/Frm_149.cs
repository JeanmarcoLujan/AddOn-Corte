﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Forms
{
    class Frm_149
    {
        public const string csFormType = "149";

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

                oItem = oForm.Items.Add("btnProceso", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = captuon_btn;


                oItem.Top = 120;
                oItem.Left = (oItem.Width);
                oItem.Width = 150;

                switch (oForm.Mode)
                {

                    case SAPbouiCOM.BoFormMode.fm_EDIT_MODE:
                    case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                    case SAPbouiCOM.BoFormMode.fm_OK_MODE:
                        oItem.Enabled = true;
                        break;
                    case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
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
    }
}