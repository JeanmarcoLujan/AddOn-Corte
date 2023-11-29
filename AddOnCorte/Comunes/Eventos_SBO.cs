using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Comunes
{
    class Eventos_SBO
    {
        #region Constructores

        /// <summary>
        /// Constructor de la clase.
        /// </summary>
        public Eventos_SBO()
        {
            try
            {
                //Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title = Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title.Replace(" #" + Conexion.Conexion_SBO.m_SBO_Appl.AppId.ToString(), "");
                //Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title = Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title + " #" + Conexion.Conexion_SBO.m_SBO_Appl.AppId.ToString();
                RegistrarEventos();
                //RegistrarFiltros();
               // RegistrarMenu();
                //Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(MSS_Asistente_PMD.Properties.Resources.NombreAddon + " Conectado con exito",
                //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }

        #endregion

        #region Eventos

        //void m_SBO_Appl_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        //{
        //    try
        //    {
        //        switch (EventType)
        //        {
        //            case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
        //            case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
        //            case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
        //                Application.Exit();
        //                break;
        //            case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
        //            case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
        //                RegistrarMenu();
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
        //    }
        //}
        void m_SBO_Appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                {
                    case "149":
                    case "139":
                    case "140":
                    case "1250000940":
                    case "720":
                    case "721":
                        Forms.Frm_149 oFrm_149 = null;
                        oFrm_149 = new AddOnCorte.Forms.Frm_149();
                        oFrm_149.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        oFrm_149 = null;
                        break;
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }
        void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.BeforeAction)
                {
                    case false:
                        switch (pVal.MenuUID)
                        {
                            //case "MSS_APMD":
                            //    Formularios.Frm_APMD_P1 oFrm_APMD_P1 = null;
                            //    oFrm_APMD_P1 = new MSS_Asistente_PMD.Formularios.Frm_APMD_P1(true);
                            //    oFrm_APMD_P1 = null;
                            //    break;
                            //case "MSS_CPMD":
                            //    Formularios.Frm_MSSL_PMD oFrm_MSSL_PMD = null;
                            //    oFrm_MSSL_PMD = new MSS_Asistente_PMD.Formularios.Frm_MSSL_PMD(true);
                            //    oFrm_MSSL_PMD = null;
                                //break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        #endregion

        #region Metodos

        /// <summary>
        /// Método para registrar la opción del menú dentro del formulario de menus de SAP B1.
        /// </summary>
        //private void RegistrarMenu()
        //{
        //    try
        //    {
        //        //CreaMenu("MSS_APMD", "Asistente Pagos Masivos y Detracciones", "43538", SAPbouiCOM.BoMenuType.mt_STRING);
        //        CreaMenu("MSS_MPMD", "Pagos Masivos y Detracciones", "43538", SAPbouiCOM.BoMenuType.mt_POPUP); //43520
        //        CreaMenu("MSS_CPMD", "Definiciones PMD", "MSS_MPMD", SAPbouiCOM.BoMenuType.mt_STRING);
        //        CreaMenu("MSS_APMD", "Asistente Pagos Masivos y Detracciones", "MSS_MPMD", SAPbouiCOM.BoMenuType.mt_STRING);//43538

        //    }
        //    catch (Exception ex)
        //    {
        //        FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
        //    }
        //}

        /// <summary>
        /// Método para registrar los eventos de la aplicación en SAP B1.
        /// </summary>
        private void RegistrarEventos()
        {
            try
            {
              //  Clases.Globales.oApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(m_SBO_Appl_AppEvent);
                Clases.Globales.oApp.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(m_SBO_Appl_FormDataEvent);
                Clases.Globales.oApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(m_SBO_Appl_ItemEvent);
                Clases.Globales.oApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(m_SBO_Appl_MenuEvent);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar los filtros en SAP B1.
        /// </summary>
        //private void RegistrarFiltros()
        //{
        //    SAPbouiCOM.EventFilter oEF = null;
        //    SAPbouiCOM.EventFilters oEFs = null;
        //    try
        //    {
        //        oEFs = new SAPbouiCOM.EventFilters();
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
        //        oEF.AddEx("Frm_APMD_P1");
        //        oEF.AddEx("Frm_APMD_P2");
        //        oEF.AddEx("Frm_APMD_P3PP");
        //        oEF.AddEx("Frm_APMD_P3PD");
        //        oEF.AddEx("Frm_APMD_P3PAD");
        //        oEF.AddEx("Frm_APMD_P4");
        //        oEF.AddEx("Frm_APMD_P4PR");
        //        oEF.AddEx("Frm_APMD_P5");
        //        oEF.AddEx("Frm_APMD_P6");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
        //        oEF.AddEx("Frm_APMD_P1");
        //        oEF.AddEx("Frm_APMD_P2");
        //        oEF.AddEx("65052");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
        //        oEF.AddEx("Frm_APMD_P1");
        //        oEF.AddEx("Frm_APMD_P2");
        //        oEF.AddEx("Frm_APMD_P3PP");
        //        oEF.AddEx("Frm_APMD_P3PD");
        //        oEF.AddEx("Frm_APMD_P3PAD");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
        //        oEF.AddEx("65052");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
        //        oEF.AddEx("Frm_APMD_P1");
        //        oEF.AddEx("Frm_APMD_P2");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
        //        oEF.AddEx("65052");
        //        oEF.AddEx("Frm_MSSL_PMD");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
        //        oEF.AddEx("65052");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
        //        oEF.AddEx("Frm_APMD_P3PP");
        //        oEF.AddEx("Frm_APMD_P3PD");
        //        oEF.AddEx("Frm_APMD_P3PAD");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
        //        oEF.AddEx("Frm_APMD_P3PP");
        //        oEF.AddEx("Frm_APMD_P3PD");
        //        oEF.AddEx("Frm_APMD_P3PAD");
        //        oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
        //        oEF.AddEx("Frm_MSSL_PMD");
        //        Conexion.Conexion_SBO.m_SBO_Appl.SetFilter(oEFs);
        //    }
        //    catch (Exception ex)
        //    {
        //        FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
        //    }
        //}

        /// <summary>
        /// Método para registrar opciones de menú en el menu principal de SAP B1.
        /// </summary>
        /// <param name="uniqueId"></param>
        /// <param name="name"></param>
        /// <param name="principalMenuId"></param>
        /// <param name="type"></param>
        //private void CreaMenu(string uniqueId, string name, string principalMenuId, SAPbouiCOM.BoMenuType type)
        //{
        //    SAPbouiCOM.MenuCreationParams objParams;
        //    SAPbouiCOM.Menus objSubMenu;

        //    try
        //    {
        //        objSubMenu = Conexion.Conexion_SBO.m_SBO_Appl.Menus.Item(principalMenuId).SubMenus;

        //        if (Conexion.Conexion_SBO.m_SBO_Appl.Menus.Exists(uniqueId) == false)
        //        {
        //            objParams = (SAPbouiCOM.MenuCreationParams)Conexion.Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
        //            objParams.Type = type;
        //            objParams.UniqueID = uniqueId;
        //            objParams.String = name;
        //            objParams.Position = -1;
        //            objSubMenu.AddEx(objParams);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
        //    }
        //}

        #endregion
    }
}
