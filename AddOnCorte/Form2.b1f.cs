using AddOnCorte.Clases;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOnCorte
{
    [FormAttribute("AddOnCorte.Form2", "Form2.b1f")]
    class Form2 : UserFormBase
    {
        SAPbouiCOM.Form oForm;
        public Form2()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            //Clases.Globales.oApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(this.m_SBO_Appl_MenuEvent);
            Clases.Globales.oApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(m_SBO_Appl_ItemEvent);
        }

        public void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.BeforeAction)
                {
                    //case true:
                    //    switch (pVal.EventType)
                    //    {

                    //        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                    //            switch (pVal.ItemUID)
                    //            {
                    //                case "Item_0":
                    //                    if (pVal.ColUID.Equals("Equipo"))
                    //                        AjustarComboBoxSerie(pVal.Row);
                    //                        //Globales.oApp.MessageBox("hellow" + pVal.Row);
                    //                    break;
                    //            }
                    //            break;
                    //    }
                    //    break;
                    case false:
                        switch (pVal.EventType)
                        {

                            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                                switch (pVal.ItemUID)
                                {
                                    case "Item_0":
                                        if (pVal.ColUID.Equals("Equipo"))
                                            AjustarComboBoxSerie(pVal.Row);
                                        //Globales.oApp.MessageBox("hellow" + pVal.Row);
                                        break;
                                }
                                break;
                        }
                        break;

                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            this.ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
           // this.ComboBox0.Selected.Value = "A";
            ListSolicitudes();

        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            ListSolicitudes();
        }

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;


        private void AjustarComboBox(int row, string grid, int typeC)
        {
            //SAPbouiCOM.Form oForm = null;

            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.ComboBoxColumn oCombo = null;

            SAPbobsCOM.SBObob oSboBob = null;
            SAPbobsCOM.Recordset oRS = null;

            // SAPbobsCOM.Recordset oRecSet = default(SAPbobsCOM.Recordset);
            try
            {
                if (row >= 0)
                {
                    //oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item("Frm_APMD_P3PP");
                    oForm.Freeze(true);
                    //oBanks = (SAPbobsCOM.Banks)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBanks);
                    oSboBob = (SAPbobsCOM.SBObob) Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                    //oBP = (SAPbobsCOM.BusinessPartners)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                    oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item(grid).Specific);

                    oCombo = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(typeC == 1 ?"Equipo": "AlmacenOrigen");
                    while (oCombo.ValidValues.Count != 0)
                        oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oCombo.ValidValues.Add("", "");

                    oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    
                    oRS.DoQuery( typeC == 1 ? Comunes.Consultas.GetEquipoList() : Comunes.Consultas.GetAlmacenList() );

                    if (oRS.RecordCount > 0)
                    {
                        while (oRS.EoF == false)
                        {
                            oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());

                            oRS.MoveNext();
                        }

                    }

                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private void ListSolicitudes()
        {
            SAPbobsCOM.Recordset oRS = null;
            SAPbouiCOM.Grid oGrid = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item("Item_0").Specific);
                oGrid.DataTable = oForm.DataSources.DataTables.Item("dt_gr");

                if (this.ComboBox0.Selected.Value.ToString() != "")
                {

                    if (this.ComboBox0.Selected.Value.ToString() == "A")
                    {
                        this.Button1.Item.Enabled = true;
                        this.Button2.Item.Enabled = false;
                    }
                    else
                    {
                        this.Button1.Item.Enabled = false;
                        this.Button2.Item.Enabled = true;
                    }


                    oGrid.DataTable.ExecuteQuery(Comunes.Consultas.GetSolicitudesAgenda(this.ComboBox0.Selected.Value));


                    oGrid.Columns.Item("Codigo").Editable = false;
                    oGrid.Columns.Item("Codigo").TitleObject.Caption = "Código";
                    oGrid.Columns.Item("FechaSolicitud").Editable = false;
                    oGrid.Columns.Item("FechaSolicitud").TitleObject.Caption = "Fecha de solicitud";
                    oGrid.Columns.Item("Cliente").Editable = false;
                    ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Cliente")).LinkedObjectType = "2";
                    oGrid.Columns.Item("Nombre").Editable = false;
                    oGrid.Columns.Item("Artículo").Editable = false;
                    ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Artículo")).LinkedObjectType = "4";
                    oGrid.Columns.Item("Descripción").Editable = false;
                    oGrid.Columns.Item("Ofertav").Editable = false;
                    oGrid.Columns.Item("Ofertav").TitleObject.Caption = "Oferta de venta";
                    ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Ofertav")).LinkedObjectType = "23";
                    oGrid.Columns.Item("OrdenVenta").Editable = false;
                    oGrid.Columns.Item("OrdenVenta").TitleObject.Caption = "Orden de venta";
                    ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("OrdenVenta")).LinkedObjectType = "17";
                    oGrid.Columns.Item("AlmacenOrigen").TitleObject.Caption = "Almacén origen";
                    oGrid.Columns.Item("AlmacenOrigen").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                    ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("AlmacenOrigen")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                    //oGrid.Columns.Item("FechaCorte").Type = SAPbouiCOM.BoGridColumnType.;

                    oGrid.Columns.Item("Equipo").TitleObject.Caption = "Equipo";
                    oGrid.Columns.Item("Equipo").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                    ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Equipo")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;

                    oGrid.Columns.Item("Seleccion").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    oGrid.Columns.Item("Seleccion").TitleObject.Caption = "Check";

                    oGrid.Columns.Item("DiferenciaDias").Editable = false;
                    oGrid.Columns.Item("DiferenciaDias").TitleObject.Caption = "Diferencia días";

                    oGrid.Columns.Item("Indicador").Editable = false;
                    oGrid.Columns.Item("Indicador").TitleObject.Caption = "Indicador 911";

                    oGrid.Columns.Item("Ancho").Editable = false;
                    oGrid.Columns.Item("Largo").Editable = false;
                    oGrid.Columns.Item("Cantidad(m2)").Editable = false;

                    if (this.ComboBox0.Value.ToString() == "A")
                    {
                        oGrid.Columns.Item("FechaCorte").Visible = true;
                        oGrid.Columns.Item("FechaCorte").Editable = true;
                        oGrid.Columns.Item("Equipo").Visible = true;
                        oGrid.Columns.Item("Equipo").Editable = true;
                        oGrid.Columns.Item("Serial").Visible = true;
                        oGrid.Columns.Item("Serial").Editable = true;
                        oGrid.Columns.Item("AlmacenOrigen").Visible = false;
                        AjustarComboBox(2, "Item_0", 1);
                    }
                    else
                    {
                        oGrid.Columns.Item("FechaCorte").Visible = true;
                        oGrid.Columns.Item("FechaCorte").Editable = false;
                        oGrid.Columns.Item("Equipo").Visible = true;
                        oGrid.Columns.Item("Equipo").Editable = false;
                        oGrid.Columns.Item("Serial").Visible = true;
                        oGrid.Columns.Item("Serial").Editable = false;
                        oGrid.Columns.Item("AlmacenOrigen").Visible = true;
                        AjustarComboBox(2, "Item_0", 2);
                    }

                    oGrid.AutoResizeColumns();


                }

                else
                    Globales.oApp.MessageBox("Seleccionar el estado de las solicitudes a obtener");
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }
        }

        private void AjustarComboBoxSerie(int row)
        {

            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.ComboBoxColumn oCombo = null;
            SAPbouiCOM.EditTextColumn oEditText = null;

            SAPbobsCOM.Recordset oRS = null;
            string NroCta = "";
            try
            {
                if (row >= 0)
                {

                   // oForm.Freeze(true);

                    oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item("Item_0").Specific);
                   // var asdad = oGrid.DataTable.GetValue("Equipo", oGrid.GetDataTableRowIndex(row)).ToString();
                    
                    var ff1 = oGrid.DataTable.GetValue("Equipo", row);
                    if(ff1.ToString() != "")
                    {
                        oRS.DoQuery(Comunes.Consultas.GetEquipoSerie(ff1.ToString()));

                        if (oRS.RecordCount > 0)
                        {
                            oGrid.DataTable.SetValue("Serial", row, oRS.Fields.Item(0).Value.ToString().Trim());

                        }
                    }else
                        oGrid.DataTable.SetValue("Serial", row, "");


                    // oEditText = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Serial");





                    //oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
               // oForm.Freeze(false);
               // Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {

                List<Agenda> listaAgenda = new List<Agenda>();
                int ircontando = 0;
                int incompleto = 0;
                for (int i = 0; i < oForm.DataSources.DataTables.Item("dt_gr").Rows.Count; i++)
                {

                    if (oForm.DataSources.DataTables.Item("dt_gr").GetValue("Seleccion", i).ToString().Equals("Y"))
                    {
                        ircontando++;

                        var fecha = oForm.DataSources.DataTables.Item("dt_gr").GetValue("FechaCorte", i);
                        string equipo = oForm.DataSources.DataTables.Item("dt_gr").GetValue("Equipo", i).ToString();
                        string serie = oForm.DataSources.DataTables.Item("dt_gr").GetValue("Serial", i).ToString();
                        string docEntry = oForm.DataSources.DataTables.Item("dt_gr").GetValue("Codigo", i).ToString();


                        if (fecha != null && equipo != "" && serie != "")
                        {
                            Agenda agenda = new Agenda();
                            agenda.DocEntry = docEntry;
                            agenda.Fecha = fecha.ToString();
                            agenda.Equipo = equipo;
                            agenda.Serie = serie;

                            listaAgenda.Add(agenda);
                        }
                        else
                            incompleto++;


                        if (incompleto == 0)
                        {
                            foreach (var item in listaAgenda)
                            {
                                Comunes.FuncionesComunes.UpdateUDOAgendarSolicitud(item.DocEntry, item.Fecha, item.Equipo, item.Serie);
                            }
                        }
                            
                    
                    }
                }

                if (ircontando == 0)
                    Globales.oApp.MessageBox("Debe marcar las solicitudes a agendar, no ha marcado ninguna");
                else if(incompleto!=0)
                    Globales.oApp.MessageBox("En todas las líneas que marco, debe completar la fecha, equipo y serie.");
                else
                    ListSolicitudes();

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            
            SAPbobsCOM.Recordset oRS = null;


            try
            {

                List<Agendado> listaAgenda = new List<Agendado>();
                int ircontando = 0;
                int incompleto = 0;
                List<string> listaAgendaRespuesta = new List<string>();


                listaAgendaRespuesta.Add("Resumen - Resultado");
                

                string respuestaGlobal = "";


                if (Globales.oApp.MessageBox("¿Esta Ud. seguro de generar la(s) solicitud de traslado?", 1, "Continuar", "Cancelar", "") == 1)
                {
                    for (int i = 0; i < oForm.DataSources.DataTables.Item("dt_gr").Rows.Count; i++)
                    {

                        if (oForm.DataSources.DataTables.Item("dt_gr").GetValue("Seleccion", i).ToString().Equals("Y"))
                        {
                            ircontando++;

                            var almacen = oForm.DataSources.DataTables.Item("dt_gr").GetValue("AlmacenOrigen", i);
                            string docEntry = oForm.DataSources.DataTables.Item("dt_gr").GetValue("Codigo", i).ToString();


                            if (almacen != null && almacen != "")
                            {
                                Agendado agenda = new Agendado();
                                agenda.DocEntry = docEntry;
                                agenda.Almacen = almacen.ToString();
                                listaAgenda.Add(agenda);
                            }
                            else
                                incompleto++;


                        }

                    }

                    if (incompleto == 0)
                    {
                        foreach (Agendado item in listaAgenda)
                        {
                            listaAgendaRespuesta.Add(GenerateSolicitudTransferencia(item));
                        }

                    }
                }



                if(listaAgendaRespuesta.Count > 1)
                {
                    foreach (var item in listaAgendaRespuesta)
                    {
                        respuestaGlobal = respuestaGlobal + item.ToString() + "\n";
                    }

                    Globales.oApp.MessageBox(respuestaGlobal);
                }


                

                if (ircontando == 0)
                    Globales.oApp.MessageBox("Debe marcar las solicitudes agendadas para generar la solicitud de transferencia, no ha marcado ninguna");
                else if (incompleto != 0)
                    Globales.oApp.MessageBox("En todas las líneas que marco, debe completar o seleccionar el almacén.");
                else
                    ListSolicitudes();

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }
    
        private string GenerateSolicitudTransferencia(Agendado agendado)
        {
            SAPbobsCOM.StockTransfer oTransferReq = null;
            SAPbobsCOM.Recordset oRS = null;
            int iErrCod;
            string sErrMsg = "";
            string rpta = "";
            try
            {
                oTransferReq = (SAPbobsCOM.StockTransfer)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                Solicitud sol = Comunes.FuncionesComunes.GetUDOSolicitudAgendada(agendado.DocEntry);

                oTransferReq.CardCode = sol.MGS_CL_CLIE.ToString();
                oTransferReq.DocDate = DateTime.Now;
                oTransferReq.TaxDate = DateTime.Now;
                oTransferReq.DueDate = sol.MGS_CL_AFECOR;
                oTransferReq.UserFields.Fields.Item("U_MGS_CL_SOLCOR").Value = agendado.DocEntry;
                oTransferReq.UserFields.Fields.Item("U_MGS_CL_EFCO").Value = sol.MGS_CL_EFCO;
                oTransferReq.FromWarehouse = agendado.Almacen.ToString();
                oTransferReq.ToWarehouse = "CORTE";

                string modelo = "";
                string ccrCod = "";
                oRS.DoQuery(Comunes.Consultas.GetItemData(sol.MGS_CL_ARTC));
                if (oRS.RecordCount > 0)
                {
                    modelo = oRS.Fields.Item(0).Value.ToString();
                    ccrCod = oRS.Fields.Item(1).Value.ToString();
                }

                foreach (SolicitudDetalle item in sol.Detalle)
                {
                    oTransferReq.Lines.ItemCode = sol.MGS_CL_ARTC;
                    oTransferReq.Lines.UserFields.Fields.Item("U_MGS_CL_LARGO").Value = item.MGS_CL_MLAR.ToString();
                    oTransferReq.Lines.UserFields.Fields.Item("U_MGS_CL_ANCHO").Value = item.MGS_CL_ANCM.ToString();
                    oTransferReq.Lines.UserFields.Fields.Item("U_MGS_CL_CANBOB").Value = item.MGS_CL_MNBO.ToString();
                    var asdasd = double.Parse(item.MGS_CL_MCAN.ToString());
                    oTransferReq.Lines.Quantity = double.Parse(item.MGS_CL_MCAN.ToString());
                    oTransferReq.Lines.UserFields.Fields.Item("U_MGS_CL_MODELO").Value = oRS.Fields.Item(0).Value.ToString();
                    oTransferReq.Lines.DistributionRule = ccrCod;
                    oTransferReq.Lines.FromWarehouseCode = agendado.Almacen;
                    oTransferReq.Lines.WarehouseCode = "CORTE";
                    oTransferReq.Lines.BatchNumbers.BatchNumber = item.MGS_CL_LOTE;
                    oTransferReq.Lines.BatchNumbers.Quantity = double.Parse(item.MGS_CL_MCAN.ToString());

                    oTransferReq.Lines.Add();
                }

                iErrCod = oTransferReq.Add();
                if (iErrCod != 0)
                {

                    Globales.oCompany.GetLastError(out iErrCod, out sErrMsg);


                   // Globales.oApp.MessageBox(" Solicitud de traslado : " + sErrMsg);
                    rpta = "Solic. corte: " + agendado.DocEntry.ToString() + " - " + sErrMsg;

                }
                else
                {

                    oTransferReq.GetByKey(int.Parse(Globales.oCompany.GetNewObjectKey()));

                    var sfsdf = Globales.oCompany.GetNewObjectKey();

                    Comunes.FuncionesComunes.UpdateUDO(agendado.DocEntry, Globales.oCompany.GetNewObjectKey(), "U_MGS_CL_SOLTRA");

                    //Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se generó la solicitud de traslado",
                    //SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    rpta = "Solic. corte: " + agendado.DocEntry.ToString() + " - Se generó con éxito" ;

                    //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                }



            }
            catch (Exception ex)
            {
                //Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                rpta = "Solic. corte: " + agendado.DocEntry.ToString() + " - " + ex.Message.ToString();
            }

            return rpta;
        }
    }
}
