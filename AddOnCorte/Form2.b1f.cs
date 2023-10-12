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

                if (this.ComboBox0.Value.ToString() != "")
                {

                    if (this.ComboBox0.Value.ToString() == "A")
                    {
                        this.Button1.Item.Enabled = true;
                        this.Button2.Item.Enabled = false;
                    }
                    else
                    {
                        this.Button1.Item.Enabled = false;
                        this.Button2.Item.Enabled = true;
                    }


                    oGrid.DataTable.ExecuteQuery(Comunes.Consultas.GetSolicitudesAgenda(this.ComboBox0.Value));


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
                    oGrid.Columns.Item("AlmacenOrigen").TitleObject.Caption = "Almacén origen";
                    oGrid.Columns.Item("AlmacenOrigen").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                    ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("AlmacenOrigen")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                    //oGrid.Columns.Item("FechaCorte").Type = SAPbouiCOM.BoGridColumnType.gct_Date;

                    oGrid.Columns.Item("Equipo").TitleObject.Caption = "Equipo";
                    oGrid.Columns.Item("Equipo").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                    ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Equipo")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;

                    oGrid.Columns.Item("Seleccion").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    oGrid.Columns.Item("Seleccion").TitleObject.Caption = "Check";

                    if (this.ComboBox0.Value.ToString() == "A")
                    {
                        oGrid.Columns.Item("FechaCorte").Visible = true;
                        oGrid.Columns.Item("Equipo").Visible = true;
                        oGrid.Columns.Item("Serial").Visible = true;
                        oGrid.Columns.Item("AlmacenOrigen").Visible = false;
                        AjustarComboBox(2, "Item_0", 1);
                    }
                    else
                    {
                        oGrid.Columns.Item("FechaCorte").Visible = false;
                        oGrid.Columns.Item("Equipo").Visible = false;
                        oGrid.Columns.Item("Serial").Visible = false;
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

                    oForm.Freeze(true);

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





                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int ircontando = 0;
                for (int i = 0; i < oForm.DataSources.DataTables.Item("dt_gr").Rows.Count; i++)
                {

                    if (oForm.DataSources.DataTables.Item("dt_gr").GetValue("Seleccion", i).ToString().Equals("Y"))
                    {
                        ircontando++;
                        Globales.oApp.MessageBox("Se agendó");
                    }
                }

                if (ircontando == 0)
                    Globales.oApp.MessageBox("Debe marcar las solicitudes a agendar, no ha marcado ninguna");
                else
                    ListSolicitudes();

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }
    }
}
