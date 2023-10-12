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
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            
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
                    //oGrid.Columns.Item("FechaCorte").Type = SAPbouiCOM.BoGridColumnType.gct_Date;

                    

                    oGrid.Columns.Item("Seleccion").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    oGrid.Columns.Item("Seleccion").TitleObject.Caption = "Check";

                    if (this.ComboBox0.Value.ToString() == "A")
                    {
                        oGrid.Columns.Item("FechaCorte").Visible = true;
                        oGrid.Columns.Item("Equipo").Visible = true;
                        oGrid.Columns.Item("Serial").Visible = true;
                        oGrid.Columns.Item("AlmacenOrigen").Visible = false;
                    }
                    else
                    {
                        oGrid.Columns.Item("FechaCorte").Visible = false;
                        oGrid.Columns.Item("Equipo").Visible = false;
                        oGrid.Columns.Item("Serial").Visible = false;
                        oGrid.Columns.Item("AlmacenOrigen").Visible = true;
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

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
    }
}
