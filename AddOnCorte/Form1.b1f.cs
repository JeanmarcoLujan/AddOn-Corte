using AddOnCorte.Clases;
using AddOnCorte.Comunes;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace AddOnCorte
{
    [FormAttribute("AddOnCorte.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        SAPbouiCOM.Form oForm;
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_12").Specific));
            this.EditText4.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText4_LostFocusAfter);
            this.EditText4.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText4_ChooseFromListAfter);
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_14").Specific));
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_15").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_16").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_17").Specific));
            this.Matrix1.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix1_KeyDownAfter);
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_18").Specific));
            this.Matrix2.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.Matrix2_ComboSelectAfter);
            this.Matrix2.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix2_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_20").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_22").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_23").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_25").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_26").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("Item_27").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_30").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_31").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_32").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_33").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_34").Specific));
            this.EditText12.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText12_LostFocusAfter);
            this.EditText12.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText12_ChooseFromListAfter);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_35").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_36").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_37").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_38").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_10").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_19").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_28").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_39").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_40").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_41").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_42").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_43").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_44").Specific));
            this.Button4.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button4_PressedAfter);
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_45").Specific));
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_46").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_47").Specific));
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_48").Specific));
            this.StaticText20 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_49").Specific));
            this.StaticText21 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_50").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_51").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("Item_52").Specific));
            this.EditText17 = ((SAPbouiCOM.EditText)(this.GetItem("Item_53").Specific));
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("Item_54").Specific));
            this.EditText19 = ((SAPbouiCOM.EditText)(this.GetItem("Item_55").Specific));
            this.StaticText22 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_56").Specific));
            this.EditText21 = ((SAPbouiCOM.EditText)(this.GetItem("Item_58").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_59").Specific));
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("Item_57").Specific));
            this.Button5.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button5_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            Clases.Globales.oApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(this.m_SBO_Appl_MenuEvent);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        public void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            try
            {
                switch (pVal.MenuUID)
                {

                    case "1282":
                        //oForm.DataSources.UserDataSources.Item("UD_0").Value = "sdfsd";
                        var asd = "asdad";
                        this.Button3.Item.Enabled = false;
                        Comunes.Funciones.AutoResizeColumnsMatrix(oForm);
                        break;
                    case "1281":
                        var asdasd = "";
                        this.Button3.Item.Enabled = false;
                        break;
                    case "1290":
                    case "1289":
                    case "1288":
                    case "1291":
                        var ssss = "";
                        break;
                }
            }
            catch (Exception ex)
            {
                //ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
            }
        }

        private SAPbouiCOM.Folder Folder0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            Comunes.Funciones.AutoResizeColumnsMatrixInicial(oForm);
            CondiconesCFLsSN();
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.ComboBox ComboBox0;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
           // throw new System.NotImplementedException();

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.ComboBox oCombo = null;
            //throw new System.NotImplementedException();

            var sda = oForm.Mode;

            string obtenerInfo = "NONE";
            switch (oForm.Mode)
            {

                case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                    obtenerInfo = "ADD";
                    break;
                case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                    obtenerInfo = "UPDATE";
                    break;
                default:
                    obtenerInfo = "NONE";
                    //this.EditText9.Item.Enabled = false;
                    break;
            }

            if(obtenerInfo == "ADD")
            {
                
                List<string> oListErr = new List<string>();

                string cardCode = this.EditText3.Value;
                if(cardCode == "")
                    oListErr.Add("Debe seleccionar un cliente");

                string itemCode = this.EditText4.Value;
                if (itemCode == "")
                    oListErr.Add("Debe seleccionar un artículo");

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;
                if (oMatrix.RowCount == 0)
                {
                    oListErr.Add("Debe especificar informacion en la pestaña MASTER");
                }

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;
                if (oMatrix.RowCount == 0)
                {
                    oListErr.Add("Debe especificar informacion en la pestaña RESUMEN");
                }
                else
                {
                    //SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbComboBox").Specific;
                    //string valorSeleccionado = oComboBox.Selected.Value;
                    int contar = 0;
                    for (int j = 1; j <= oMatrix.RowCount - 1; j++)
                    {
                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(j).Specific;
                        if (oCombo.Selected == null)
                            contar++;
                    }
                    if (contar > 0)
                        oListErr.Add("En el RESUMEN, debe seleccionar los lotes");
                    
                    if( int.Parse(this.EditText15.Value) != (int.Parse(this.EditText17.Value) + int.Parse(this.EditText18.Value)) )
                        oListErr.Add("En el RESUMEN, debe distribuir el total del número de bobinas ");

                }

                if (oListErr.Count > 0)
                {
                    string msmValidate = "";
                    foreach (var msm in oListErr)
                    {
                        msmValidate = msmValidate + " > " + msm + "\n";
                    }

                    Globales.oApp.MessageBox("La siguiente informacion es obligatoria: \n" + msmValidate);
                }

                if (oListErr.Count > 0)
                    BubbleEvent = false;
                else
                    BubbleEvent = true;
            }
            else
            {
                BubbleEvent = true;
            }




        }

        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;


        private void CondiconesCFLsSN()
        {
            //SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Conditions oConds = null;
            SAPbouiCOM.Condition oCond = null;
            SAPbouiCOM.EditText oEditText111 = null;
            try
            {

                oForm.ChooseFromLists.Item("cfl_SN").SetConditions(null);

                oConds = new SAPbouiCOM.Conditions();
                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "C";



                oForm.ChooseFromLists.Item("cfl_SN").SetConditions(oConds);
                // oForm.ChooseFromLists.Item("cfl_SN").SetConditions(oConds);


                oEditText111 = (SAPbouiCOM.EditText)oForm.Items.Item("Item_34").Specific;
                oEditText111.DataBind.SetBound(true, "@MGS_CL_COCABE", "U_MGS_CL_CLIE");
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private SAPbouiCOM.EditText EditText12;

        private void EditText12_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                var currentForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource oDBDS = currentForm.DataSources.DBDataSources.Item("@MGS_CL_COCABE");

                this.EditText12.Value = chooseFromListEvent.SelectedObjects.GetValue(0, 0).ToString();
                oDBDS.SetValue("U_MGS_CL_CLID", 0, chooseFromListEvent.SelectedObjects.GetValue(1, 0).ToString());
            }
            catch
            {

            }

        }

        private void EditText4_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                var currentForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource oDBDS = currentForm.DataSources.DBDataSources.Item("@MGS_CL_COCABE");

                //this.EditText12.Value = chooseFromListEvent.SelectedObjects.GetValue(0, 0).ToString();
                oDBDS.SetValue("U_MGS_CL_ARTD", 0, chooseFromListEvent.SelectedObjects.GetValue(1, 0).ToString());
            }
            catch
            {

            }

        }

        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Items.Item("Item_36").Width = 190;
            oForm.Items.Item("Item_36").Height = 317;
            


        }

        private void EditText4_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                bool continuar = true;

                if (this.EditText4.Value == "")
                {
                    Globales.oApp.MessageBox("Debe seleccionar un artículo.");
                    continuar = false;
                }

                if (continuar == true && pVal.FormMode == 3)
                {
                    oItem = oForm.Items.Item("Item_36");

                    oForm.PaneLevel = 1;
                    oGrid = (SAPbouiCOM.Grid)oItem.Specific;

                    oGrid.DataTable = oForm.DataSources.DataTables.Item("dt_lt");



                    oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    

                    oGrid.DataTable.ExecuteQuery(Comunes.Consultas.GetAllLotesByItem(this.EditText4.Value.ToString()));

                    
                    oGrid.Columns.Item("Lote").Editable = false;

                    oGrid.Columns.Item("Seleccion").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    oGrid.Columns.Item("Seleccion").TitleObject.Caption = "Check";

                    oGrid.Columns.Item("U_MGS_CL_ANCHO").TitleObject.Caption = "Ancho";
                    oGrid.Columns.Item("U_MGS_CL_ANCHO").Editable = false;

                    oGrid.Columns.Item("U_MGS_CL_LARGO").TitleObject.Caption = "Largo";
                    oGrid.Columns.Item("U_MGS_CL_LARGO").Editable = false;

                    oGrid.Columns.Item("U_MGS_CL_BOBINA").TitleObject.Caption = "Nro bobinas";
                    oGrid.Columns.Item("U_MGS_CL_BOBINA").Editable = false;

                    oGrid.Columns.Item("U_MGS_CL_CANT").TitleObject.Caption = "Cantidad";
                    oGrid.Columns.Item("U_MGS_CL_CANT").Editable = false;


                    oGrid.AutoResizeColumns();

                    GetPrecioArticulo();

                }


            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.ProgressBar oPB = null;


            try
            {

                int cont = 0;
                oPB = Clases.Globales.oApp.StatusBar.CreateProgressBar("Obteniendo la data", 27, true);
                oPB.Text = "Obteniendo la data";
                oPB.Value = 10;


                
                oForm.Freeze(true);


                double total_ancho = 0;
                double total_largo = 0;
                double total_cantidad = 0;
                for (int i = 0; i < oForm.DataSources.DataTables.Item("dt_lt").Rows.Count; i++)
                {
                    if (oForm.DataSources.DataTables.Item("dt_lt").GetValue("Seleccion", i).ToString().Equals("Y"))
                    {
                      
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;
                        oMatrix.AddRow();
                        oMatrix.FlushToDataSource();
                        //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                        //else bAddNewRow = false;
                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_COMAST");

                        var asdasd = oForm.DataSources.DataTables.Item("dt_lt").GetValue("Lote", i).ToString();

                        var sdsd = oDBDataSource.Size - 1;
                        oDBDataSource.Offset = oDBDataSource.Size - 1;
                        oDBDataSource.SetValue("U_MGS_CL_LOTE", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("Lote", i).ToString());
                        oDBDataSource.SetValue("U_MGS_CL_ANCM", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_ANCHO", i).ToString());
                        oDBDataSource.SetValue("U_MGS_CL_MLAR", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_LARGO", i).ToString());
                        oDBDataSource.SetValue("U_MGS_CL_MNBO", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_BOBINA", i).ToString());
                        oDBDataSource.SetValue("U_MGS_CL_MCAN", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_CANT", i).ToString());

                        total_ancho = total_ancho + double.Parse( oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_ANCHO", i).ToString());
                        total_largo = total_largo + double.Parse( oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_LARGO", i).ToString());
                        total_cantidad = total_cantidad + double.Parse( oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_CANT", i).ToString());


                        oMatrix.LoadFromDataSource();
                        oMatrix.AutoResizeColumns();

                        cont++;
                        if (oPB.Value < 25)
                            oPB.Value = oPB.Value + 10;
                    }
                }

                this.EditText2.Value = (total_ancho/cont).ToString();
                this.EditText13.Value = total_largo.ToString();
                this.EditText14.Value = total_cantidad.ToString();

                //oDBDataSource.SetValue("U_MGS_CL_LOTE", cont - 1, "Total");

                //oMatrix.AddRow();
                //oMatrix.FlushToDataSource();



                System.Threading.Thread.Sleep(1000);
                oPB.Stop();
                Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);




            }
            catch (Exception ex)
            {
                //ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                if (oPB != null)
                {
                    oPB.Stop();
                    Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);
                }
            }
            finally
            {
                Comunes.FuncionesComunes.LiberarObjetoGenerico(null);

                oForm.Freeze(false);
            }

        }

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            SAPbouiCOM.Matrix oMatrix = null;
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;
                oMatrix.Clear();

                this.EditText2.Value = "0";
                this.EditText13.Value = "0";
                this.EditText14.Value = "0";

            }
            catch
            {

            }

        }

        private SAPbouiCOM.Button Button3;

        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            
            List<string[,]> oListErr = null;


            bool continuar = true;

            switch (oForm.Mode)
            {

                case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                    continuar = false;
                    break;
                default:
                    continuar = true;
                    //this.EditText9.Item.Enabled = false;
                    break;
            }

            if (continuar)
            {
      
                //if (this.EditText21.Value == "")
                if (true)
                {
                    if (Globales.oApp.MessageBox("¿Esta Ud. seguro de generar la oferta de venta?, revise toda la información", 1, "Continuar", "Cancelar", "") == 1)
                    {


                        GenerateOferta(out oListErr);
                    }
                }
                else
                {
                    Globales.oApp.MessageBox("La oferta de ventas para este registro, ya se generó");
                }
                

            }

        }

        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.StaticText StaticText15;

        private void Matrix1_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;

                double ancho = this.EditText2.Value == "" ? 0: double.Parse(this.EditText2.Value);

                for (int k = 2; k < 12; k++)
                {
                    double c1_sub = 0;
                    double c1_total = 0;
                    double c1_largo = 0;

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(2).Specific; //Cast the Cell of 
                    c1_total = double.Parse(oEditText.Value.ToString());

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(1).Specific; //Cast the Cell of 
                    c1_largo = double.Parse(oEditText.Value.ToString());


                    for (int i = 3; i <= 17; i++)
                    {
                        
                        //oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_CORRID");
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(i).Specific; //Cast the Cell of 
                        c1_sub = c1_sub + double.Parse(oEditText.Value.ToString());
                        oMatrix.FlushToDataSource();

                    }

                    if(c1_sub != 0)
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(18).Specific; //Cast the Cell of 
                        oEditText.Value = c1_sub.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(19).Specific; //Cast the Cell of 
                        oEditText.Value = (c1_sub + c1_total).ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(20).Specific; //Cast the Cell of 
                        oEditText.Value = (ancho - (c1_sub + c1_total)).ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(k).Cells.Item(21).Specific; //Cast the Cell of 
                        oEditText.Value = (Math.Round(c1_sub * c1_largo * 2.54 * 2.54 * 12 / 10000, 2)).ToString();
                    }

                    
                }
                



            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private SAPbouiCOM.Button Button4;

        private void Button4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                bool continuar = false;
                if (this.EditText3.Value.ToString() == "")
                {
                    continuar = false;
                    Globales.oApp.MessageBox("Seleccione el cliente, para poder generar el resumen adecuadamente");
                }else if (this.EditText4.Value.ToString() == "")
                {
                    continuar = false;
                    Globales.oApp.MessageBox("Seleccione el artículo, para poder generar el resumen adecuadamente");
                }
                else
                {
                    GetPrecioArticulo();
                    continuar = true;
                }
                    



                if (pVal.FormMode == 3 && continuar == true)
                {
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;

                    int columnCount = oMatrix.Columns.Count;
                    List<ColumnValueCount> valueCounts = new List<ColumnValueCount>();

                    // Recorrer las columnas de la matriz
                    for (int colIndex = 2; colIndex <= columnCount - 1; colIndex++)
                    {
                        string columnName = oMatrix.Columns.Item(colIndex).TitleObject.Caption;

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colIndex).Cells.Item(1).Specific;
                        string largo = oEditText.Value.ToString();

                        // Recorrer las filas de la columna y contar los valores repetidos
                        Dictionary<string, int> columnValueCount = new Dictionary<string, int>();
                        for (int rowIndex = 3; rowIndex <= oMatrix.RowCount - 4; rowIndex++)
                        {
                            //string cellValue = oMatrix.Columns.Item(colIndex).Cells.Item(rowIndex).Specific.Value;

                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colIndex).Cells.Item(rowIndex).Specific;
                            string cellValue = oEditText.Value.ToString();

                            if (cellValue != "0")
                            {
                                if (columnValueCount.ContainsKey(cellValue))
                                {
                                    columnValueCount[cellValue]++;
                                }
                                else
                                {
                                    columnValueCount.Add(cellValue, 1);
                                }
                            }
                        }

                        // Agregar los resultados a la lista de objetos
                        foreach (var kvp in columnValueCount)
                        {
                            valueCounts.Add(new ColumnValueCount
                            {
                                ColumnName = columnName,
                                CellValue = kvp.Key,
                                Count = kvp.Value,
                                LargoValue = double.Parse(largo),
                                ColumnId = colIndex - 1
                            });
                        }
                    }


                    var adasd = valueCounts;

                    var result = valueCounts
                    .GroupBy(vc => new { vc.CellValue, vc.LargoValue })
                    .Select(group => new ColumnValueCount
                    {
                        CellValue = group.Key.CellValue,
                        LargoValue = group.Key.LargoValue,
                        Count = group.Sum(vc => vc.Count),
                    });

                    var sdsdf = result.ToList();

                    int cont = 1;
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;
                    oMatrix.Clear();
                    double sum_count = 0;
                    double sum_rcan = 0;
                    foreach (var item in result)
                    {
                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_CORESU");
                        var sdsd = oDBDataSource.Size - 1;
                        oDBDataSource.Offset = oDBDataSource.Size - 1;
                        oDBDataSource.SetValue("U_MGS_CL_RANC", oDBDataSource.Size - 1, item.CellValue);
                        oDBDataSource.SetValue("U_MGS_CL_RLAR", oDBDataSource.Size - 1, item.LargoValue.ToString());
                        sum_count = sum_count + item.Count;
                        oDBDataSource.SetValue("U_MGS_CL_RBOB", oDBDataSource.Size - 1, item.Count.ToString());
                        sum_rcan = sum_rcan + Math.Round(double.Parse(item.CellValue) * item.LargoValue * item.Count * 2.54 * 2.54 * 12 / 10000, 2);
                        oDBDataSource.SetValue("U_MGS_CL_RCAN", oDBDataSource.Size - 1, (Math.Round(double.Parse(item.CellValue) * item.LargoValue * item.Count * 2.54 * 2.54 * 12 / 10000, 2)).ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RBOA", oDBDataSource.Size - 1, "0");
                        oDBDataSource.SetValue("U_MGS_CL_RBOV", oDBDataSource.Size - 1, "0");
                        oDBDataSource.SetValue("U_MGS_CL_RCAV", oDBDataSource.Size - 1, "0");

                        //Math.Round(numeroOriginal, 2)
                        oMatrix.AddRow();
                        oMatrix.FlushToDataSource();

                        oMatrix.LoadFromDataSource();
                        cont++;

                    }

                    this.EditText15.Value = sum_count.ToString();
                    this.EditText16.Value = sum_rcan.ToString();
                    this.EditText17.Value = "0";
                    this.EditText18.Value = "0";
                    this.EditText19.Value = "0";


                    List<string> listaLotes = new List<string>();
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;

                    for (int m = 1; m <= oMatrix.RowCount; m++)
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(1).Cells.Item(m).Specific;
                        string cellValue = oEditText.Value.ToString();

                        listaLotes.Add(oEditText.Value.ToString());

                    }


                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;

                    for (int j = 1; j <= oMatrix.RowCount - 1; j++)
                    {
                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(j).Specific;
                        foreach (var item in listaLotes)
                        {
                            oCombo.ValidValues.Add(item, item);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
               // Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private void Matrix2_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;

                if(pVal.ColUID=="Col_5" || pVal.ColUID == "Col_6")
                {
                    

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(6).Cells.Item(pVal.Row).Specific; 
                    int columnValue6 = FuncionesComunes.ValidateNumberInt( oEditText.Value.ToString());

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(7).Cells.Item(pVal.Row).Specific;
                    int columnValue7 = FuncionesComunes.ValidateNumberInt(oEditText.Value.ToString());

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(4).Cells.Item(pVal.Row).Specific;
                    int columnValue4 = int.Parse(oEditText.Value.ToString());

                    

                    if ((columnValue6 + columnValue7) > columnValue4)
                        Globales.oApp.MessageBox("No debe superar el número de bobina disponibles");
                    else
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific;
                        double columnValue5 = double.Parse(oEditText.Value.ToString());

                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(pVal.Row).Specific;
                        var columnValue1 = oCombo.Selected;

                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_CORESU");
                        string calculo = (Math.Round((columnValue5 / columnValue4) * columnValue7, 2)).ToString();
                        if (columnValue1 != null)
                            oDBDataSource.SetValue("U_MGS_CL_LOTE", pVal.Row - 1, columnValue1.Value.ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RBOA", pVal.Row - 1, columnValue6.ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RBOV", pVal.Row - 1, columnValue7.ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RCAV", pVal.Row - 1, calculo);
                        oMatrix.LoadFromDataSource();

                        
                    }
                    double sum_ba = 0;
                    double sum_bv = 0;
                    double sum_cv = 0;
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(6).Cells.Item(i).Specific;
                        sum_ba = sum_ba + double.Parse(oEditText.Value.ToString());
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(7).Cells.Item(i).Specific;
                        sum_bv = sum_bv + double.Parse(oEditText.Value.ToString());
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(8).Cells.Item(i).Specific;
                        sum_cv = sum_cv + double.Parse(oEditText.Value.ToString());
                    }
                    this.EditText17.Value = sum_ba.ToString();
                    this.EditText18.Value = sum_bv.ToString();
                    this.EditText19.Value = Math.Round(sum_cv,2).ToString();
                    this.EditText9.Value = (Math.Round(sum_cv, 2)* double.Parse(this.EditText8.Value.ToString()) ).ToString();
                    this.EditText10.Value = (double.Parse(this.EditText9.Value.ToString()) * 1.16).ToString();
                    this.EditText11.Value = ( 100*double.Parse(this.EditText19.Value.ToString())/ double.Parse(this.EditText16.Value.ToString()) ).ToString();

                }

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private void Matrix2_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;

                if (pVal.ColUID == "Col_0")
                {


                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(6).Cells.Item(pVal.Row).Specific;
                    int columnValue6 = FuncionesComunes.ValidateNumberInt(oEditText.Value.ToString());

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(7).Cells.Item(pVal.Row).Specific;
                    int columnValue7 = FuncionesComunes.ValidateNumberInt(oEditText.Value.ToString());

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(4).Cells.Item(pVal.Row).Specific;
                    int columnValue4 = int.Parse(oEditText.Value.ToString());



                    if ((columnValue6 + columnValue7) > columnValue4)
                        Globales.oApp.MessageBox("No debe superar el Número de bobina disponibles");
                    else
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(pVal.Row).Specific;
                        double columnValue5 = double.Parse(oEditText.Value.ToString());

                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(pVal.Row).Specific;
                        var columnValue1 = oCombo.Selected;

                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_CORESU");
                        string calculo = ( Math.Round((columnValue5 / columnValue4) * columnValue7,2)).ToString();
                        if (columnValue1 != null)
                            oDBDataSource.SetValue("U_MGS_CL_LOTE", pVal.Row - 1, columnValue1.Value.ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RBOA", pVal.Row - 1, columnValue6.ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RBOV", pVal.Row - 1, columnValue7.ToString());
                        oDBDataSource.SetValue("U_MGS_CL_RCAV", pVal.Row - 1, calculo);
                        oMatrix.LoadFromDataSource();

                    }

                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.StaticText StaticText17;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.StaticText StaticText19;
        private SAPbouiCOM.StaticText StaticText20;
        private SAPbouiCOM.StaticText StaticText21;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.EditText EditText17;
        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.EditText EditText19;


        private void GenerateOferta(out List<string[,]> oListErr)
        {
            oListErr = null;
            SAPbobsCOM.Documents oSalesOpportunity = null;
            //SalesOpportunities
            SAPbouiCOM.Matrix oMatrix = null;
            int iErrCod;
            string sErrMsg = "";
            SAPbouiCOM.ProgressBar oPB = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbobsCOM.Recordset oRS = null;

            try
            {
                oPB = Clases.Globales.oApp.StatusBar.CreateProgressBar("Generando la oferta de venta", 27, true);
                oPB.Text = "GGenerando la oferta de venta";
                oPB.Value = 10;

                oSalesOpportunity = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oSalesOpportunity.CardCode = this.EditText3.Value;
                oSalesOpportunity.TaxDate = DateTime.Now;
                oSalesOpportunity.DocDate = DateTime.Now;
                oSalesOpportunity.DocDueDate = DateTime.Now.AddDays(30);

                oSalesOpportunity.UserFields.Fields.Item("U_MGS_CL_SOLCOR").Value = this.EditText0.Value.ToString();
                oSalesOpportunity.UserFields.Fields.Item("U_MGS_CL_EFCO").Value = this.EditText11.Value.ToString();
                oSalesOpportunity.UserFields.Fields.Item("U_MGS_CL_TIPORD").Value = "2";



                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;

                int salesPersonCode = 1;
                var asdas = oMatrix.RowCount;
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    //oSalesOpportunity.Lines.SetCurrentLine(i-1);
                    oSalesOpportunity.Lines.ItemCode = this.EditText4.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(3).Cells.Item(i).Specific;
                    oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_LARGO").Value = oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(2).Cells.Item(i).Specific;
                    oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_ANCHO").Value = oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(7).Cells.Item(i).Specific;
                    oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_CANBOB").Value = oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(8).Cells.Item(i).Specific;
                    oSalesOpportunity.Lines.Quantity = double.Parse(oEditText.Value.ToString());

                    oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                    var sdsd = oCombo.Selected.Value.ToString();
                    oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_LOTE").Value  = oCombo.Selected.Value.ToString();

                    oRS.DoQuery(Comunes.Consultas.GetItemData(this.EditText4.Value.ToString()));
                    if (oRS.RecordCount > 0)
                    {
                        oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_MODELO").Value = oRS.Fields.Item(0).Value.ToString();
                        oSalesOpportunity.Lines.CostingCode = oRS.Fields.Item(1).Value.ToString();
                    }

                    oSalesOpportunity.Lines.Add();


                    if (oPB.Value < 25)
                        oPB.Value = oPB.Value + 15;


                }

                //oSalesOpportunity.SalesPersonCode = salesPersonCode==-1?1: salesPersonCode;

                var ss1 = oSalesOpportunity.GetAsXML();

                iErrCod = oSalesOpportunity.Add();
                if (iErrCod != 0)
                {

                    Globales.oCompany.GetLastError(out iErrCod, out sErrMsg);
                    

                    Globales.oApp.MessageBox("Oferta de venta: " + sErrMsg);

                }
                else
                {

                    oSalesOpportunity.GetByKey(int.Parse(Globales.oCompany.GetNewObjectKey()));

                    var sfsdf = Globales.oCompany.GetNewObjectKey();

                    Comunes.FuncionesComunes.UpdateUDO(this.EditText0.Value.ToString(), Globales.oCompany.GetNewObjectKey(), "U_MGS_CL_OFEV");

                    Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se generó la oferta de venta ",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                }

                System.Threading.Thread.Sleep(1000);
                oPB.Stop();
                Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);


                //oPurchaseReturns.

            }
            catch (Exception ex)
            {
                //Comunes.Funciones.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                if (oPB != null)
                {
                    oPB.Stop();
                    Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);
                }
                //Comunes.Funciones.RemoveRegisterUDO(docEntryUDO);
            }
        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                this.Button3.Item.Enabled = true;

                List<string> listaLotes = new List<string>();
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;

                for (int m = 1; m <= oMatrix.RowCount; m++)
                {
                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(1).Cells.Item(m).Specific;
                    string cellValue = oEditText.Value.ToString();

                    listaLotes.Add(oEditText.Value.ToString());

                }


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_18").Specific;

                
                for (int j = 1; j <= oMatrix.RowCount - 1; j++)
                {
                    oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(j).Specific;
                    int validValueCount = oCombo.ValidValues.Count;
                    for (int i = validValueCount - 1; i >= 0; i--)
                    {
                        oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    foreach (var item in listaLotes)
                    {
                        oCombo.ValidValues.Add(item, item);
                    }

                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }
            //throw new System.NotImplementedException();
            



        }

        private SAPbouiCOM.StaticText StaticText22;
        private SAPbouiCOM.EditText EditText21;
        private SAPbouiCOM.LinkedButton LinkedButton1;

        private void GetPrecioArticulo()
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Comunes.Consultas.GetPrecioArticulo(this.EditText4.Value.ToString(), this.EditText3.Value.ToString()));

                if (oRS.RecordCount > 0)
                {
                    this.EditText8.Value = oRS.Fields.Item(0).Value.ToString();
                }
                else
                    this.EditText8.Value = "0";
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }
        }

        

        private void EditText12_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.FormMode == 3)
                    GetPrecioArticulo();
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private SAPbouiCOM.Button Button5;

        private void Button5_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            Form3 activeForm = new Form3();
            activeForm.Show();

        }
    }
}