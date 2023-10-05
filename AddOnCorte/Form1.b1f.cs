using AddOnCorte.Clases;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
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
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_18").Specific));
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
            this.EditText12.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText12_ChooseFromListAfter);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_35").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_36").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_37").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_38").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_10").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            
            Globales.oApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(this.m_SBO_Appl_MenuEvent);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

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
                        Comunes.Funciones.AutoResizeColumnsMatrix(oForm);
                        break;
                    case "1281":
                        var asdasd = "";
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
            //throw new System.NotImplementedException();

            var sda = oForm.Mode;


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




                if (continuar)
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

                
                
                for (int i = 0; i < oForm.DataSources.DataTables.Item("dt_lt").Rows.Count; i++)
                {
                    if (oForm.DataSources.DataTables.Item("dt_lt").GetValue("Seleccion", i).ToString().Equals("Y"))
                    {
                        cont++;
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


                        oMatrix.LoadFromDataSource();
                        oMatrix.AutoResizeColumns();

                        cont++;
                        if (oPB.Value < 25)
                            oPB.Value = oPB.Value + 10;
                    }
                }

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

            }
            catch
            {

            }

        }

        private SAPbouiCOM.Button Button3;
    }
}