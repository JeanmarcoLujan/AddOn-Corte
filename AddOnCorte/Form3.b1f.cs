using AddOnCorte.Clases;
using AddOnCorte.Comunes;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOnCorte
{
    [FormAttribute("AddOnCorte.Form3", "Form3.b1f")]
    class Form3 : UserFormBase
    {
        SAPbouiCOM.Form oForm;
        public Form3()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_7").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_13").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_15").Specific));
            this.Folder1.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder1_PressedAfter);
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_16").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_17").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_3").Specific));
            this.Matrix1.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix1_KeyDownAfter);
            this.Matrix1.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.Matrix1_ComboSelectAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_14").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_20").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_22").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_25").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_26").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_27").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("Item_28").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_29").Specific));
            this.CheckBox0.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox0_PressedAfter);
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_30").Specific));
            this.CheckBox1.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox1_PressedAfter);
            this.CheckBox2 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_31").Specific));
            this.CheckBox2.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox2_PressedAfter);
            this.CheckBox3 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_32").Specific));
            this.CheckBox3.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox3_PressedAfter);
            this.CheckBox4 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_33").Specific));
            this.CheckBox4.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox4_PressedAfter);
            this.CheckBox5 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_34").Specific));
            this.CheckBox5.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox5_PressedAfter);
            this.CheckBox6 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_35").Specific));
            this.CheckBox6.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox6_PressedAfter);
            this.CheckBox7 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_36").Specific));
            this.CheckBox7.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox7_PressedAfter);
            this.CheckBox8 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_37").Specific));
            this.CheckBox8.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox8_PressedAfter);
            this.CheckBox9 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_38").Specific));
            this.CheckBox9.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox9_PressedAfter);
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_39").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_40").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_41").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_42").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_43").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_44").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_45").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_46").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_47").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_48").Specific));
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_50").Specific));
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_51").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_52").Specific));
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_53").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_54").Specific));
            this.StaticText20 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_55").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("Item_56").Specific));
            this.StaticText21 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_57").Specific));
            this.EditText17 = ((SAPbouiCOM.EditText)(this.GetItem("Item_58").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_59").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("Item_60").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_61").Specific));
            this.StaticText22 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_62").Specific));
            this.EditText19 = ((SAPbouiCOM.EditText)(this.GetItem("Item_63").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_64").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_65").Specific));
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_66").Specific));
            this.StaticText23 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_67").Specific));
            this.EditText20 = ((SAPbouiCOM.EditText)(this.GetItem("Item_68").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_69").Specific));
            this.StaticText25 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_71").Specific));
            this.StaticText26 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_72").Specific));
            this.StaticText27 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_73").Specific));
            this.EditText21 = ((SAPbouiCOM.EditText)(this.GetItem("Item_74").Specific));
            this.EditText22 = ((SAPbouiCOM.EditText)(this.GetItem("Item_75").Specific));
            this.CheckBox10 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_76").Specific));
            this.CheckBox10.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox10_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            Funciones.CbxSolCorte(ref oForm, Globales.oApp, this.ComboBox0);
            Clases.Globales.oApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(this.m_SBO_Appl_MenuEvent);

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
                        this.Button2.Item.Enabled = false;
                        this.Button3.Item.Enabled = false;
                        
                        break;
                    case "1281":
                        this.Button2.Item.Enabled = false;
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


       


        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.ComboBox ComboBox0;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.CheckBox oCheckBox = null;
            try
            {
                string[] posiblesFormatos = {
                    "MM/dd/yyyy hh:mm:ss tt",
                    "MM/d/yyyy hh:mm:ss tt",
                    "dd/MM/yyyy hh:mm:ss tt",
                    "d/MM/yyyy hh:mm:ss tt",
                    "d/M/yyyy hh:mm:ss tt",
                    "yyyy-MM-dd",
                    "dd/MM/yyyy",
                    "yyyyMMdd",
                    "yyyy/MM/dd",
                    "MM/dd/yyyy",
                    "dd-MMM-yyyy"
                };

                if (pVal.FormMode == 3)
                {
                    var asd = this.ComboBox0.Selected.Value;
                    Solicitud sol = Comunes.FuncionesComunes.GetUDOSolicitudAgendada(asd);

                    if(sol != null)
                    {
                        this.EditText2.Value = sol.MGS_CL_CLIE;
                        this.EditText3.Value = sol.MGS_CL_CLID;
                        this.EditText4.Value = sol.MGS_CL_ARTC;
                        this.EditText5.Value = sol.MGS_CL_ARTD;
                        this.EditText10.Value = DateTime.Now.ToString("yyyyMMdd");
                        this.EditText11.Value = sol.MGS_CL_PRAR;
                        this.EditText12.Value = sol.MGS_CL_OFVE;
                        this.EditText13.Value = sol.MGS_CL_OFVI;
                        this.EditText14.Value = sol.MGS_CL_EFCO;
                        this.EditText15.Value = sol.MGS_CL_MTANC;
                        this.EditText16.Value = sol.MGS_CL_MTLAR;
                        this.EditText17.Value = sol.MGS_CL_MTCANT;

                        

                        //oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("Item_29").Specific;
                        //oCheckBox.Checked = true;

                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_51").Specific;
                        oMatrix.Clear();

                        foreach (SolicitudDetalle item in sol.Detalle)
                        {
                            
                            oMatrix.AddRow();
                            oMatrix.FlushToDataSource();
                            //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                            //else bAddNewRow = false;
                            oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_RCOMAST");


                            var sdsd = oDBDataSource.Size - 1;
                            oDBDataSource.Offset = oDBDataSource.Size - 1;
                            oDBDataSource.SetValue("U_MGS_CL_LOTE", oDBDataSource.Size - 1, item.MGS_CL_LOTE );
                            oDBDataSource.SetValue("U_MGS_CL_ANCM", oDBDataSource.Size - 1, item.MGS_CL_ANCM);
                            oDBDataSource.SetValue("U_MGS_CL_MLAR", oDBDataSource.Size - 1, item.MGS_CL_MLAR );
                            oDBDataSource.SetValue("U_MGS_CL_MNBO", oDBDataSource.Size - 1, item.MGS_CL_MNBO);
                            oDBDataSource.SetValue("U_MGS_CL_MCAN", oDBDataSource.Size - 1, item.MGS_CL_MCAN);

                            DateTime fecha;
                            if (DateTime.TryParseExact(item.MGS_CL_FECADM, posiblesFormatos, null, System.Globalization.DateTimeStyles.None, out fecha))
                            {
                                oDBDataSource.SetValue("U_MGS_CL_FECADM", oDBDataSource.Size - 1, fecha.ToString("yyyyMMdd"));
                            }

                            //var ssss = item.MGS_CL_FECADM;
                            ////DateTime fecha = DateTime.ParseExact(item.MGS_CL_FECADM, "dd/MM/yyyy", null);
                            //DateTime fecha = DateTime.ParseExact(item.MGS_CL_FECADM, "MM/dd/yyyy hh:mm:ss tt", null);
                            //string fechaFormateada = fecha.ToString("yyyyMMdd");

                            
                            oDBDataSource.SetValue("U_MGS_CL_FIFO", oDBDataSource.Size - 1, item.MGS_CL_FIFO);

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }


                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
                        oMatrix.Clear();
                        foreach (CorridasDetalle item in sol.DetalleCorridas)
                        {
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
                            oMatrix.AddRow();
                            oMatrix.FlushToDataSource();
                            //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                            //else bAddNewRow = false;
                            oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_RCORRID");


                            var sdsd = oDBDataSource.Size - 1;
                            oDBDataSource.Offset = oDBDataSource.Size - 1;
                            oDBDataSource.SetValue("U_MGS_CL_TITU", oDBDataSource.Size - 1, item.MGS_CL_TITU);
                            oDBDataSource.SetValue("U_MGS_CL_C1", oDBDataSource.Size - 1, item.MGS_CL_C1);
                            oDBDataSource.SetValue("U_MGS_CL_C2", oDBDataSource.Size - 1, item.MGS_CL_C2);
                            oDBDataSource.SetValue("U_MGS_CL_C3", oDBDataSource.Size - 1, item.MGS_CL_C3);
                            oDBDataSource.SetValue("U_MGS_CL_C4", oDBDataSource.Size - 1, item.MGS_CL_C4);
                            oDBDataSource.SetValue("U_MGS_CL_C5", oDBDataSource.Size - 1, item.MGS_CL_C5);
                            oDBDataSource.SetValue("U_MGS_CL_C6", oDBDataSource.Size - 1, item.MGS_CL_C6);
                            oDBDataSource.SetValue("U_MGS_CL_C7", oDBDataSource.Size - 1, item.MGS_CL_C7);
                            oDBDataSource.SetValue("U_MGS_CL_C8", oDBDataSource.Size - 1, item.MGS_CL_C8);
                            oDBDataSource.SetValue("U_MGS_CL_C9", oDBDataSource.Size - 1, item.MGS_CL_C9);
                            oDBDataSource.SetValue("U_MGS_CL_C10", oDBDataSource.Size - 1, item.MGS_CL_C10);
                            oDBDataSource.SetValue("U_MGS_CL_C11", oDBDataSource.Size - 1,"0");

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }

                        Form_ResizeAfter(pVal);

                        //for (int colIndex = 1; colIndex <= oMatrix.Columns.Count - 1; colIndex++)
                        //{
                        //    oMatrix.CommonSetting.SetCellEditable(18, colIndex, false);
                        //    oMatrix.CommonSetting.SetCellEditable(19, colIndex, false);
                        //    oMatrix.CommonSetting.SetCellEditable(20, colIndex, false);
                        //    oMatrix.CommonSetting.SetCellEditable(21, colIndex, false);
                        //}

                        //for (int i = 0; i < sol.Detalle.Count; i++)
                        //{
                        //    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_16").Specific;
                        //    oMatrix.AddRow();
                        //    oMatrix.FlushToDataSource();
                        //    //if (!bItsTheSameLine) oDBDataSource.InsertRecord(index);
                        //    //else bAddNewRow = false;
                        //    oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_COMAST");

                        //    var asdasd = oForm.DataSources.DataTables.Item("dt_lt").GetValue("Lote", i).ToString();

                        //    var sdsd = oDBDataSource.Size - 1;
                        //    oDBDataSource.Offset = oDBDataSource.Size - 1;
                        //    oDBDataSource.SetValue("U_MGS_CL_LOTE", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("Lote", i).ToString());
                        //    oDBDataSource.SetValue("U_MGS_CL_ANCM", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_ANCHO", i).ToString());
                        //    oDBDataSource.SetValue("U_MGS_CL_MLAR", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_LARGO", i).ToString());
                        //    oDBDataSource.SetValue("U_MGS_CL_MNBO", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_BOBINA", i).ToString());
                        //    oDBDataSource.SetValue("U_MGS_CL_MCAN", oDBDataSource.Size - 1, oForm.DataSources.DataTables.Item("dt_lt").GetValue("U_MGS_CL_CANT", i).ToString());

                        //    oMatrix.LoadFromDataSource();
                        //    oMatrix.AutoResizeColumns();
                        //}


                    }



                }
                
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.CheckBox CheckBox2;
        private SAPbouiCOM.CheckBox CheckBox3;
        private SAPbouiCOM.CheckBox CheckBox4;
        private SAPbouiCOM.CheckBox CheckBox5;
        private SAPbouiCOM.CheckBox CheckBox6;
        private SAPbouiCOM.CheckBox CheckBox7;
        private SAPbouiCOM.CheckBox CheckBox8;
        private SAPbouiCOM.CheckBox CheckBox9;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.StaticText StaticText17;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.StaticText StaticText19;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.StaticText StaticText20;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.StaticText StaticText21;
        private SAPbouiCOM.EditText EditText17;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                bool continuar = false;

                if(this.EditText18.Value != "")
                {
                    continuar = false;
                    Globales.oApp.MessageBox("No es posible generar el resumen, porque ya se genero la salida/entrada");
                }
                else if (this.EditText3.Value.ToString() == "")
                {
                    continuar = false;
                    Globales.oApp.MessageBox("Seleccione el cliente, para poder generar el resumen adecuadamente");
                }
                else if (this.EditText4.Value.ToString() == "")
                {
                    continuar = false;
                    Globales.oApp.MessageBox("Seleccione el artículo, para poder generar el resumen adecuadamente");
                }
                else
                {
                    //GetPrecioArticulo();
                    continuar = true;
                }




                if (pVal.FormMode == 3 && continuar == true)
                {
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;

                    int columnCount = oMatrix.Columns.Count;

                    List<ColumnValueCount> valueCounts = new List<ColumnValueCount>();

                    // Recorrer las columnas de la matriz
                    int contar_marcado = 0;
                    for (int colIndex = 2; colIndex <= columnCount - 1; colIndex++)
                    {
                        bool validar = false;
                        switch (colIndex)
                        {
                            case 2:
                                validar = this.CheckBox0.Checked;
                                break;
                            case 3:
                                validar = this.CheckBox1.Checked;
                                break;
                            case 4:
                                validar = this.CheckBox2.Checked;
                                break;
                            case 5:
                                validar = this.CheckBox3.Checked;
                                break;
                            case 6:
                                validar = this.CheckBox4.Checked;
                                break;
                            case 7:
                                validar = this.CheckBox5.Checked;
                                break;
                            case 8:
                                validar = this.CheckBox6.Checked;
                                break;
                            case 9:
                                validar = this.CheckBox7.Checked;
                                break;
                            case 10:
                                validar = this.CheckBox8.Checked;
                                break;
                            case 11:
                                validar = this.CheckBox9.Checked;
                                break;
                        }

                        
                        if (validar)
                        {
                            contar_marcado++;
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

                        
                    }

                    if(contar_marcado == 0)
                    {
                        Globales.oApp.MessageBox("Debe marcar las corridas a procesar en la pestaña \"Corridas\" ");
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
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;
                    oMatrix.Clear();
                    double sum_count = 0;
                    double sum_rcan = 0;
                    foreach (var item in result)
                    {
                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_RCORESU");
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

                    

                    this.EditText1.Value = sum_count.ToString();
                    this.EditText6.Value = sum_rcan.ToString();
                    this.EditText7.Value = "0";
                    this.EditText8.Value = "0";
                    this.EditText9.Value = "0";


                    List<string> listaLotes = new List<string>();
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_51").Specific;

                    for (int m = 1; m <= oMatrix.RowCount; m++)
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(1).Cells.Item(m).Specific;
                        string cellValue = oEditText.Value.ToString();

                        listaLotes.Add(oEditText.Value.ToString());

                    }


                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;

                    for (int j = 1; j <= oMatrix.RowCount; j++)
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

        private void Folder1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            try
            {
                if (pVal.FormMode == 3)
                {
                    //this.CheckBox0.Checked = true;
                    //this.CheckBox1.Checked = false;
                    //Globales.oApp.MessageBox("asdasd");

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;

                    if(oMatrix.Columns.Count > 0)
                    {
                        for (int colIndex = 1; colIndex <= oMatrix.Columns.Count - 1; colIndex++)
                        {
                            oMatrix.CommonSetting.SetCellEditable(18, colIndex, false);
                            oMatrix.CommonSetting.SetCellEditable(19, colIndex, false);
                            oMatrix.CommonSetting.SetCellEditable(20, colIndex, false);
                            oMatrix.CommonSetting.SetCellEditable(21, colIndex, false);
                        }
                    }

                    Form_ResizeAfter(pVal);
                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }
            //throw new System.NotImplementedException();



        }

        private void CheckBox1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox1, pVal, 3);
        }


        private void CheckBoxManagement(SAPbouiCOM.CheckBox checkBox, SAPbouiCOM.SBOItemEventArg pVal, int column)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            try
            {
                var asdasd = pVal.InnerEvent;
                if (pVal.FormMode == 3 && pVal.InnerEvent == false)
                {
                    

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
                    double sumar = 0;
                    for (int i = 1; i <= oMatrix.RowCount - 4; i++)
                    {
                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(column).Cells.Item(i).Specific; //Cast the Cell of 
                        sumar = sumar + double.Parse(oEditText.Value.ToString());
                    }

                    if (sumar == 0)
                    {
                        checkBox.Checked = false;
                        Globales.oApp.MessageBox("El proceso de corrida, no es válido, porque todos los valores estan en cero.");
                        
                    }
                        
                    else
                    {
                        for (int i = 1; i <= oMatrix.RowCount - 4; i++)
                        {
                            oMatrix.CommonSetting.SetCellEditable(i, column, checkBox.Checked);
                        }
                    }
                    
                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }
        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;

            try
            {
                this.Button2.Item.Enabled = true;
                this.Button2.Item.Enabled = true;

                List<string> listaLotes = new List<string>();
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_51").Specific;

                for (int m = 1; m <= oMatrix.RowCount; m++)
                {
                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(1).Cells.Item(m).Specific;
                    string cellValue = oEditText.Value.ToString();

                    listaLotes.Add(oEditText.Value.ToString());

                }


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;


                for (int j = 1; j <= oMatrix.RowCount; j++)
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

        }

        private void CheckBox0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox0, pVal, 2);

        }

        private void CheckBox2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox2, pVal, 4);

        }

        private void CheckBox3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox3, pVal, 5);

        }

        private void CheckBox4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox4, pVal, 6);

        }

        private void CheckBox5_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox5, pVal, 7);

        }

        private void CheckBox6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox6, pVal, 8);

        }

        private void CheckBox7_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox7, pVal, 9);

        }

        private void CheckBox8_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox8, pVal, 10);

        }

        private void CheckBox9_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox9, pVal,11);
           

        }

        private void Matrix1_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;

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

                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_RCORESU");
                        string calculo = (Math.Round((columnValue5 / columnValue4) * columnValue7, 2)).ToString();
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

        private void Matrix1_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;

                if (pVal.ColUID == "Col_5" || pVal.ColUID == "Col_6")
                {


                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(6).Cells.Item(pVal.Row).Specific;
                    int columnValue6 = FuncionesComunes.ValidateNumberInt(oEditText.Value.ToString());

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

                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MGS_CL_RCORESU");
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


                    this.EditText7.Value = sum_ba.ToString();
                    this.EditText8.Value = sum_bv.ToString();
                    this.EditText9.Value = Math.Round(sum_cv, 2).ToString();
                    this.EditText12.Value = (Math.Round(sum_cv, 2) * double.Parse(this.EditText11.Value.ToString())).ToString();
                    this.EditText13.Value = (double.Parse(this.EditText12.Value.ToString()) * 1.16).ToString();
                    this.EditText14.Value = (100 * double.Parse(this.EditText9.Value.ToString()) / double.Parse(this.EditText6.Value.ToString())).ToString();
                    // this.EditText8.Value = (double.Parse(this.EditText9.Value.ToString()) * 1.16).ToString();
                    //this.EditText9.Value = (100 * double.Parse(this.EditText8.Value.ToString()) / double.Parse(this.EditText16.Value.ToString())).ToString();


                }

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());

            }

        }

        

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.ComboBox oCombo = null;
            //throw new System.NotImplementedException();



            if (pVal.FormMode == 3)
            {

                List<string> oListErr = new List<string>();


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;
                if (oMatrix.RowCount == 0)
                {
                    oListErr.Add("Debe especificar informacion en la pestaña RESUMEN");
                }
                else
                {
                    //SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbComboBox").Specific;
                    //string valorSeleccionado = oComboBox.Selected.Value;
                    int contar = 0;
                    for (int j = 1; j <= oMatrix.RowCount; j++)
                    {
                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(j).Specific;
                        if (oCombo.Selected == null)
                            contar++;
                    }
                    if (contar > 0)
                        oListErr.Add("En el RESUMEN, debe seleccionar los lotes");

                    if (int.Parse(this.EditText1.Value) != (int.Parse(this.EditText7.Value) + int.Parse(this.EditText8.Value)))
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
                {
                    GenerateSalida();
                    //GenerateEntrada();
                    BubbleEvent = true;
                }
                    
            }
            else
            {
                BubbleEvent = true;
            }

        }

        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText22;
        private SAPbouiCOM.EditText EditText19;

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Documents oInventoryExit = null;
            SAPbobsCOM.Recordset oRS = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            int iErrCod;
            string sErrMsg = "";
            try
            {
                oInventoryExit = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //Solicitud sol = Comunes.FuncionesComunes.GetUDOSolicitudAgendada(agendado.DocEntry);

                if (Globales.oApp.MessageBox("¿Esta Ud. seguro de generar la salida de inventario?", 1, "Continuar", "Cancelar", "") == 1)
                {
                    // oInventoryExit.CardCode = this.EditText2.Value.ToString();
                    oInventoryExit.DocDate = DateTime.Now;
                    oInventoryExit.TaxDate = DateTime.Now;

                    string modelo = "";
                    string ccrCod = "";
                    oRS.DoQuery(Comunes.Consultas.GetItemData(this.EditText4.Value.ToString()));
                    if (oRS.RecordCount > 0)
                    {
                        modelo = oRS.Fields.Item(0).Value.ToString();
                        ccrCod = oRS.Fields.Item(1).Value.ToString();
                    }


                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_51").Specific;

                    int salesPersonCode = 1;
                    var asdas = oMatrix.RowCount;
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {

                        oInventoryExit.Lines.ItemCode = this.EditText4.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(3).Cells.Item(i).Specific;
                        oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_LARGO").Value = oEditText.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(2).Cells.Item(i).Specific;
                        oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_ANCHO").Value = oEditText.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(4).Cells.Item(i).Specific;
                        oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_CANBOB").Value = oEditText.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(i).Specific;
                        oInventoryExit.Lines.Quantity = double.Parse(oEditText.Value.ToString());

                        oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_MODELO").Value = modelo;
                        oInventoryExit.Lines.CostingCode = ccrCod;

                        oInventoryExit.Lines.WarehouseCode = "CORTE";
                        oInventoryExit.Lines.AccountCode = "5110190";

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                        oInventoryExit.Lines.BatchNumbers.BatchNumber = oEditText.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(i).Specific;
                        oInventoryExit.Lines.BatchNumbers.Quantity = double.Parse(oEditText.Value.ToString());

                        oInventoryExit.Lines.Add();
                    }



                    iErrCod = oInventoryExit.Add();
                    if (iErrCod != 0)
                    {

                        Globales.oCompany.GetLastError(out iErrCod, out sErrMsg);


                        Globales.oApp.MessageBox("Salida de inventario: " + sErrMsg);

                    }
                    else
                    {

                        oInventoryExit.GetByKey(int.Parse(Globales.oCompany.GetNewObjectKey()));

                        var sfsdf = Globales.oCompany.GetNewObjectKey();

                        Comunes.FuncionesComunes.UpdateUDORecibo(this.EditText0.Value, Globales.oCompany.GetNewObjectKey(), "U_MGS_CL_REFSAL");

                        Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se generó la salida con éxito",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    }

                }



            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }

        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.LinkedButton LinkedButton1;

        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            List<string[,]>  oListErr = null;
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
                oPB = Clases.Globales.oApp.StatusBar.CreateProgressBar("Generando la entrada de mercancia", 27, true);
                oPB.Text = "Generando la entrada de mercancia";
                oPB.Value = 10;
                if (Globales.oApp.MessageBox("¿Esta Ud. seguro de generar la entrada de inventario?", 1, "Continuar", "Cancelar", "") == 1)
                {
                    oSalesOpportunity = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                    oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oSalesOpportunity.TaxDate = DateTime.Now;
                    oSalesOpportunity.DocDate = DateTime.Now;

                    string modelo = "";
                    string ccrCod = "";
                    oRS.DoQuery(Comunes.Consultas.GetItemData(this.EditText4.Value.ToString()));
                    if (oRS.RecordCount > 0)
                    {
                        modelo = oRS.Fields.Item(0).Value.ToString();
                        ccrCod = oRS.Fields.Item(1).Value.ToString();
                    }


                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;

               
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

                        //oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                        //var sdsd = oCombo.Selected.Value.ToString();
                        //oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_LOTE").Value = oCombo.Selected.Value.ToString();

                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                        var sdsd = oCombo.Selected.Value.ToString();
                        oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_LOTE").Value = oCombo.Selected.Value.ToString();


                        oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_MODELO").Value = modelo;
                        oSalesOpportunity.Lines.CostingCode = ccrCod;

                        oSalesOpportunity.Lines.WarehouseCode = "CORTE";
                        oSalesOpportunity.Lines.AccountCode = "5110190";

                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                        oSalesOpportunity.Lines.BatchNumbers.BatchNumber = oCombo.Selected.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(8).Cells.Item(i).Specific;
                        oSalesOpportunity.Lines.BatchNumbers.Quantity = double.Parse(oEditText.Value.ToString());

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


                        Globales.oApp.MessageBox("Entrada de inventario: " + sErrMsg);

                    }
                    else
                    {

                        oSalesOpportunity.GetByKey(int.Parse(Globales.oCompany.GetNewObjectKey()));

                        var sfsdf = Globales.oCompany.GetNewObjectKey();

                        Comunes.FuncionesComunes.UpdateUDORecibo(this.EditText0.Value, Globales.oCompany.GetNewObjectKey(), "U_MGS_CL_REFENT");

                        Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se generó la entrada de mercancía con éxito ",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                    }
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

        private SAPbouiCOM.LinkedButton LinkedButton2;
        private SAPbouiCOM.StaticText StaticText23;
        private SAPbouiCOM.EditText EditText20;
        private SAPbouiCOM.Button Button4;


        private bool generateDocuments()
        {
            //try
            //{

            //}
            //catch ()
            //{

            //}

            return true;
        }

        private bool GenerateSalida()
        {

            SAPbobsCOM.Documents oInventoryExit = null;
            SAPbobsCOM.Recordset oRS = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            
            int iErrCod;
            string sErrMsg = "";
            bool rpta = true;
            try
            {
                oInventoryExit = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oInventoryExit.DocDate = DateTime.Now;
                oInventoryExit.TaxDate = DateTime.Now;

                string modelo = "";
                string ccrCod = "";
                oRS.DoQuery(Comunes.Consultas.GetItemData(this.EditText4.Value.ToString()));
                if (oRS.RecordCount > 0)
                {
                    modelo = oRS.Fields.Item(0).Value.ToString();
                    ccrCod = oRS.Fields.Item(1).Value.ToString();
                }


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_51").Specific;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {

                    oInventoryExit.Lines.ItemCode = this.EditText4.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(3).Cells.Item(i).Specific;
                    oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_LARGO").Value = oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(2).Cells.Item(i).Specific;
                    oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_ANCHO").Value = oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(4).Cells.Item(i).Specific;
                    oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_CANBOB").Value = oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(i).Specific;
                    oInventoryExit.Lines.Quantity = 1; // double.Parse(oEditText.Value.ToString());

                    oInventoryExit.Lines.UserFields.Fields.Item("U_MGS_CL_MODELO").Value = modelo;
                    oInventoryExit.Lines.CostingCode = ccrCod;

                    oInventoryExit.Lines.WarehouseCode = "CORTE";
                    oInventoryExit.Lines.AccountCode = "5110190";

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                    oInventoryExit.Lines.BatchNumbers.BatchNumber =  oEditText.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(i).Specific;
                    oInventoryExit.Lines.BatchNumbers.Quantity = 1; // double.Parse(oEditText.Value.ToString());

                    oInventoryExit.Lines.Add();
                }



                iErrCod = oInventoryExit.Add();
                if (iErrCod != 0)
                {

                    Globales.oCompany.GetLastError(out iErrCod, out sErrMsg);


                    Globales.oApp.MessageBox("Salida de inventario: " + sErrMsg);
                    rpta = false;

                }
                else
                {

                    oInventoryExit.GetByKey(int.Parse(Globales.oCompany.GetNewObjectKey()));

                    

                    var sfsdf = Globales.oCompany.GetNewObjectKey();

                    //Comunes.FuncionesComunes.UpdateUDORecibo(this.EditText0.Value, Globales.oCompany.GetNewObjectKey(), "U_MGS_CL_REFSAL");

                    Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se generó la salida con éxito",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    rpta = true;
                }

                GenerateEntrada();


            }
            catch (Exception ex)
            {
               // Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                throw ex;
                rpta = false;
            }

            return rpta;
        }

        private bool GenerateEntrada()
        {
            
            SAPbobsCOM.Documents oSalesOpportunity = null;
            //SalesOpportunities
            SAPbouiCOM.Matrix oMatrix = null;
            int iErrCod;
            string sErrMsg = "";
            SAPbouiCOM.ProgressBar oPB = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbobsCOM.Recordset oRS = null;
            bool rpta = true;
            try
            {
                oPB = Clases.Globales.oApp.StatusBar.CreateProgressBar("Generando la entrada de mercancia", 27, true);
                oPB.Text = "Generando la entrada de mercancia";
                oPB.Value = 10;

                oSalesOpportunity = (SAPbobsCOM.Documents)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oSalesOpportunity.TaxDate = DateTime.Now;
                oSalesOpportunity.DocDate = DateTime.Now;

                string modelo = "";
                string ccrCod = "";
                oRS.DoQuery(Comunes.Consultas.GetItemData(this.EditText4.Value.ToString()));
                if (oRS.RecordCount > 0)
                {
                    modelo = oRS.Fields.Item(0).Value.ToString();
                    ccrCod = oRS.Fields.Item(1).Value.ToString();
                }


                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;

            
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
                    oSalesOpportunity.Lines.Quantity = 1; // double.Parse(oEditText.Value.ToString());

                    //oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                    //var sdsd = oCombo.Selected.Value.ToString();
                    //oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_LOTE").Value = oCombo.Selected.Value.ToString();

                    oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                    var sdsd = oCombo.Selected.Value.ToString();
                    oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_LOTE").Value = oCombo.Selected.Value.ToString();


                    oSalesOpportunity.Lines.UserFields.Fields.Item("U_MGS_CL_MODELO").Value = modelo;
                    oSalesOpportunity.Lines.CostingCode = ccrCod;

                    oSalesOpportunity.Lines.WarehouseCode = "CORTE";
                    oSalesOpportunity.Lines.AccountCode = "5110190";

                    oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(1).Cells.Item(i).Specific;
                    oSalesOpportunity.Lines.BatchNumbers.BatchNumber = oCombo.Selected.Value.ToString();

                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(8).Cells.Item(i).Specific;
                    oSalesOpportunity.Lines.BatchNumbers.Quantity = 1; // double.Parse(oEditText.Value.ToString());

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
                    rpta = false;

                    Globales.oApp.MessageBox("Entrada de inventario: " + sErrMsg);

                }
                else
                {

                    oSalesOpportunity.GetByKey(int.Parse(Globales.oCompany.GetNewObjectKey()));

                    var sfsdf = Globales.oCompany.GetNewObjectKey();

                   // Comunes.FuncionesComunes.UpdateUDORecibo(this.EditText0.Value, Globales.oCompany.GetNewObjectKey(), "U_MGS_CL_REFENT");

                    Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se generó la entrada de mercancía con éxito ",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    rpta = true;


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
                throw ex;
                rpta = false;
                //Comunes.Funciones.RemoveRegisterUDO(docEntryUDO);
            }

            return rpta;
        }
        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            SAPbouiCOM.Matrix oMatrix = null;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
            var mat_left = oForm.Items.Item("Item_17").Left ;
            var col_0 = oMatrix.Columns.Item("Col_0").Width;
            var col_1 = oMatrix.Columns.Item("Col_1").Width;
            var inicial = mat_left + col_0 + col_1 / 2;

            oForm.Items.Item("Item_29").Left =  inicial;
            oForm.Items.Item("Item_30").Left =  inicial + col_1;
            oForm.Items.Item("Item_31").Left =  inicial + col_1*2;
            oForm.Items.Item("Item_32").Left =  inicial + col_1 * 3;
            oForm.Items.Item("Item_33").Left =  inicial + col_1 * 4;
            oForm.Items.Item("Item_34").Left =  inicial + col_1 * 5;
            oForm.Items.Item("Item_35").Left =  inicial + col_1 * 6;
            oForm.Items.Item("Item_36").Left =  inicial + col_1 * 7;
            oForm.Items.Item("Item_37").Left =  inicial + col_1 * 8;
            oForm.Items.Item("Item_38").Left =  inicial + col_1 * 9;
            oForm.Items.Item("Item_76").Left =  inicial + col_1 * 10;



        }

        private SAPbouiCOM.StaticText StaticText24;
        private SAPbouiCOM.StaticText StaticText25;
        private SAPbouiCOM.StaticText StaticText26;
        private SAPbouiCOM.StaticText StaticText27;
        private SAPbouiCOM.EditText EditText21;
        private SAPbouiCOM.EditText EditText22;
        private SAPbouiCOM.CheckBox CheckBox10;

        private void CheckBox10_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CheckBoxManagement(this.CheckBox10, pVal, 12);

        }
    }
}
