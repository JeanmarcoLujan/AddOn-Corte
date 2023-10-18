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
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_15").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_16").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_17").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_3").Specific));
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
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_30").Specific));
            this.CheckBox2 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_31").Specific));
            this.CheckBox3 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_32").Specific));
            this.CheckBox4 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_33").Specific));
            this.CheckBox5 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_34").Specific));
            this.CheckBox6 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_35").Specific));
            this.CheckBox7 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_36").Specific));
            this.CheckBox8 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_37").Specific));
            this.CheckBox9 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_38").Specific));
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
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_49").Specific));
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
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_59").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            Funciones.CbxSolCorte(ref oForm, Globales.oApp, this.ComboBox0);
            this.CheckBox0.Checked = true;
            this.CheckBox1.Checked = true;
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
                if(pVal.FormMode == 3)
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

                        var sdfsdf = this.CheckBox0.Checked;

                        this.CheckBox0.Checked = true;
                        this.CheckBox1.Checked = true;

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

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }


                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_17").Specific;
                        oMatrix.Clear();
                        foreach (CorridasDetalle item in sol.DetalleCorridas)
                        {
                            
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

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }

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
                if (this.EditText3.Value.ToString() == "")
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
    }
}
