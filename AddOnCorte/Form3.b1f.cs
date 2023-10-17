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
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
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
        }

        private SAPbouiCOM.StaticText StaticText0;
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
            var asd = this.ComboBox0.Selected.Value;
            Globales.oApp.MessageBox("asdasd");

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
    }
}
