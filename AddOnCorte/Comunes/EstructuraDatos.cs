﻿using AddOnCorte.Clases;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Comunes
{
    class EstructuraDatos
    {
        #region _Attributes_

        int m_iErrCode = 0;
        string m_sErrMsg = "";
        private string m_sNombreAddon = Properties.Resources.NombreAddon;
        private string m_sDescripcionUDTAddon = Properties.Resources.DescripcionTablaAddon;
        private string m_sVersion = Properties.Resources.VersionAddon;

        #endregion

        #region _Constructor_

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        public EstructuraDatos()
        {
            try
            {
                if (ValidaVersion(m_sNombreAddon, m_sDescripcionUDTAddon, m_sVersion))
                {
                    RegistrarVersion(m_sNombreAddon, m_sVersion);
                    CrearUDTAddOn();
                    CrearUDFAddOn();
                    CrearUDOAddOn();
                    PrecargarDatosAddOn();
                    // GenerateDataStructureReports();
                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        #endregion

        #region _Functions_

        /// <summary>
        /// Método para validar la versión instalada del AddOn dentro de la sociedad donde se está iniciando.
        /// </summary>
        /// <param name="NombreAddon">Nombre del AddOn que sale de los recursos del compilado.</param>
        /// <param name="Version">Versión del AddOn que sale de los recuros del compilado.</param>
        /// <returns>Returna TRUE o FALSE según el resultado de la operación de verificación en la sociedad.</returns>
        private bool ValidaVersion(string NombreAddon, string DescripUDTAddOn, string Version)
        {
            bool bRetorno = false;
            SAPbobsCOM.UserTable oUT = null;
            SAPbobsCOM.Recordset oRS = null;
            string NombreTabla = "";

            try
            {
                NombreTabla = NombreAddon.ToUpper();
                try
                {
                    oUT = Globales.oCompany.UserTables.Item(NombreTabla);
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("invalid field name")) oUT = null;
                    else throw ex;
                }

                if (oUT == null)
                {
                    CreaTablaMD(NombreTabla, DescripUDTAddOn, SAPbobsCOM.BoUTBTableType.bott_NoObject);
                    bRetorno = true;
                }
                else
                {
                    oRS = (SAPbobsCOM.Recordset)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery(Consultas.ConsultaTablaConfiguracion(Globales.oCompany.DbServerType, NombreAddon, "", true));
                    if (oRS.RecordCount == 0)
                    {
                        bRetorno = true;
                        Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Actualizará la esturctura de datos",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    else
                    {
                        if (int.Parse(Version.Replace(".", "").ToString()) > int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", "")))
                        {
                            bRetorno = true;
                            Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Actualizará la esturctura de datos",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }

                        if (int.Parse(Version.Replace(".", "").ToString()) < int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", "")))
                            Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Detectó una version superior para este Addon",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                FuncionesComunes.LiberarObjetoGenerico(oRS);
                oRS = null;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
            return bRetorno;
        }

        #endregion

        #region _Methods_

        /// <summary>
        /// Método para registar el AddOn y la versión del mismo dentro de la sociedad donde se está iniciando el AddOn.
        /// </summary>        
        /// <param name="NombreAddon">Nombre del AddOn que sale de los recursos del compilado.</param>
        /// <param name="Version">Versión del AddOn que sale de los recuros del compilado.</param>     
        private void RegistrarVersion(string NombreAddon, string Version)
        {
            SAPbobsCOM.UserTable oUT;
            string NombreTabla = "";
            try
            {
                NombreTabla = NombreAddon.ToUpper();
                oUT = Globales.oCompany.UserTables.Item(NombreTabla);
                oUT.Code = Version;
                oUT.Name = NombreAddon + " V-" + Version;
                m_iErrCode = oUT.Add();

                if (m_iErrCode == 0)
                    Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se ingreso un nuevo registro a la BD ",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                else
                    Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Error ingresar el registro en la tabla: "
                        + NombreTabla + ". Error: " + Globales.oCompany.GetLastErrorDescription().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de tablas de usuario (UDT) en la sociedad.
        /// </summary>
        private void CrearUDTAddOn()
        {
            try
            {
                //CreaTablaMD("MSS_TIPOCUENTA", "Tipo de Cuenta", SAPbobsCOM.BoUTBTableType.bott_NoObject);

                CreaTablaMD("NX_LIQU", "Liq proveedores CAB", SAPbobsCOM.BoUTBTableType.bott_Document);
                CreaTablaMD("NX_LIEM", "Liq proveedores EM", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                CreaTablaMD("NX_LISG", "Liq proveedores SGI", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

                CreaTablaMD("NX_BCPA", "Bonif/Castigo parametros", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                CreaTablaMD("NX_BYC", "Bonificacion y Castigo", SAPbobsCOM.BoUTBTableType.bott_NoObject);


                //CreaTablaMD("NX_ROMA", "Aplicacion romana I", SAPbobsCOM.BoUTBTableType.bott_Document);


            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de campos de usuarios (UDF) en la sociedad.
        /// </summary>
        private void CrearUDFAddOn()
        {
            try
            {

                //Cabecera
                string[] ValidValues = null;
                string[] ValidDescrip = null;
                CreaCampoMD("@NX_LIQU", "NX_PROV", "Proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_PRDE", "Nombre del proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_TTRI", "Tipo de trigo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIQU", "NX_PRCO", "Persona de contacto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 40, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIQU", "NX_LREF", "Referencia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                // CreaCampoMD("@NX_LIQU", "NX_MONE", "Moneda", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 40, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");

                ValidValues = new string[3] { "A", "F", "E" };
                ValidDescrip = new string[3] { "Abierto", "Facturado", "Enviado" };
                CreaCampoMD("@NX_LIQU", "NX_ESTD", "Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, SAPbobsCOM.BoYesNoEnum.tNO, ValidValues, ValidDescrip, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FCVE", "Fecha de vencimiento", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FCDO", "Fecha del documento", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_COMM", "Comentarios", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_TVTR", "Valor trigo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_TRAL", "Trigo a liquidar", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_BONI", "Bonificacion porcentaje", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_CAST", "Castigo porcentaje", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_VBCA", "Valor bonficacion/castigo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_PREQ", "Precio por Q", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_UNES", "Unidad estandar", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_VNET", "Valor neto", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_IVAP", "Iva pagado", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_VTOT", "Valor total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_ANTC", "Anticipos", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FLET", "Fletes", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_TPAG", "Total a pagar", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_RDEV", "Referencia devoluacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_RFAC", "Referencia factura", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_RNPR", "Precio nuevo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 200, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FOPR", "Serie factura", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FONU", "Correlativo factura", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FIND", "Indicator factura", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_FENV", "Fecha envio correo", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIQU", "NX_HENV", "Hora envio correo", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");



                //Detalle
                CreaCampoMD("@NX_LIEM", "NX_ENME", "Entrada de mercancía", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIEM", "NX_CANT", "Cantidad", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIEM", "NX_PRCO", "Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LIEM", "NX_DNME", "Numero EM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");

                //CreaCampoMD("@NX_LIEM", "NX_PRCN", "Nuevo precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");

                //CreaCampoMD("@NX_LIEM", "NX_ITEM", "Número de articulo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIEM", "NX_ITDE", "Descripción de articulo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIEM", "NX_QBTY", "Cantidad", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIEM", "NX_ALMC", "Almacén", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIEM", "NX_PRIC", "Precio por unidad", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                //CreaCampoMD("@NX_LIEM", "NX_INDC", "Indicardor de impuesto", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_ENME", "Entrada de mercancía", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_LIQU", "SGI code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_SEDI", "Sedimentacion", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_PEES", "Peso especifico", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_HUME", "Humedad", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_IMPR", "Impuereza", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_GRAP", "Grano partido", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_GRAD", "Grano daniado calor", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_GRAH", "Grano helado", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_GRAB", "Grano brotado", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_GRAN", "Grano punta negra", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_BOCA", "Bonificacion - castigo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_PTOL", "Porcentaje del total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_PROP", "Proporcion de la suma BC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_LISG", "NX_DNME", "Numero EM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");


                CreaCampoMD("@NX_BCPA", "NX_PARA", "Parametro", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");

                CreaCampoMD("@NX_BYC", "NX_Parametro", "Parametro", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_BYC", "NX_Val1", "Valor inicial", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_BYC", "NX_Val2", "Valor final", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_BYC", "NX_CASBON", "Cast/Bon", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");
                CreaCampoMD("@NX_BYC", "NX_Seq", "Secuencia", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, "");

                //U_NX_DNME
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de objetos de usuarios (UDO) en la sociedad.
        /// </summary>
        private void CrearUDOAddOn()
        {

            string[] FindColumns = null;
            string[] ChildTables = null;

            try
            {
                FindColumns = new string[2] { "DocEntry", "DocNum" };
                ChildTables = new string[2] { "NX_LIEM", "NX_LISG" };

                CreaUDOMD(
                    "NX_LIQU", //Code
                    "LiquidacionProveedores", //name
                    "NX_LIQU", //table name
                    FindColumns, // findcolumns
                    ChildTables, //childtable
                    SAPbobsCOM.BoYesNoEnum.tNO, // cancel 
                    SAPbobsCOM.BoYesNoEnum.tNO, //close
                    SAPbobsCOM.BoYesNoEnum.tYES, //delete
                    SAPbobsCOM.BoYesNoEnum.tYES, //defaultForm
                    null, //formcolumns
                    SAPbobsCOM.BoYesNoEnum.tYES, //canfind
                    SAPbobsCOM.BoYesNoEnum.tNO, //canlog
                    SAPbobsCOM.BoUDOObjType.boud_Document, //objectType
                    SAPbobsCOM.BoYesNoEnum.tNO, //managerseries
                    SAPbobsCOM.BoYesNoEnum.tNO, //enableEnhanced
                    SAPbobsCOM.BoYesNoEnum.tNO, //rebuildEnhanced
                    null, //ChilFormColumns 
                    null); //iChildBlock

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de objetos de usuarios (UDO) en la sociedad.
        /// </summary>
        private void PrecargarDatosAddOn()
        {
            try
            {


                PreloadUDO("NX_CONF", null, new string[] { "Code", "Name", "U_NX_ITMG", "U_NX_CPRC", "U_NX_SMTP", "U_NX_EMAI", "U_NX_CONT", "U_NX_PORT", "U_NX_DIRE", "U_NX_EMCC", "U_NX_RGUI", "U_NX_RNOT", "U_NX_RINF" }, new string[] { "NX", "NX", "ABR_114120", "Precio", "smtp.gmail.com", "sapmolinos@gmail.com", "Sap$2022", "587", "C:\\Users\\nxalujan\\Documents\\nx\\files\\", "rcoloma@inxap.com", "Guia recepcion", "Nota recepcion", "Informe laboratorio" }, null, null);

                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "1", "Sedimentación", "11" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "2", "Peso Especifico", "13" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "3", "Humedad", "6" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "4", "Impurezas", "15" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "5", "Granos Partidos", "16" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "6", "Dañados Calor", "17" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "7", "Helados/Inmaduros", "18" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "8", "Brotados", "19" }, null, null);
                PreloadUDO("NX_BCPA", null, new string[] { "Code", "Name", "U_NX_PARA" }, new string[] { "9", "Punta Negra", "20" }, null, null);




            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        //private string[] GetFormColumns(nx_TableName m_bo_TableName)
        //{
        //    string[] m_Detail = null;

        //    try
        //    {
        //        switch (m_bo_TableName)
        //        {
        //            case nx_TableName.tn_NX_ROMANA:
        //                m_Detail = new string[] {
        //                    "DocEntry",
        //                    "U_NX_IDPE",
        //                    "U_NX_PROV",
        //                    "U_NX_PRNA",
        //                    "U_NX_NUGU",
        //                    "U_NX_PATE",
        //                    "U_NX_CHRU",
        //                    "U_NX_CHNO",
        //                    "U_NX_WHSC",
        //                    "U_NX_OPOR",
        //                    "U_NX_ALAB",
        //                    "U_NX_NUML",
        //                    "U_NX_TIPT",
        //                    "U_NX_PREC",
        //                    "U_NX_ENME",
        //                    "U_NX_ESTD",
        //                    "U_NX_TREC",
        //                    "U_NX_PBRT",
        //                    "U_NX_PNET",
        //                    "U_NX_PREA",
        //                    "U_NX_FREG",
        //                    "U_NX_USER",
        //                    "U_NX_COMM"
        //                };
        //                break;
        //        }

        //        return m_Detail;
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }
        //}


        #endregion

        #region _Base Methods_

        /// <summary>
        /// Método para crear objetos de tipo Tablas de Usuarios (UDT) en la sociedad conectada.
        /// </summary>
        /// <param name="NombTabla">Nombre de la tabla de usuario.</param>
        /// <param name="DescTabla">Descripción de la tabla de usuario.</param>
        /// <param name="tipoTabla">Tipo de tabla de usuario.</param>
        /// <returns>Returna TRUE o FALSE como resultado de la operación de registro.</returns>
        private bool CreaTablaMD(string NombTabla, string DescTabla, SAPbobsCOM.BoUTBTableType tipoTabla)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            try
            {
                oUserTablesMD = null;
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!oUserTablesMD.GetByKey(NombTabla))
                {
                    oUserTablesMD.TableName = NombTabla;
                    oUserTablesMD.TableDescription = DescTabla;
                    oUserTablesMD.TableType = tipoTabla;
                    m_iErrCode = oUserTablesMD.Add();

                    if (m_iErrCode != 0)
                    {
                        Globales.oCompany.GetLastError(out m_iErrCode, out m_sErrMsg);
                        Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Error al crear  tabla: " + NombTabla,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else
                        Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se ha creado la tabla: " + NombTabla,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    FuncionesComunes.LiberarObjetoGenerico(oUserTablesMD);
                    oUserTablesMD = null;
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oUserTablesMD);
                oUserTablesMD = null;
            }
        }

        /// <summary>
        /// Método para crear objetos de tipo Campos de Usuarios (UDF) en la sociedad conectada.
        /// </summary>
        /// <param name="NombreTabla">Nombre de la tabla donde se creará el campo.</param>
        /// <param name="NombreCampo">Nombre del campo de usuario.</param>
        /// <param name="DescCampo">Descripción del campo de usario</param>
        /// <param name="TipoCampo">Tipo de campo de usuario.</param>
        /// <param name="SubTipo">Subtipo de campo de usuario.</param>
        /// <param name="Tamano">Tamaño del campo de usuario.</param>
        /// <param name="Obligatorio">Indicador de si el campo es obligatorio o no.</param>
        /// <param name="validValues">Arreglo de valores validos.</param>
        /// <param name="validDescription">Arreglo de descripción de valores validos.</param>
        /// <param name="valorPorDef">Valor por defecto.</param>
        /// <param name="tablaVinculada">Tabla vinculada al campo de usuario.</param>
        private void CreaCampoMD(string NombreTabla, string NombreCampo, string DescCampo, SAPbobsCOM.BoFieldTypes TipoCampo,
           SAPbobsCOM.BoFldSubTypes SubTipo, int Tamano, SAPbobsCOM.BoYesNoEnum Obligatorio, string[] validValues,
            string[] validDescription, string valorPorDef, string tablaVinculada)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;

            try
            {
                if (NombreTabla == null) NombreTabla = "";
                if (NombreCampo == null) NombreCampo = "";
                if (Tamano == 0) Tamano = 10;
                if (validValues == null) validValues = new string[0];
                if (validDescription == null) validDescription = new string[0];
                if (valorPorDef == null) valorPorDef = "";
                if (tablaVinculada == null) tablaVinculada = "";

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NombreTabla;
                oUserFieldsMD.Name = NombreCampo;
                oUserFieldsMD.Description = DescCampo;
                oUserFieldsMD.Type = TipoCampo;
                if (TipoCampo != SAPbobsCOM.BoFieldTypes.db_Date) oUserFieldsMD.EditSize = Tamano;
                oUserFieldsMD.SubType = SubTipo;

                if (tablaVinculada != "") oUserFieldsMD.LinkedTable = tablaVinculada;
                else
                {
                    if (validValues.Length > 0)
                    {
                        for (int i = 0; i <= (validValues.Length - 1); i++)
                        {
                            oUserFieldsMD.ValidValues.Value = validValues[i];
                            if (validDescription.Length > 0) oUserFieldsMD.ValidValues.Description = validDescription[i];
                            else oUserFieldsMD.ValidValues.Description = validValues[i];
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }
                    oUserFieldsMD.Mandatory = Obligatorio;
                    if (valorPorDef != "") oUserFieldsMD.DefaultValue = valorPorDef;
                }

                m_iErrCode = oUserFieldsMD.Add();

                if (m_iErrCode != 0)
                {
                    Globales.oCompany.GetLastError(out m_iErrCode, out m_sErrMsg);
                    if ((m_iErrCode != -5002) && (m_iErrCode != -2035))
                        Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Error al crear campo de usuario: " + NombreCampo
                            + "en la tabla: " + NombreTabla + " Error: " + m_sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                    Globales.oApp.StatusBar.SetText(AddOnCorte.Properties.Resources.NombreAddon + " Se ha creado el campo de usuario: " + NombreCampo
                            + " en la tabla: " + NombreTabla, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oUserFieldsMD);
                oUserFieldsMD = null;
            }
        }

        /// <summary>
        /// Método para crear objetos de tipo Objetos de Usuarios (UDO) en la sociedad conectada.
        /// </summary>
        /// <param name="s_Code"></param>
        /// <param name="s_Name"></param>
        /// <param name="s_TableName"></param>
        /// <param name="s_FindColumns"></param>
        /// <param name="s_ChildTables"></param>
        /// <param name="e_CanCancel"></param>
        /// <param name="e_CanClose"></param>
        /// <param name="e_CanDelete"></param>
        /// <param name="e_CanCreateDefaultForm"></param>
        /// <param name="s_FormColumns"></param>
        /// <param name="e_CanFind"></param>
        /// <param name="e_CanLog"></param>
        /// <param name="e_ObjectType"></param>
        /// <param name="e_ManageSeries"></param>
        /// <param name="e_EnableEnhancedForm"></param>
        /// <param name="e_RebuildEnhancedForm"></param>
        /// <param name="s_ChildFormColumns"></param>
        /// <param name="i_ChildBlock"></param>
        /// <returns></returns>
        private bool CreaUDOMD(string s_Code, string s_Name, string s_TableName, string[] s_FindColumns = null,
            string[] s_ChildTables = null, SAPbobsCOM.BoYesNoEnum e_CanCancel = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_CanClose = SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum e_CanDelete = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO, string[] s_FormColumns = null,
            SAPbobsCOM.BoYesNoEnum e_CanFind = SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum e_CanLog = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoUDOObjType e_ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData,
            SAPbobsCOM.BoYesNoEnum e_ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            string[] s_ChildFormColumns = null, int[] iChildBlock = null
            )
        {

            /* ,
            SAPbobsCOM.BoYesNoEnum e_EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            string[] s_ChildFormColumns = null, int[] iChildBlock= null */

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            int i_Result = 0;
            string s_Result = "";
            int beginChild = 0;

            try
            {
                oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Globales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                oUserObjectsMD.Code = "";


                if (!oUserObjectsMD.GetByKey(s_Code))
                {
                    oUserObjectsMD.Code = s_Code;
                    oUserObjectsMD.Name = s_Name;
                    oUserObjectsMD.ObjectType = e_ObjectType;
                    oUserObjectsMD.TableName = s_TableName;
                    oUserObjectsMD.CanCancel = e_CanCancel;
                    oUserObjectsMD.CanClose = e_CanClose;
                    oUserObjectsMD.CanDelete = e_CanDelete;
                    oUserObjectsMD.CanCreateDefaultForm = e_CanCreateDefaultForm;
                    oUserObjectsMD.EnableEnhancedForm = e_EnableEnhancedForm;
                    oUserObjectsMD.RebuildEnhancedForm = e_RebuildEnhancedForm;
                    oUserObjectsMD.CanFind = e_CanFind;
                    oUserObjectsMD.CanLog = e_CanLog;
                    oUserObjectsMD.ManageSeries = e_ManageSeries;

                    if (s_FindColumns != null)
                    {
                        for (int i = 0; i < s_FindColumns.Length - 1; i++)
                        {
                            oUserObjectsMD.FindColumns.ColumnAlias = s_FindColumns[i].ToString();
                            oUserObjectsMD.FindColumns.Add();
                        }
                    }

                    if (s_ChildTables != null)
                    {
                        for (int j = 0; j < s_ChildTables.Length; j++)
                        {
                            oUserObjectsMD.ChildTables.TableName = s_ChildTables[j];
                            oUserObjectsMD.FindColumns.Add();
                        }
                    }

                    if (s_FormColumns != null)
                    {
                        oUserObjectsMD.UseUniqueFormType = SAPbobsCOM.BoYesNoEnum.tYES;

                        for (int k = 0; k < s_FormColumns.Length; k++)
                        {
                            oUserObjectsMD.FormColumns.FormColumnAlias = s_FormColumns[k];
                            oUserObjectsMD.FormColumns.Add();
                        }
                    }

                    if (s_ChildFormColumns != null)
                    {
                        if (s_ChildTables != null)
                        {
                            beginChild = 1;

                            for (int i = 0; i < s_ChildFormColumns.Length; i++)
                            {
                                oUserObjectsMD.FormColumns.SonNumber = beginChild;
                                oUserObjectsMD.FormColumns.FormColumnAlias = s_ChildFormColumns[i];
                                oUserObjectsMD.FormColumns.Add();

                                if (iChildBlock[(beginChild - 1)].Equals((i + 1)))
                                {
                                    beginChild = beginChild + 1;
                                }
                            }
                        }
                    }

                    i_Result = oUserObjectsMD.Add();

                    if (!i_Result.Equals(0))
                    {
                        Globales.oCompany.GetLastError(out i_Result, out s_Result);
                        FuncionesComunes.DisplayErrorMessages((i_Result.ToString() + " - " + s_Result).ToString(), System.Reflection.MethodBase.GetCurrentMethod());
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oUserObjectsMD);
                oUserObjectsMD = null;
            }
        }

        /// <summary>
        /// Método para registrar valores validos en tablas de usuarios dentro de SAP B1.
        /// </summary>
        /// <param name="s_TableName">Nombre de la tabla de usuario.</param>
        /// <param name="s_CodeValue">Código del valor, representado en la columna Code de la tabla de usuario.</param>
        /// <param name="s_NameValue">Nombre o descripción del valor, representado en la columna Name de la tabla de usuario.</param>
        private void CargarValoresUDT(string s_TableName, string s_CodeValue, string s_NameValue)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            int i_Result = 0;
            int i_Error = 0;
            string s_Error = "";

            try
            {
                oUserTable = Globales.oCompany.UserTables.Item(s_TableName);
                if (!oUserTable.GetByKey(s_CodeValue))
                {
                    oUserTable.Code = s_CodeValue;
                    oUserTable.Name = s_NameValue;
                    i_Result = oUserTable.Add();

                    if (i_Result != 0)
                    {
                        Globales.oCompany.GetLastError(out i_Error, out s_Error);
                        FuncionesComunes.DisplayErrorMessages((i_Error.ToString() + s_Error).ToString(), System.Reflection.MethodBase.GetCurrentMethod());
                    }
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }


        private void PreloadUDO(string udoName, string udoChildName, string[] columnsUDOMain, string[] valuesUDOMain, string[] columnsUDOChild, string[] valuesUDOChid)
        {
            Metadata m_oMetadata = null;
            string errorMessage = "";

            try
            {
                m_oMetadata = new Metadata(Globales.oCompany);

                if (!m_oMetadata.PreloadUDO(udoName, udoChildName, columnsUDOMain, valuesUDOMain, columnsUDOChild,
                                            valuesUDOChid))
                {
                    if (m_oMetadata.IsError)
                    {
                        errorMessage = "[" + m_oMetadata.ErrorCode.ToString() + "] " + m_oMetadata.ErrorMessage;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                m_oMetadata = null;
            }
        }


        #endregion
    }
}
