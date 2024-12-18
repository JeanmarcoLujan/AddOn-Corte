﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Comunes
{
    class Consultas
    {
        #region Atributos
        private static StringBuilder m_sSQL = new StringBuilder(); //Variable para la construccion de strings
        #endregion

        public static string ConsultaTablaConfiguracion(SAPbobsCOM.BoDataServerTypes bo_ServerTypes, string NAddon, string Version, bool Ordenamiento)
        {
            m_sSQL.Length = 0;

            //ES UNA PRUEBA DE AGRITOP

            switch (bo_ServerTypes)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    m_sSQL.AppendFormat("SELECT * FROM \"@{0}\"", NAddon.ToUpper());
                    if (NAddon != "" || Version != "")
                    {
                        m_sSQL.Append(" WHERE ");
                        if (NAddon != "")
                        {
                            m_sSQL.AppendFormat("\"Name\" Like '{0}%'", NAddon);
                            if (Version != "") m_sSQL.AppendFormat(" AND \"Code\" = '{0}'", Version);
                        }
                        else if (Version != "") m_sSQL.AppendFormat("\"Code\" = '{0}'", Version);
                    }
                    if (Ordenamiento) m_sSQL.Append(" ORDER BY CAST( REPLACE(\"Code\", '.' , '') AS NUMERIC) DESC");
                    break;
                default:
                    m_sSQL.AppendFormat("SELECT * FROM [@{0}]", NAddon.ToUpper());
                    if (NAddon != "" || Version != "")
                    {
                        m_sSQL.Append(" WHERE ");
                        if (NAddon != "")
                        {
                            m_sSQL.AppendFormat("name Like '{0}%'", NAddon);
                            if (Version != "") m_sSQL.AppendFormat(" AND code = '{0}'", Version);
                        }
                        else if (Version != "") m_sSQL.AppendFormat("code = '{0}'", Version);
                    }
                    if (Ordenamiento) m_sSQL.Append(" ORDER BY CONVERT(NUMERIC, REPLACE(Code, '.' , '')) DESC");
                    break;
            }

            return m_sSQL.ToString();
        }

        public static string GetAllLotesByItem(string itemCode)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('LOTES_INICIAL','{0}','','',''); ", itemCode);

            return m_sSQL.ToString();
        }

        public static string GetItemData(string itemCode)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('ITEM_DATA','{0}','','',''); ", itemCode);

            return m_sSQL.ToString();
        }

        public static string GetAlmacenCore()
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('ALMACEN_CORE','','','',''); ");

            return m_sSQL.ToString();
        }

        public static string ValidateArticuloLotes(string itemCode)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('VALIDATE_ART_LOTE','{0}','','',''); ", itemCode);

            return m_sSQL.ToString();
        }

        public static string GetSolicitudesAgenda(string estado)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('AGENDAR_SOLICITUDES','{0}','','',''); ", estado);

            return m_sSQL.ToString();
        }

        public static string GetPrecioArticulo(string itemCode, string cardCode)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('PRECIO_ARTICULO','{0}','{1}','',''); ", itemCode, cardCode);

            return m_sSQL.ToString();
        }

        public static string GetEquipoList()
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('EQUIPO_LIST','','','',''); ");

            return m_sSQL.ToString();
        }

        public static string GetAlmacenList()
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('ALMACEN_LIST','','','',''); ");

            return m_sSQL.ToString();
        }

        public static string GetEquipoSerie(string itemCode)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('EQUIPO_SERIE','{0}','','',''); ", itemCode);

            return m_sSQL.ToString();
        }

        public static string GetSolicitudCorte()
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('SOLICITUD_CORTE','','','',''); ");

            return m_sSQL.ToString();
        }

        public static string GetCortesCmb()
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('LISTAR_SOLICITUDES','','','',''); ");

            return m_sSQL.ToString();
        }

        public static string GetCortesCmbAll(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('LISTAR_SOLICI_ALL','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string GetPrecioOrdenVenta(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('PRECIO_OV','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string ValidateSalidaRef(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('SALIDA_REF','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string ValidateCorridasDelRecibo(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('RECIBO_PROD','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string ValidateArticulo(string itemCode)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('VALIDATE_ITEM','{0}','','',''); ", itemCode);

            return m_sSQL.ToString();
        }


        public static string GetUltimaCorrida(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('ULTIMA_CORRIDA','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string GetSumQtyEntregas(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('SUM_QTY_ENTREGA','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string GetAlmacenOfTrasferenciaAgenda(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('ALMACEN_DE_AGEN','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string ValidateDocCancel(string oferta)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('VALID_OFERTA','{0}','','',''); ", oferta);

            return m_sSQL.ToString();
        }

        public static string ValidateEntregaEnRecibosAnteriores(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('VALID_REC_ENT_ANT','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }

        public static string GetNumAtCardDoc(string indicator, string series)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('NUMATCARD','{0}','{1}','',''); ", indicator, series);

            return m_sSQL.ToString();
        }

        public static string GetNroCon(string indicator, string series)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('NROCON','{0}','{1}','',''); ", indicator, series);

            return m_sSQL.ToString();
        }

        public static string ValidateSolicitudCancel(string solicitud)
        {
            m_sSQL.Length = 0;

            m_sSQL.AppendFormat(" CALL MGS_HDB_PE_SP_ADDON_CORTE_GETVALUE ('VALID_SOL','{0}','','',''); ", solicitud);

            return m_sSQL.ToString();
        }
    }
}
