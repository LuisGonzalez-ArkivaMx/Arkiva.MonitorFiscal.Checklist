using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using MFiles.VAF.Common;

namespace Arkiva.MonitorFiscal.Checklist.Sql
{
    public class Query
    {
        Conexion cnx = new Conexion();

        public void InsertarDocumentosFaltantesChecklist(
            bool bDelete,
            string sProveedor = "", 
            int iProveedorID = 0, 
            string sProyecto = "",
            int iProyectoID = 0,
            string sEmpleado = "", 
            int iEmpleadoID = 0, 
            string sCategoria = "", 
            string sTipoDocumento = "", 
            string sNombreDocumento = "", 
            int iDocumentoID = 0,
            string sVigencia = "", 
            string sPeriodo = "")
        {
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();

            try
            {
                string sPeriodoFaltante = "";
                string sPeriodoFecha = "";

                if (sPeriodo == "Faltante")
                    sPeriodoFaltante = sPeriodo;
                else
                    sPeriodoFecha = sPeriodo;

                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@bDelete", bDelete);
                cmd.Parameters.AddWithValue("@Proveedor", sProveedor);
                cmd.Parameters.AddWithValue("@Proveedor_ID", iProveedorID);
                cmd.Parameters.AddWithValue("@Proyecto", sProyecto);
                cmd.Parameters.AddWithValue("@Proyecto_ID", iProyectoID);
                cmd.Parameters.AddWithValue("@Empleado", sEmpleado);
                cmd.Parameters.AddWithValue("@Empleado_ID", iEmpleadoID);
                cmd.Parameters.AddWithValue("@Categoria", sCategoria);
                cmd.Parameters.AddWithValue("@Tipo_Documento", sTipoDocumento);
                cmd.Parameters.AddWithValue("@Nombre_Documento", sNombreDocumento);
                cmd.Parameters.AddWithValue("@Documento_ID", iDocumentoID);
                cmd.Parameters.AddWithValue("@Vigencia", sVigencia);
                cmd.Parameters.AddWithValue("@Periodo", sPeriodoFaltante);
                cmd.Parameters.AddWithValue("@PeriodoFecha", sPeriodoFecha);
                cmd.CommandText = "spInsertaDocumentosFaltantesChecklist";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection = cnx.conectar;

                cnx.AbrirConexion();

                cmd.ExecuteNonQuery();

                cnx.CerrarConexion();
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error al insertar informacion de documentos faltantes, " + ex);
                cnx.CerrarConexion();
            }
        }

        public void InsertarDocumentosCaducados(
            string sProveedor = "",
            int iProveedorID = 0,
            string sProyecto = "",
            int iProyectoID = 0,
            string sEmpleado = "",
            int iEmpleadoID = 0,
            string sCategoria = "",
            string sTipoDocumento = "",
            string sNombreDocumento = "",
            int iDocumentoID = 0,
            string sVigencia = "",
            string sPeriodo = "")
        {
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();

            try
            {
                string sPeriodoFaltante = "";
                string sPeriodoFecha = "";

                if (sPeriodo == "Faltante")
                {
                    sPeriodoFaltante = sPeriodo;
                }                    
                else
                {
                    if (sPeriodo == "")
                    {
                        string sFechaActual = DateTime.Now.ToString("yyyy-MM-dd");
                        sPeriodoFecha = sFechaActual;
                    }
                    else
                    {
                        sPeriodoFecha = sPeriodo;
                    }                    
                }                    

                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@Proveedor", sProveedor);
                cmd.Parameters.AddWithValue("@Proveedor_ID", iProveedorID);
                cmd.Parameters.AddWithValue("@Proyecto", sProyecto);
                cmd.Parameters.AddWithValue("@Proyecto_ID", iProyectoID);
                cmd.Parameters.AddWithValue("@Empleado", sEmpleado);
                cmd.Parameters.AddWithValue("@Empleado_ID", iEmpleadoID);
                cmd.Parameters.AddWithValue("@Categoria", sCategoria);
                cmd.Parameters.AddWithValue("@Tipo_Documento", sTipoDocumento);
                cmd.Parameters.AddWithValue("@Nombre_Documento", sNombreDocumento);
                cmd.Parameters.AddWithValue("@Documento_ID", iDocumentoID);
                cmd.Parameters.AddWithValue("@Vigencia", sVigencia);
                cmd.Parameters.AddWithValue("@Periodo", sPeriodoFaltante);
                cmd.Parameters.AddWithValue("@PeriodoFecha", sPeriodoFecha);
                cmd.CommandText = "spInsertaDocumentosCaducados";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection = cnx.conectar;

                cnx.AbrirConexion();

                cmd.ExecuteNonQuery();

                cnx.CerrarConexion();
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error al insertar informacion de documentos caducados, " + ex);
                cnx.CerrarConexion();
            }
        }

        public void InsertarComprobantesPago(
            string sComprobante, 
            int iComprobanteID, 
            string sProveedor, 
            int iProveedorID,
            int iDocumentoRelacionadoID = 0)
        {
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();

            try
            {
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@Comprobante", sComprobante);
                cmd.Parameters.AddWithValue("@Comprobante_ID", iComprobanteID);
                cmd.Parameters.AddWithValue("@Proveedor", sProveedor);
                cmd.Parameters.AddWithValue("@Proveedor_ID", iProveedorID);
                cmd.Parameters.AddWithValue("@Documento_relacionado_ID", iDocumentoRelacionadoID);
                cmd.CommandText = "spInsertaComprobantesPago";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection = cnx.conectar;

                cnx.AbrirConexion();

                cmd.ExecuteNonQuery();

                cnx.CerrarConexion();
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error al insertar informacion de comprobantes de pago, " + ex);
                cnx.CerrarConexion();
            }
        }
    }
}
