using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using MFiles.VAF.Common;

namespace Arkiva.MonitorFiscal.Checklist.Sql
{
    public class Conexion
    {
        //Cadena de conexion
        //string cadena =
        //    ConfigurationManager.
        //    ConnectionStrings["Arkiva.MonitorFiscal.Checklist.Properties.Settings.connectionString"].
        //    ConnectionString;

        string cadena = @"Data Source=WINDOWS-A28KAVE\SQLEXPRESS;Initial Catalog=Prueba;Persist Security Info=True;User ID=sa;Password=arkivasql";
        //string cadena = @"Data Source=MONITOR\SQLEXPRESS;Initial Catalog=MFSQL_MonitorBackOffice;Persist Security Info=True;User ID=sa;Password=Ark1V@.$Ql";
        //string cadena = @"Data Source=ARKIVA-HETZ-SQL;Initial Catalog=MFSQL_Monicom-Dev;Persist Security Info=True;User ID=sa;Password=Z4Qfk@DE$2tiuK";

        public SqlConnection conectar = new SqlConnection();

        //Constructor
        public Conexion()
        {
            conectar.ConnectionString = cadena;
        }

        //Metodo para abrir conexion
        public void AbrirConexion()
        {
            try
            {
                conectar.Open();
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error al conectar a la BD, " + ex);
            }
        }

        //Metodo para cerrar la conexion
        public void CerrarConexion()
        {
            try
            {
                conectar.Close();
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error al cerrar la conexion a la BD, " + ex);
            }
        }
    }
}
