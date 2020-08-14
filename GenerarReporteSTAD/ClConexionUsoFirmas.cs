using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;

namespace GenerarReporteSTAD
{
    class ClConexionUsoFirmas
    {
        string Servidor = ConfigurationManager.AppSettings["server"].ToString();
        string BaseDeDatos = ConfigurationManager.AppSettings["staddb"].ToString();
        string Usuario = ConfigurationManager.AppSettings["user"].ToString();
        string Contraseña = ConfigurationManager.AppSettings["password"].ToString();

        protected SqlConnection Conexion;
        protected SqlCommand Comando;
        protected SqlDataAdapter Adaptador;

        public ClConexionUsoFirmas()
        {
            string ConnString = "Server=" + Servidor + ";Database=" + BaseDeDatos + ";User Id=" + Usuario + ";Password=" + Contraseña + ";";
            Conexion = new SqlConnection(ConnString);
            Comando = new SqlCommand("", Conexion);
            Comando.CommandTimeout = 90;
            Adaptador = new SqlDataAdapter(Comando);
        }

        public bool Open()
        {
            try
            {
                if (Conexion.State == System.Data.ConnectionState.Open)
                {
                    return true;
                }
                else
                {
                    Conexion.Open();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public bool Close()
        {
            try
            {
                if (Conexion.State == System.Data.ConnectionState.Open)
                {
                    Conexion.Close();
                    return true;
                }
                else
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
