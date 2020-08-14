using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace CargarReporteSTAD
{
    class ClMonitoreo:ClConexion
    {

        private DateTime StringToDate(string Date, string Time)
        {
            DateTime date;
            int Año, Mes, Dia, Hora, Minuto, Segundo;
            Año = int.Parse(Date.Substring(6, 4));
            Mes = int.Parse(Date.Substring(3, 2));
            Dia = int.Parse(Date.Substring(0, 2));
            Hora = int.Parse(Time.Substring(0, 2));
            Minuto = int.Parse(Time.Substring(3, 2));
            Segundo = int.Parse(Time.Substring(6, 2));
            date = new DateTime(Año, Mes, Dia, Hora, Minuto, Segundo);
            return date;
        }

        private int StringToInt(string Seconds)
        {
            Seconds = Seconds.Replace(",", "");
            int Time = int.Parse(Seconds);
            return Time;
        }

        private double StringToDouble(string Transfered)
        {
            Transfered = Transfered.Replace(",", "");
            double KBytes = double.Parse(Transfered);
            return KBytes;
        }

        public int Select_Prerequisitos()
        {
            int Prerequisitos = 0;
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Select Count(*) As PRE from STAD_STEPS Where Programa in (Select Distinct ProgramaAnterior from STAD_STEPS Where Programa = 'CargarReporteSTAD') And Efectuado = ' '";
                Prerequisitos = (int)Comando.ExecuteScalar();
                Close();
            }
            catch
            {
                Close();
                Prerequisitos = -1;
            }
            return Prerequisitos;
        }

        public int Select_Finalizado()
        {
            int Finalizado = 0;
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Select Count(*) As FIN from STAD_STEPS Where Efectuado = ' ' And Programa = 'CargarReporteSTAD'";
                Finalizado = (int)Comando.ExecuteScalar();
                Close();
            }
            catch
            {
                Close();
                Finalizado = -1;
            }
            return Finalizado;
        }

        public DataTable Select_Next_Step()
        {
            DataTable Next_Step = new DataTable();
            try 
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Select Top 1 * from STAD_STEPS Where Programa = 'CargarReporteSTAD' And Efectuado in  (' ', '-') Order by Paso";
                Adaptador.Fill(Next_Step);
                Close();
            }
            catch 
            {
                Close();
            }
            return Next_Step;
        }

        public bool Update_Next_Step(int Paso, string Status)
        {
            bool resultado = false;
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Update STAD_STEPS Set Efectuado = '" + Status + "' Where Paso = " + Paso.ToString() + " And Programa = 'CargarReporteSTAD'";
                Comando.ExecuteNonQuery();
                Close();
                resultado = true;
            }
            catch
            {
                Close();
                resultado = false;
            }
            return resultado;
        }

        public bool Insert_STAD(string Fecha, string Started, string Server, string Transaction, string Program, string TScreen, string Screen, string WP, string User, string ResponseTime, string TimeInWPS, string WaitTime, string CPUTime, string DBReqTime, string VMCelapsedtime, string MemoryUsed, string TransferedKBytes, string Mandante)
        {
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Insert into STAD Values (@FECHA, @SERVIDOR, @TRANSACCION, @PROGRAMA, @TPANTALLA, @PANTALLA, @WP, @USUARIO, @TRESPUESTAMS, @TWPMS, @TESPERAMS, @TCPUMS, @TDBMS, @TVMCMS, @MEMORIAKB, @TRANSFERIDOSKB, @MANDANTE)";
                Comando.Parameters.Add("FECHA", SqlDbType.DateTime).Value = StringToDate(Fecha, Started);
                Comando.Parameters.Add("SERVIDOR", SqlDbType.NVarChar, 17).Value = Server;
                Comando.Parameters.Add("TRANSACCION", SqlDbType.NVarChar, 21).Value = Transaction;
                Comando.Parameters.Add("PROGRAMA", SqlDbType.NVarChar, 41).Value = Program;
                Comando.Parameters.Add("TPANTALLA", SqlDbType.NVarChar, 1).Value = TScreen;
                Comando.Parameters.Add("PANTALLA", SqlDbType.Int).Value = StringToInt(Screen);
                Comando.Parameters.Add("WP", SqlDbType.Int).Value = StringToInt(WP);
                Comando.Parameters.Add("USUARIO", SqlDbType.NVarChar, 12).Value = User;
                Comando.Parameters.Add("TRESPUESTAMS", SqlDbType.Int).Value = StringToInt(ResponseTime);
                Comando.Parameters.Add("TWPMS", SqlDbType.Int).Value = StringToInt(TimeInWPS);
                Comando.Parameters.Add("TESPERAMS", SqlDbType.Int).Value = StringToInt(WaitTime);
                Comando.Parameters.Add("TCPUMS", SqlDbType.Int).Value = StringToInt(CPUTime);
                Comando.Parameters.Add("TDBMS", SqlDbType.Int).Value = StringToInt(DBReqTime);
                Comando.Parameters.Add("TVMCMS", SqlDbType.Int).Value = StringToInt(VMCelapsedtime);
                Comando.Parameters.Add("MEMORIAKB", SqlDbType.Int).Value = StringToInt(MemoryUsed);
                Comando.Parameters.Add("TRANSFERIDOSKB", SqlDbType.Real).Value = StringToDouble(TransferedKBytes);
                Comando.Parameters.Add("MANDANTE", SqlDbType.Int).Value = StringToInt(Mandante);
                Comando.ExecuteNonQuery();
                Close();
                return true;
            }
            catch
            {
                Close();
                return false;
            }
        }
    }
}
