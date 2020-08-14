using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Data;

namespace GenerarReporteSTAD
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

        //public bool Insert_STAD(string Fecha, string Started, string Server, string Transaction, string Program, string TScreen, string WP, string User, string ResponseTime, string TimeInWPS, string WaitTime, string CPUTime, string DBReqTime, string VMCelapsedtime, string MemoryUsed, string TransferedKBytes, string Mandante)
        //{
        //    try
        //    {
        //        Open();
        //        Comando.Parameters.Clear();
        //        Comando.CommandText = "Insert into STAD Values (@FECHA, @SERVIDOR, @TRANSACCION, @PROGRAMA, @TPANTALLA, @USUARIO, @TRESPUESTAMS, @TWPMS, @TESPERAMS, @TCPUMS, @TDBMS, @TVMCMS, @MEMORIAKB, @TRANSFERIDOSKB, @MANDANTE)";
        //        Comando.Parameters.Add("FECHA", SqlDbType.DateTime).Value = StringToDate(Fecha, Started);
        //        Comando.Parameters.Add("SERVIDOR", SqlDbType.NVarChar, 17).Value = Server;
        //        Comando.Parameters.Add("TRANSACCION", SqlDbType.NVarChar, 21).Value = Transaction;
        //        Comando.Parameters.Add("PROGRAMA", SqlDbType.NVarChar, 41).Value = Program;
        //        Comando.Parameters.Add("TPANTALLA", SqlDbType.NVarChar, 1).Value = TScreen;
        //        Comando.Parameters.Add("USUARIO", SqlDbType.NVarChar, 12).Value = User;
        //        Comando.Parameters.Add("TRESPUESTAMS", SqlDbType.Int).Value = StringToInt(ResponseTime);
        //        Comando.Parameters.Add("TWPMS", SqlDbType.Int).Value = StringToInt(TimeInWPS);
        //        Comando.Parameters.Add("TESPERAMS", SqlDbType.Int).Value = StringToInt(WaitTime);
        //        Comando.Parameters.Add("TCPUMS", SqlDbType.Int).Value = StringToInt(CPUTime);
        //        Comando.Parameters.Add("TDBMS", SqlDbType.Int).Value = StringToInt(DBReqTime);
        //        Comando.Parameters.Add("TVMCMS", SqlDbType.Int).Value = StringToInt(VMCelapsedtime);
        //        Comando.Parameters.Add("MEMORIAKB", SqlDbType.Int).Value = StringToInt(MemoryUsed);
        //        Comando.Parameters.Add("TRANSFERIDOSKB", SqlDbType.Real).Value = StringToDouble(TransferedKBytes);
        //        Comando.Parameters.Add("MANDANTE", SqlDbType.Int).Value = StringToInt(Mandante);
        //        Comando.ExecuteNonQuery();
        //        Close();
        //        return true;
        //    }
        //    catch
        //    {
        //        Close();
        //        return false;
        //    }
        //}

        public DataTable Select_STAD(string Select)
        {
            DataTable DT = new DataTable();
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = Select;
                Adaptador.Fill(DT);
                Close();
                return DT;
            }
            catch
            {
                Close();
                return DT;
            }
        }

        //public bool Delete_STAD()
        //{
        //    try
        //    {
        //        Open();
        //        Comando.CommandText = "Delete STAD";
        //        Comando.ExecuteNonQuery();
        //        Close();
        //        return true;
        //    }
        //    catch
        //    {
        //        Close();
        //        return false;
        //    }
        //}

        public bool Insert_Stad_Usuarios()
        { 
            try
            {
                Open();
                Comando.CommandText = "Insert into STAD_USUARIOS Select Distinct Usuario, '', '', '', '', '' from STAD Where Usuario Not In (Select Distinct Usuario from STAD_USUARIOS) Order by Usuario";
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

        public bool Update_Stad_Usuarios(string Nombre, string Unidad, string Departamento, string Gerencia, string Empresa, string Usuario)
        {
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Update STAD_USUARIOS Set Nombre = @NOMBRE, Unidad = @UNIDAD, Departamento = @DEPARTAMENTO, Gerencia = @GERENCIA, Empresa = @EMPRESA Where Usuario = @USUARIO";
                Comando.Parameters.Add("NOMBRE", SqlDbType.NVarChar, 40).Value = Nombre;
                Comando.Parameters.Add("UNIDAD", SqlDbType.NVarChar, 40).Value = Unidad;
                Comando.Parameters.Add("DEPARTAMENTO", SqlDbType.NVarChar, 40).Value = Departamento;
                Comando.Parameters.Add("GERENCIA", SqlDbType.NVarChar, 40).Value = Gerencia;
                Comando.Parameters.Add("EMPRESA", SqlDbType.NVarChar, 40).Value = Empresa;
                Comando.Parameters.Add("USUARIO", SqlDbType.NVarChar, 12).Value = Usuario;
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

        public bool Insert_Stad_Log(string Usuario)
        {
            try 
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Insert into STAD_LOG Select Distinct Fecha, Servidor, Transaccion, Programa, TPantalla, Usuario  from STAD Where Usuario = @USUARIO And Transaccion in (Select TCODE from STAD_TRANSACCIONES_RESTRINGIDAS) And Transaccion not in (Select TCODE from STAD_USUARIO_PERFIL A, STAD_PERFIL_TRANSACCION B Where Usuario = @USUARIO And  A.Perfil = B.Perfil)";
                Comando.Parameters.Add("USUARIO", SqlDbType.NVarChar, 12).Value = Usuario;
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

        public int Select_Prerequisitos()
        {
            int Prerequisitos = 0;
            try
            {
                Open();
                Comando.Parameters.Clear();
                Comando.CommandText = "Select Count(*) As PRE from STAD_STEPS Where Programa in (Select Distinct ProgramaAnterior from STAD_STEPS Where Programa = 'GenerarReporteSTAD') And Efectuado = ' '";
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
                Comando.CommandText = "Select Count(*) As FIN from STAD_STEPS Where Efectuado = ' ' And Programa = 'GenerarReporteSTAD'";
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
                Comando.CommandText = "Select Top 1 * from STAD_STEPS Where Programa = 'GenerarReporteSTAD' And Efectuado in  (' ', '-') Order by Paso";
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
                Comando.CommandText = "Update STAD_STEPS Set Efectuado = '" + Status + "' Where Paso = " + Paso.ToString() + " And Programa = 'GenerarReporteSTAD'";
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
    }
}
