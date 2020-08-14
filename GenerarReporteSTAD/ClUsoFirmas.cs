using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Data;

namespace GenerarReporteSTAD
{
    class ClUsoFirmas:ClConexionUsoFirmas
    {

        public DataTable Select(string Select)
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
    }
}
