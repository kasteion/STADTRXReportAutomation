using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace CargarReporteSTAD
{
    public partial class Form1 : Form
    {
        ClMonitoreo monitoreo = new ClMonitoreo();
        DateTime Fechus = DateTime.Now.AddDays(-1);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "A.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "A.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "B.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "B.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "C.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "C.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "D.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "D.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "E.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "E.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "F.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "F.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "G.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "G.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "H.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "H.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "I.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "I.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "J.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "J.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "K.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "K.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "L.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "L.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "M.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "M.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "N.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "N.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Ñ.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Ñ.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "O.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "O.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "P.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "P.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Q.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Q.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "R.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "R.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "S.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "S.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "T.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "T.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "U.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "U.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "V.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "V.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "W.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "W.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "X.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "X.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Y.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Y.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Z.txt", @"C:\Monitoreo\300" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + "Z.txt");
            //System.IO.File.Copy(@"\\C16078\STAD\500" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + ".txt", @"C:\Monitoreo\500" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + ".txt");
            //System.IO.File.Copy(@"\\C16078\STAD\800" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + ".txt", @"C:\Monitoreo\800" + Fechus.Year.ToString().Trim() + Fechus.Month.ToString().Trim().PadLeft(2, '0') + Fechus.Day.ToString().Trim().PadLeft(2, '0') + ".txt");
            timer.Start();
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\CargaReporteSTAD.log");
            int Pre = monitoreo.Select_Prerequisitos();
            int Fin = monitoreo.Select_Finalizado();
            if ((Pre == 0) && (Fin > 0))
            {
                sw.WriteLine(Pre.ToString());
                sw.WriteLine(Fin.ToString());
                try
                {
                    DataTable dt = monitoreo.Select_Next_Step();
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["Efectuado"].Equals(" "))
                        {
                            monitoreo.Update_Next_Step((int)dt.Rows[0]["Paso"], "-");
                            DateTime Fecha = (DateTime)dt.Rows[0]["FechaInicial"];
                            string Ruta = "";
                            switch (dt.Rows[0]["Letra"].ToString().Trim())
                            { 
                                case "A":
                                case "B":
                                case "C":
                                case "D":
                                case "E":
                                case "F":
                                case "G":
                                case "H":
                                case "I":
                                case "J":
                                case "K":
                                case "L":
                                case "M":
                                case "N":
                                case "Ñ":
                                case "O":
                                case "P":
                                case "Q":
                                case "R":
                                case "S":
                                case "T":
                                case "U":
                                case "V":
                                case "W":
                                case "X":
                                case "Y":
                                case "Z":
                                    Ruta = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\300" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + dt.Rows[0]["Letra"].ToString().Trim() + ".txt";
                                    break;
                                case "500":
                                    Ruta = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\500" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt";
                                    break;
                                case "800":
                                    Ruta = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\800" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt";
                                    break;
                            }
                            StreamReader sr = new StreamReader(Ruta, Encoding.UTF7);
                            //MessageBox.Show(Ruta);
                            string linea = "";
                            string fecha = "";
                            sr.ReadLine();
                            linea = sr.ReadLine().Trim();
                            if (linea.StartsWith("Analysed time:"))
                            {
                                fecha = linea.Replace("Analysed time:", "").Substring(0, 13).Trim();
                            }
                            sr.ReadLine();
                            sr.ReadLine();
                            sr.ReadLine();
                            sr.ReadLine();
                            sr.ReadLine();
                            sr.ReadLine();
                            sr.ReadLine();
                            while (!sr.EndOfStream)
                            {

                                linea = sr.ReadLine();
                                switch (linea)
                                {
                                    default:
                                        try
                                        {
                                            linea = linea.Substring(1, linea.Length - 1);
                                            //MessageBox.Show(linea);
                                            if (linea.StartsWith("Start of"))
                                            {
                                                fecha = linea.Replace("Start of", "").Substring(0, linea.Length - 9).Trim();
                                            }
                                            else
                                            {
                                                string started = linea.Substring(0, 8).Trim();
                                                linea = linea.Substring(8, linea.Length - 8);
                                                string server = linea.Substring(0, 18).Trim();
                                                linea = linea.Substring(18, linea.Length - 18);
                                                string transaction = linea.Substring(0, 21).Trim();
                                                linea = linea.Substring(21, linea.Length - 21);
                                                string program = linea.Substring(0, 41).Trim();
                                                linea = linea.Substring(41, linea.Length - 41);
                                                string TScreen = linea.Substring(0, 2).Trim();
                                                linea = linea.Substring(2, linea.Length - 2);
                                                string Screen = linea.Substring(0, 5).Trim();
                                                linea = linea.Substring(5, linea.Length - 5);
                                                string WP = linea.Substring(0, 2).Trim();
                                                linea = linea.Substring(2, linea.Length - 2);
                                                string User = linea.Substring(0, 13).Replace("|", "").Trim();
                                                linea = linea.Substring(13, linea.Length - 13);
                                                string ResponseTime = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string TimeInWPS = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string WaitTime = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string CPUTime = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string DBReqTime = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string VMCelapsed = linea.Substring(0, 12).Replace("|", "").Trim();
                                                linea = linea.Substring(12, linea.Length - 12);
                                                string MemoryUsed = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string TransferedKBytes = linea.Substring(0, 11).Replace("|", "").Trim();
                                                linea = linea.Substring(11, linea.Length - 11);
                                                string Client = "";
                                                switch (dt.Rows[0]["Letra"].ToString().Trim())
                                                {
                                                    case "A":
                                                    case "B":
                                                    case "C":
                                                    case "D":
                                                    case "E":
                                                    case "F":
                                                    case "G":
                                                    case "H":
                                                    case "I":
                                                    case "J":
                                                    case "K":
                                                    case "L":
                                                    case "M":
                                                    case "N":
                                                    case "Ñ":
                                                    case "O":
                                                    case "P":
                                                    case "Q":
                                                    case "R":
                                                    case "S":
                                                    case "T":
                                                    case "U":
                                                    case "V":
                                                    case "W":
                                                    case "X":
                                                    case "Y":
                                                    case "Z":
                                                        Client = "300";
                                                        break;
                                                    case "500":
                                                        Client = "500";
                                                        break;
                                                    case "800":
                                                        Client = "800";
                                                        break;
                                                }
                                                monitoreo.Insert_STAD(fecha, started, server, transaction, program, TScreen, Screen, WP, User, ResponseTime, TimeInWPS, WaitTime, CPUTime, DBReqTime, VMCelapsed, MemoryUsed, TransferedKBytes, Client);
                                                //MessageBox.Show("End");
                                            }
                                        }
                                        catch { }
                                        break;
                                }
                            }
                            sr.Close();
                            System.IO.File.Delete(Ruta);
                            monitoreo.Update_Next_Step((int)dt.Rows[0]["Paso"], "X");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                if ((Pre == 0) && (Fin == 0))
                {
                    this.Close();
                }
            }
        }
    }
}

