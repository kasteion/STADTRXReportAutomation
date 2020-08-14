using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace GenerarReporteSTAD
{
    public partial class Form1 : Form
    {
        ClMonitoreo monitoreo = new ClMonitoreo();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer.Start();
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            int Pre = monitoreo.Select_Prerequisitos();
            int Fin = monitoreo.Select_Finalizado();
            if ((Pre == 0) && (Fin > 0))
            {
                try
                {
                    System.Data.DataTable steps = monitoreo.Select_Next_Step();
                    if (steps.Rows.Count > 0)
                    {
                        if (steps.Rows[0]["Efectuado"].Equals(" "))
                        {
                            monitoreo.Update_Next_Step((int)steps.Rows[0]["Paso"], "-");
                            DateTime Fecha = (DateTime)steps.Rows[0]["FechaInicial"];
                            switch (steps.Rows[0]["Letra"].ToString().Trim())
                            {
                                case "A":
                                    #region Reporte
                                    try
                                    {
                                        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                                        if (xlApp == null)
                                        {
                                            //Console.WriteLine("No se pudo iniciar EXCEL");
                                        }
                                        xlApp.Visible = false;
                                        xlApp.DisplayAlerts = false;
                                        Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                                        Worksheet ws = (Worksheet)wb.Worksheets[1];
                                        if (ws == null)
                                        {
                                            //Console.WriteLine("No se pudo crear el Worksheet");
                                        }

                                        int CuentaFilas = 0, CuentaFilas2 = 0;
                                        System.Data.DataTable Datos = new System.Data.DataTable();

                                        ws.Name = "Transacciones";

                                        ws.Columns[2].ColumnWidth = 15;
                                        ws.Columns[3].ColumnWidth = 15;
                                        ws.Columns[4].ColumnWidth = 25;
                                        ws.Columns[5].ColumnWidth = 15;
                                        ws.Columns[6].ColumnWidth = 15;
                                        ws.Columns[7].ColumnWidth = 15;
                                        ws.Columns[8].ColumnWidth = 15;
                                        ws.Columns[9].ColumnWidth = 25;
                                        ws.Columns[10].ColumnWidth = 15;

                                        Datos = monitoreo.Select_STAD("Select * from Encabezado_TODOS_STAD");

                                        ws.Cells[2, 2].Font.Bold = true;
                                        ws.Cells[2, 2] = "Fecha:";

                                        for (int i = 0; i < Datos.Rows.Count; i++)
                                        {
                                            ws.Cells[2, 3] = Datos.Rows[i]["Fecha"].ToString();
                                        }

                                        #region DB13_SM51
                                        CuentaFilas = 4;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "DB13";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from DB13_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        CuentaFilas2 = 4;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SM51";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SM51_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SM59_SMLG
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SM59";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SM59_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SMLG";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SMLG_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region RZ10_SCC4
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "RZ10";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from RZ10_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SCC4";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SCC4_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region STMS_SE11_OLD
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "STMS";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from STMS_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SE11_OLD";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SE11_OLD_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SNOTE_SE14
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SNOTE";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SNOTE_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SE14";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SE14_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SE16_UASE16
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SE16N";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SE16_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "UASE16N";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from UASE16_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region LSMW_SU10
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "LSMW";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from LSMW_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SU10";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SU10_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SU01_SE38
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SU01";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SU01_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SE38";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SE38_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SM66_FS10N
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SM66";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SM66_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "FS10N";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from FS10N_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SE03_SM18
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SE03";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SE03_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SM18";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SM18_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region SM19_SM20
                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2] = "Transacción:";
                                        ws.Cells[CuentaFilas, 3] = "SM19";

                                        CuentaFilas = CuentaFilas + 2;

                                        ws.Cells[CuentaFilas, 2].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 2] = "Hora";
                                        ws.Cells[CuentaFilas, 3].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 3] = "Transacción";
                                        ws.Cells[CuentaFilas, 4].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 4] = "Programa";
                                        ws.Cells[CuentaFilas, 5].Font.Bold = true;
                                        ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas, 5] = "Usuario";

                                        CuentaFilas = CuentaFilas + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SM19_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 2] = "---------------";
                                            ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 3] = "---------------";
                                            ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 4] = "-------------------------------";
                                            ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas, 5] = "---------------";
                                            CuentaFilas = CuentaFilas + 2;
                                        }

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7] = "Transacción:";
                                        ws.Cells[CuentaFilas2, 8] = "SM20";

                                        CuentaFilas2 = CuentaFilas2 + 2;

                                        ws.Cells[CuentaFilas2, 7].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 7] = "Hora";
                                        ws.Cells[CuentaFilas2, 8].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 8] = "Transacción";
                                        ws.Cells[CuentaFilas2, 9].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 9] = "Programa";
                                        ws.Cells[CuentaFilas2, 10].Font.Bold = true;
                                        ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[CuentaFilas2, 10] = "Usuario";

                                        CuentaFilas2 = CuentaFilas2 + 1;

                                        Datos = monitoreo.Select_STAD("Select * from SM20_STAD Order by Hora");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
                                                ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
                                                ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
                                                ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
                                            }
                                            CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
                                        }
                                        else
                                        {
                                            ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 7] = "---------------";
                                            ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 8] = "---------------";
                                            ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 9] = "-------------------------------";
                                            ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[CuentaFilas2, 10] = "---------------";
                                            CuentaFilas2 = CuentaFilas2 + 2;
                                        }

                                        if (CuentaFilas > CuentaFilas2)
                                        {
                                            CuentaFilas2 = CuentaFilas;
                                        }
                                        else
                                        {
                                            CuentaFilas = CuentaFilas2;
                                        }
                                        #endregion

                                        #region Transacciones_SERVICE
                                        Datos = monitoreo.Select_STAD("Select * from Transacciones_SERVICE_STAD");
                                        if (Datos.Rows.Count > 0)
                                        {
                                            wb.Worksheets.Add();
                                            ws = (Worksheet)wb.Worksheets[1];

                                            ws.Name = "Transacciones SERVICE";

                                            ws.Columns[2].ColumnWidth = 30;
                                            ws.Columns[3].ColumnWidth = 35;
                                            ws.Columns[4].ColumnWidth = 10;
                                            ws.Columns[5].ColumnWidth = 30;

                                            ws.Cells[2, 2].Font.Bold = true;
                                            ws.Cells[2, 2] = "Usuario:";
                                            ws.Cells[3, 2].Font.Bold = true;
                                            ws.Cells[3, 2] = "Fecha:";
                                            ws.Cells[4, 2].Font.Bold = true;
                                            ws.Cells[4, 2] = "Horario:";

                                            Datos = monitoreo.Select_STAD("Select * from Encabezado_SERVICE_STAD");

                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
                                                ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                                ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
                                                ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
                                            }

                                            ws.Cells[6, 2].Font.Bold = true;
                                            ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 2] = "Código de la Transacción";
                                            ws.Cells[6, 3].Font.Bold = true;
                                            ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 3] = "Descripción de la Transacción";
                                            ws.Cells[6, 4].Font.Bold = true;
                                            ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 4] = "Módulo";
                                            ws.Cells[6, 5].Font.Bold = true;
                                            ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 5] = "Descripción del Módulo";

                                            Datos = monitoreo.Select_STAD("Select * from Transacciones_SERVICE_STAD");

                                            if (Datos.Rows.Count > 0)
                                            {
                                                for (int i = 0; i < Datos.Rows.Count; i++)
                                                {
                                                    ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
                                                    ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
                                                    ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
                                                    ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
                                                }
                                            }
                                            else
                                            {
                                                ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 2] = "-------------------------------------";
                                                ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 3] = "-----------------------------------------------";
                                                ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 4] = "-------------";
                                                ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 5] = "-----------------------------------------";
                                            }
                                        }
                                        #endregion

                                        #region Transacciones_AUDITOREXT
                                        Datos = monitoreo.Select_STAD("Select * from Transacciones_AUDITOREXT_STAD");
                                        if (Datos.Rows.Count > 0)
                                        {
                                            wb.Worksheets.Add();
                                            ws = (Worksheet)wb.Worksheets[1];

                                            ws.Name = "Transacciones AUDITOREXT";

                                            ws.Columns[2].ColumnWidth = 30;
                                            ws.Columns[3].ColumnWidth = 35;
                                            ws.Columns[4].ColumnWidth = 10;
                                            ws.Columns[5].ColumnWidth = 30;

                                            ws.Cells[2, 2].Font.Bold = true;
                                            ws.Cells[2, 2] = "Usuario:";
                                            ws.Cells[3, 2].Font.Bold = true;
                                            ws.Cells[3, 2] = "Fecha:";
                                            ws.Cells[4, 2].Font.Bold = true;
                                            ws.Cells[4, 2] = "Horario:";

                                            Datos = monitoreo.Select_STAD("Select * from Encabezado_AUDITOREXT_STAD");

                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
                                                ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                                ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
                                                ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
                                            }

                                            ws.Cells[6, 2].Font.Bold = true;
                                            ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 2] = "Código de la Transacción";
                                            ws.Cells[6, 3].Font.Bold = true;
                                            ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 3] = "Descripción de la Transacción";
                                            ws.Cells[6, 4].Font.Bold = true;
                                            ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 4] = "Módulo";
                                            ws.Cells[6, 5].Font.Bold = true;
                                            ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 5] = "Descripción del Módulo";

                                            Datos = monitoreo.Select_STAD("Select * from Transacciones_AUDITOREXT_STAD");

                                            if (Datos.Rows.Count > 0)
                                            {
                                                for (int i = 0; i < Datos.Rows.Count; i++)
                                                {
                                                    ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
                                                    ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
                                                    ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
                                                    ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
                                                }
                                            }
                                            else
                                            {
                                                ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 2] = "-------------------------------------";
                                                ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 3] = "-----------------------------------------------";
                                                ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 4] = "-------------";
                                                ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 5] = "-----------------------------------------";
                                            }
                                        }
                                        #endregion

                                        #region Transacciones_DDIC
                                        wb.Worksheets.Add();
                                        ws = (Worksheet)wb.Worksheets[1];

                                        ws.Name = "Transacciones DDIC";

                                        ws.Columns[2].ColumnWidth = 30;
                                        ws.Columns[3].ColumnWidth = 35;
                                        ws.Columns[4].ColumnWidth = 10;
                                        ws.Columns[5].ColumnWidth = 30;

                                        ws.Cells[2, 2].Font.Bold = true;
                                        ws.Cells[2, 2] = "Usuario:";
                                        ws.Cells[3, 2].Font.Bold = true;
                                        ws.Cells[3, 2] = "Fecha:";
                                        ws.Cells[4, 2].Font.Bold = true;
                                        ws.Cells[4, 2] = "Horario:";

                                        Datos = monitoreo.Select_STAD("Select * from Encabezado_DDIC_STAD");

                                        for (int i = 0; i < Datos.Rows.Count; i++)
                                        {
                                            ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
                                            ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                            ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
                                            ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
                                        }

                                        ws.Cells[6, 2].Font.Bold = true;
                                        ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 2] = "Código de la Transacción";
                                        ws.Cells[6, 3].Font.Bold = true;
                                        ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 3] = "Descripción de la Transacción";
                                        ws.Cells[6, 4].Font.Bold = true;
                                        ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 4] = "Módulo";
                                        ws.Cells[6, 5].Font.Bold = true;
                                        ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 5] = "Descripción del Módulo";

                                        Datos = monitoreo.Select_STAD("Select * from Transacciones_DDIC_STAD");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
                                                ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
                                                ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
                                                ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
                                            }
                                        }
                                        else
                                        {
                                            ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 2] = "-------------------------------------";
                                            ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 3] = "-----------------------------------------------";
                                            ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 4] = "-------------";
                                            ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 5] = "-----------------------------------------";
                                        }

                                        #endregion

                                        #region Transacciones_GRCLATAM
                                        wb.Worksheets.Add();
                                        ws = (Worksheet)wb.Worksheets[1];

                                        ws.Name = "Transacciones GRCLATAM";

                                        ws.Columns[2].ColumnWidth = 30;
                                        ws.Columns[3].ColumnWidth = 35;
                                        ws.Columns[4].ColumnWidth = 10;
                                        ws.Columns[5].ColumnWidth = 30;

                                        ws.Cells[2, 2].Font.Bold = true;
                                        ws.Cells[2, 2] = "Usuario:";
                                        ws.Cells[3, 2].Font.Bold = true;
                                        ws.Cells[3, 2] = "Fecha:";
                                        ws.Cells[4, 2].Font.Bold = true;
                                        ws.Cells[4, 2] = "Horario:";

                                        Datos = monitoreo.Select_STAD("Select * from Encabezado_GRCLATAM_STAD");

                                        for (int i = 0; i < Datos.Rows.Count; i++)
                                        {
                                            ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
                                            ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                            ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
                                            ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
                                        }

                                        ws.Cells[6, 2].Font.Bold = true;
                                        ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 2] = "Código de la Transacción";
                                        ws.Cells[6, 3].Font.Bold = true;
                                        ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 3] = "Descripción de la Transacción";
                                        ws.Cells[6, 4].Font.Bold = true;
                                        ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 4] = "Módulo";
                                        ws.Cells[6, 5].Font.Bold = true;
                                        ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 5] = "Descripción del Módulo";

                                        Datos = monitoreo.Select_STAD("Select * from Transacciones_GRCLATAM_STAD");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
                                                ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
                                                ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
                                                ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
                                            }
                                        }
                                        else
                                        {
                                            ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 2] = "-------------------------------------";
                                            ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 3] = "-----------------------------------------------";
                                            ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 4] = "-------------";
                                            ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 5] = "-----------------------------------------";
                                        }

                                        #endregion

                                        #region Transacciones_YNAVA
                                        wb.Worksheets.Add();
                                        ws = (Worksheet)wb.Worksheets[1];

                                        ws.Name = "Transacciones YNAVA";

                                        ws.Columns[2].ColumnWidth = 30;
                                        ws.Columns[3].ColumnWidth = 35;
                                        ws.Columns[4].ColumnWidth = 10;
                                        ws.Columns[5].ColumnWidth = 30;

                                        ws.Cells[2, 2].Font.Bold = true;
                                        ws.Cells[2, 2] = "Usuario:";
                                        ws.Cells[3, 2].Font.Bold = true;
                                        ws.Cells[3, 2] = "Fecha:";
                                        ws.Cells[4, 2].Font.Bold = true;
                                        ws.Cells[4, 2] = "Horario:";

                                        Datos = monitoreo.Select_STAD("Select * from Encabezado_YNAVA_STAD");

                                        for (int i = 0; i < Datos.Rows.Count; i++)
                                        {
                                            ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
                                            ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                            ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
                                            ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
                                        }

                                        ws.Cells[6, 2].Font.Bold = true;
                                        ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 2] = "Código de la Transacción";
                                        ws.Cells[6, 3].Font.Bold = true;
                                        ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 3] = "Descripción de la Transacción";
                                        ws.Cells[6, 4].Font.Bold = true;
                                        ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 4] = "Módulo";
                                        ws.Cells[6, 5].Font.Bold = true;
                                        ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
                                        ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                        ws.Cells[6, 5] = "Descripción del Módulo";

                                        Datos = monitoreo.Select_STAD("Select * from Transacciones_YNAVA_STAD");

                                        if (Datos.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
                                                ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
                                                ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
                                                ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
                                            }
                                        }
                                        else
                                        {
                                            ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 2] = "-------------------------------------";
                                            ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 3] = "-----------------------------------------------";
                                            ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 4] = "-------------";
                                            ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[7, 5] = "-----------------------------------------";
                                        }

                                        #endregion

                                        #region Transacciones_SERVICE1
                                        Datos = monitoreo.Select_STAD("Select * from Transacciones_SERVICE1_STAD");
                                        if (Datos.Rows.Count > 0)
                                        {
                                            wb.Worksheets.Add();
                                            ws = (Worksheet)wb.Worksheets[1];

                                            ws.Name = "Transacciones SERVICE1";

                                            ws.Columns[2].ColumnWidth = 30;
                                            ws.Columns[3].ColumnWidth = 35;
                                            ws.Columns[4].ColumnWidth = 10;
                                            ws.Columns[5].ColumnWidth = 30;

                                            ws.Cells[2, 2].Font.Bold = true;
                                            ws.Cells[2, 2] = "Usuario:";
                                            ws.Cells[3, 2].Font.Bold = true;
                                            ws.Cells[3, 2] = "Fecha:";
                                            ws.Cells[4, 2].Font.Bold = true;
                                            ws.Cells[4, 2] = "Horario:";

                                            Datos = monitoreo.Select_STAD("Select * from Encabezado_SERVICE1_STAD");

                                            for (int i = 0; i < Datos.Rows.Count; i++)
                                            {
                                                ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
                                                ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                                ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
                                                ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
                                            }

                                            ws.Cells[6, 2].Font.Bold = true;
                                            ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 2] = "Código de la Transacción";
                                            ws.Cells[6, 3].Font.Bold = true;
                                            ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 3] = "Descripción de la Transacción";
                                            ws.Cells[6, 4].Font.Bold = true;
                                            ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 4] = "Módulo";
                                            ws.Cells[6, 5].Font.Bold = true;
                                            ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
                                            ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
                                            ws.Cells[6, 5] = "Descripción del Módulo";

                                            Datos = monitoreo.Select_STAD("Select * from Transacciones_SERVICE1_STAD");

                                            if (Datos.Rows.Count > 0)
                                            {
                                                for (int i = 0; i < Datos.Rows.Count; i++)
                                                {
                                                    ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
                                                    ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
                                                    ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
                                                    ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                    ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
                                                }
                                            }
                                            else
                                            {
                                                ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 2] = "-------------------------------------";
                                                ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 3] = "-----------------------------------------------";
                                                ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 4] = "-------------";
                                                ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
                                                ws.Cells[7, 5] = "-----------------------------------------";
                                            }
                                        }
                                        #endregion

                                        wb.SaveAs(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\STAD.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                        wb.Close();
                                        xlApp.Quit();
                                        File.Copy(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\STAD.xlsx", @"\\servfile\stad\STAD.xlsx", true);
                                        //Console.WriteLine("Reporte generado exitosamente");
                                    }
                                    catch (Exception Exception) { MessageBox.Show(Exception.Message); }
                                    #endregion
                                    break;
                                case "B":
                                    try
                                    {
                                        StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\300" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", false, Encoding.UTF8);
                                        sw.WriteLine(" System: PRD       Client:  300     Number of RFCs which responded (without errors):   3 (  3)");
                                        sw.WriteLine(" Analysed time:  " + Fecha.Day.ToString().PadLeft(2, '0') + "." + Fecha.Month.ToString().PadLeft(2, '0') + "." + Fecha.Year.ToString().PadLeft(4, '0') + " / 00:00:00  -  " + Fecha.Day.ToString().PadLeft(2, '0') + "." + Fecha.Month.ToString().PadLeft(2, '0') + "." + Fecha.Year.ToString().PadLeft(4, '0') + " / 23:59:59");
                                        sw.WriteLine(" Display mode:   All statistic records, sorted by time                 Application statistic is included");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.WriteLine("|Started  Server           Transaction          Program                                  T Scr. Wp|User        |Response  |Time in   |Wait time |CPU time  |DB req.   |VMC elapsed|Memory    |Transfered|");
                                        sw.WriteLine("|                                               Function                                          |            |time (ms) | WPs (ms) |   (ms)   |   (ms)   |time (ms) |time (ms)  |used (kB) |kBytes    |");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.WriteLine("|                          *                    *                                        *        |*           |        0 |          |          |        0 |        0 |           |          |        0 |");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        System.Data.DataTable datos = monitoreo.Select_STAD("Select Fecha, Servidor, Transaccion, Programa, TPantalla, Pantalla, Wp, Usuario, TRespuestaMS, TWPMS, TEsperaMS, TCPUMS, TDBMS, TVMCMS, MemoriaKB, TransferidosKB from STAD Where Mandante = 300 Order by Fecha");
                                        for (int i = 0; i < datos.Rows.Count; i++)
                                        {
                                            DateTime Hora = (DateTime)datos.Rows[i]["Fecha"];
                                            int TRespuestaMS = (int)datos.Rows[i]["TRespuestaMS"];
                                            int TWPMS = (int)datos.Rows[i]["TWPMS"];
                                            int TEsperaMS = (int)datos.Rows[i]["TEsperaMS"];
                                            int TCPUMS = (int)datos.Rows[i]["TCPUMS"];
                                            int TDBMS = (int)datos.Rows[i]["TDBMS"];
                                            int TVMCMS = (int)datos.Rows[i]["TVMCMS"];
                                            int MemoriaKB = (int)datos.Rows[i]["MemoriaKB"];
                                            float TransferidosKB = (float)datos.Rows[i]["TransferidosKB"];
                                            sw.WriteLine("|" + Hora.Hour.ToString().PadLeft(2, '0') + ":" + Hora.Minute.ToString().PadLeft(2, '0') + ":" + Hora.Second.ToString().PadLeft(2, '0') + " " + datos.Rows[i]["Servidor"].ToString().PadRight(17, ' ') + datos.Rows[i]["Transaccion"].ToString().PadRight(21, ' ') + datos.Rows[i]["Programa"].ToString().PadRight(41, ' ') + datos.Rows[i]["TPantalla"].ToString().PadRight(2, ' ') + datos.Rows[i]["Pantalla"].ToString().PadLeft(4, '0').PadRight(5, ' ') + datos.Rows[i]["Wp"].ToString().PadRight(2, ' ') + "|" + datos.Rows[i]["Usuario"].ToString().PadRight(12, ' ') + "|" + TRespuestaMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TWPMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TEsperaMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TCPUMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TDBMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TVMCMS.ToString("N1").Replace(".0", "").PadLeft(10, ' ') + " |" + MemoriaKB.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TransferidosKB.ToString("N1").PadLeft(9, ' ') + " |");
                                        }
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.Close();
                                        File.Copy(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\300" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", @"\\servweb2\stad\300" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", true);
                                    }
                                    catch { }
                                    break;
                                case "C":
                                    try
                                    {
                                        StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\500" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", false, Encoding.UTF8);
                                        sw.WriteLine(" System: PRD       Client:  500     Number of RFCs which responded (without errors):   3 (  3)");
                                        sw.WriteLine(" Analysed time:  " + Fecha.Day.ToString().PadLeft(2, '0') + "." + Fecha.Month.ToString().PadLeft(2, '0') + "." + Fecha.Year.ToString().PadLeft(4, '0') + " / 00:00:00  -  " + Fecha.Day.ToString().PadLeft(2, '0') + "." + Fecha.Month.ToString().PadLeft(2, '0') + "." + Fecha.Year.ToString().PadLeft(4, '0') + " / 23:59:59");
                                        sw.WriteLine(" Display mode:   All statistic records, sorted by time                 Application statistic is included");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.WriteLine("|Started  Server           Transaction          Program                                  T Scr. Wp|User        |Response  |Time in   |Wait time |CPU time  |DB req.   |VMC elapsed|Memory    |Transfered|");
                                        sw.WriteLine("|                                               Function                                          |            |time (ms) | WPs (ms) |   (ms)   |   (ms)   |time (ms) |time (ms)  |used (kB) |kBytes    |");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.WriteLine("|                          *                    *                                        *        |*           |        0 |          |          |        0 |        0 |           |          |        0 |");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        System.Data.DataTable datos = monitoreo.Select_STAD("Select Fecha, Servidor, Transaccion, Programa, TPantalla, Pantalla, Wp, Usuario, TRespuestaMS, TWPMS, TEsperaMS, TCPUMS, TDBMS, TVMCMS, MemoriaKB, TransferidosKB from STAD Where Mandante = 500 Order by Fecha");
                                        for (int i = 0; i < datos.Rows.Count; i++)
                                        {
                                            DateTime Hora = (DateTime)datos.Rows[i]["Fecha"];
                                            int TRespuestaMS = (int)datos.Rows[i]["TRespuestaMS"];
                                            int TWPMS = (int)datos.Rows[i]["TWPMS"];
                                            int TEsperaMS = (int)datos.Rows[i]["TEsperaMS"];
                                            int TCPUMS = (int)datos.Rows[i]["TCPUMS"];
                                            int TDBMS = (int)datos.Rows[i]["TDBMS"];
                                            int TVMCMS = (int)datos.Rows[i]["TVMCMS"];
                                            int MemoriaKB = (int)datos.Rows[i]["MemoriaKB"];
                                            float TransferidosKB = (float)datos.Rows[i]["TransferidosKB"];
                                            sw.WriteLine("|" + Hora.Hour.ToString().PadLeft(2, '0') + ":" + Hora.Minute.ToString().PadLeft(2, '0') + ":" + Hora.Second.ToString().PadLeft(2, '0') + " " + datos.Rows[i]["Servidor"].ToString().PadRight(17, ' ') + datos.Rows[i]["Transaccion"].ToString().PadRight(21, ' ') + datos.Rows[i]["Programa"].ToString().PadRight(41, ' ') + datos.Rows[i]["TPantalla"].ToString().PadRight(2, ' ') + datos.Rows[i]["Pantalla"].ToString().PadLeft(4, '0').PadRight(5, ' ') + datos.Rows[i]["Wp"].ToString().PadRight(2, ' ') + "|" + datos.Rows[i]["Usuario"].ToString().PadRight(12, ' ') + "|" + TRespuestaMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TWPMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TEsperaMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TCPUMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TDBMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TVMCMS.ToString("N1").Replace(".0", "").PadLeft(10, ' ') + " |" + MemoriaKB.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TransferidosKB.ToString("N1").PadLeft(9, ' ') + " |");
                                        }
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.Close();
                                        File.Copy(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\500" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", @"\\servweb2\stad\500" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", true);
                                    }
                                    catch { }
                                    break;
                                case "D":
                                    try
                                    {
                                        StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\800" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", false, Encoding.UTF8);
                                        sw.WriteLine(" System: PRD       Client:  800     Number of RFCs which responded (without errors):   3 (  3)");
                                        sw.WriteLine(" Analysed time:  " + Fecha.Day.ToString().PadLeft(2, '0') + "." + Fecha.Month.ToString().PadLeft(2, '0') + "." + Fecha.Year.ToString().PadLeft(4, '0') + " / 00:00:00  -  " + Fecha.Day.ToString().PadLeft(2, '0') + "." + Fecha.Month.ToString().PadLeft(2, '0') + "." + Fecha.Year.ToString().PadLeft(4, '0') + " / 23:59:59");
                                        sw.WriteLine(" Display mode:   All statistic records, sorted by time                 Application statistic is included");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.WriteLine("|Started  Server           Transaction          Program                                  T Scr. Wp|User        |Response  |Time in   |Wait time |CPU time  |DB req.   |VMC elapsed|Memory    |Transfered|");
                                        sw.WriteLine("|                                               Function                                          |            |time (ms) | WPs (ms) |   (ms)   |   (ms)   |time (ms) |time (ms)  |used (kB) |kBytes    |");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.WriteLine("|                          *                    *                                        *        |*           |        0 |          |          |        0 |        0 |           |          |        0 |");
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        System.Data.DataTable datos = monitoreo.Select_STAD("Select Fecha, Servidor, Transaccion, Programa, TPantalla, Pantalla, Wp, Usuario, TRespuestaMS, TWPMS, TEsperaMS, TCPUMS, TDBMS, TVMCMS, MemoriaKB, TransferidosKB from STAD Where Mandante = 800 Order by Fecha");
                                        for (int i = 0; i < datos.Rows.Count; i++)
                                        {
                                            DateTime Hora = (DateTime)datos.Rows[i]["Fecha"];
                                            int TRespuestaMS = (int)datos.Rows[i]["TRespuestaMS"];
                                            int TWPMS = (int)datos.Rows[i]["TWPMS"];
                                            int TEsperaMS = (int)datos.Rows[i]["TEsperaMS"];
                                            int TCPUMS = (int)datos.Rows[i]["TCPUMS"];
                                            int TDBMS = (int)datos.Rows[i]["TDBMS"];
                                            int TVMCMS = (int)datos.Rows[i]["TVMCMS"];
                                            int MemoriaKB = (int)datos.Rows[i]["MemoriaKB"];
                                            float TransferidosKB = (float)datos.Rows[i]["TransferidosKB"];
                                            sw.WriteLine("|" + Hora.Hour.ToString().PadLeft(2, '0') + ":" + Hora.Minute.ToString().PadLeft(2, '0') + ":" + Hora.Second.ToString().PadLeft(2, '0') + " " + datos.Rows[i]["Servidor"].ToString().PadRight(17, ' ') + datos.Rows[i]["Transaccion"].ToString().PadRight(21, ' ') + datos.Rows[i]["Programa"].ToString().PadRight(41, ' ') + datos.Rows[i]["TPantalla"].ToString().PadRight(2, ' ') + datos.Rows[i]["Pantalla"].ToString().PadLeft(4, '0').PadRight(5, ' ') + datos.Rows[i]["Wp"].ToString().PadRight(2, ' ') + "|" + datos.Rows[i]["Usuario"].ToString().PadRight(12, ' ') + "|" + TRespuestaMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TWPMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TEsperaMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TCPUMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TDBMS.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TVMCMS.ToString("N1").Replace(".0", "").PadLeft(10, ' ') + " |" + MemoriaKB.ToString("N1").Replace(".0", "").PadLeft(9, ' ') + " |" + TransferidosKB.ToString("N1").PadLeft(9, ' ') + " |");
                                        }
                                        sw.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                                        sw.Close();
                                        File.Copy(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\800" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", @"\\servweb2\stad\800" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + ".txt", true);
                                    }
                                    catch { }
                                    break;
                            }
                            monitoreo.Update_Next_Step((int)steps.Rows[0]["Paso"], "X");
                        }
                    }
                }
                catch
                { }
            }
            else
            {
                if ((Pre == 0) && (Fin == 0))
                {
                    try
                    {
                        File.Copy(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\STMS.htm", @"\\servfile\stad\STMS.htm", true);
                    }
                    catch { }
                    this.Close();
                }
            }
            //DateTime Today = DateTime.Now;
            //switch(Today.DayOfWeek)
            //{
            //    case DayOfWeek.Monday:
            //    case DayOfWeek.Tuesday:
            //    case DayOfWeek.Wednesday:
            //    case DayOfWeek.Thursday:
            //    case DayOfWeek.Friday:
            //    case DayOfWeek.Saturday:
            //    case DayOfWeek.Sunday:
            //        TimeSpan Ahora = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, 0);
            //        TimeSpan HoraProgramada = new TimeSpan(9, 0, 0);
            //        if (Ahora == HoraProgramada)
            //        {
            //            DateTime Fecha = DateTime.Now;
            //            //#region Carga 00-12
            //            //Fecha = Fecha.AddDays(-1);
            //            ////if (Fecha.DayOfWeek == DayOfWeek.Monday)
            //            ////{
            //            ////    Fecha = Fecha.AddDays(-3);
            //            ////}
            //            ////else
            //            ////{
            //            ////    Fecha = Fecha.AddDays(-1);
            //            ////}
            //            ////Console.WriteLine("Iniciando carga para archivo STAD:");
            //            ////Console.WriteLine("Fecha:" + Fecha.Day.ToString() + "/" + Fecha.Month.ToString() + "/" + Fecha.Year.ToString());
            //            ////Console.WriteLine("Hora: 00:00:00 - 12:00:00");
            //            //try
            //            //{
            //            //    string Ruta = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\STAD-" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + "-00-12.txt";
            //            //    StreamReader sr = new StreamReader(Ruta, Encoding.UTF7);
            //            //    string linea = "";
            //            //    string fecha = "";
            //            //    ClMonitoreo monitoreo = new ClMonitoreo();
            //            //    monitoreo.Delete_STAD();
            //            //    sr.ReadLine();
            //            //    linea = sr.ReadLine().Trim();
            //            //    if (linea.StartsWith("Analysed time:"))
            //            //    {
            //            //        fecha = linea.Replace("Analysed time:", "").Substring(0, 13).Trim();
            //            //    }
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    while (!sr.EndOfStream)
            //            //    {

            //            //        linea = sr.ReadLine();
            //            //        switch (linea)
            //            //        {
            //            //            default:
            //            //                try
            //            //                {
            //            //                    linea = linea.Substring(1, linea.Length - 1);
            //            //                    if (linea.StartsWith("Start of"))
            //            //                    {
            //            //                        fecha = linea.Replace("Start of", "").Substring(0, linea.Length - 9).Trim();
            //            //                    }
            //            //                    else
            //            //                    {
            //            //                        string started = linea.Substring(0, 8).Trim();
            //            //                        linea = linea.Substring(8, linea.Length - 8);
            //            //                        string server = linea.Substring(0, 18).Trim();
            //            //                        linea = linea.Substring(18, linea.Length - 18);
            //            //                        string transaction = linea.Substring(0, 21).Trim();
            //            //                        linea = linea.Substring(21, linea.Length - 21);
            //            //                        string program = linea.Substring(0, 41).Trim();
            //            //                        linea = linea.Substring(41, linea.Length - 41);
            //            //                        string TScreen = linea.Substring(0, 7).Trim();
            //            //                        linea = linea.Substring(7, linea.Length - 7);
            //            //                        string WP = linea.Substring(0, 2).Trim();
            //            //                        linea = linea.Substring(2, linea.Length - 2);
            //            //                        string User = linea.Substring(0, 13).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(13, linea.Length - 13);
            //            //                        string ResponseTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string TimeInWPS = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string WaitTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string CPUTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string DBReqTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string VMCelapsed = linea.Substring(0, 12).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(12, linea.Length - 12);
            //            //                        string MemoryUsed = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string TransferedKBytes = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string Client = "300";
            //            //                        monitoreo.Insert_STAD(fecha, started, server, transaction, program, TScreen, WP, User, ResponseTime, TimeInWPS, WaitTime, CPUTime, DBReqTime, VMCelapsed, MemoryUsed, TransferedKBytes, Client);
            //            //                    }
            //            //                }
            //            //                catch { }
            //            //                break;
            //            //        }
            //            //    }
            //            //    sr.Close();
            //            //    //Console.WriteLine("Fin de Carga...");
            //            //}
            //            //catch
            //            //{
            //            //}
            //            //#endregion
            //            //#region Carga 12-24
            //            ////Console.WriteLine("Iniciando carga para archivo STAD:");
            //            ////Console.WriteLine("Fecha:" + Fecha.Day.ToString() + "/" + Fecha.Month.ToString() + "/" + Fecha.Year.ToString());
            //            ////Console.WriteLine("Hora: 12:00:01 - 11:59:59");
            //            //try
            //            //{
            //            //    string Ruta = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\STAD-" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + "-12-24.txt";
            //            //    StreamReader sr = new StreamReader(Ruta, Encoding.UTF7);
            //            //    string linea = "";
            //            //    string fecha = "";
            //            //    ClMonitoreo monitoreo = new ClMonitoreo();
            //            //    sr.ReadLine();
            //            //    linea = sr.ReadLine().Trim();
            //            //    if (linea.StartsWith("Analysed time:"))
            //            //    {
            //            //        fecha = linea.Replace("Analysed time:", "").Substring(0, 13).Trim();
            //            //    }
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    sr.ReadLine();
            //            //    while (!sr.EndOfStream)
            //            //    {

            //            //        linea = sr.ReadLine();
            //            //        switch (linea)
            //            //        {
            //            //            default:
            //            //                try
            //            //                {
            //            //                    linea = linea.Substring(1, linea.Length - 1);
            //            //                    if (linea.StartsWith("Start of"))
            //            //                    {
            //            //                        fecha = linea.Replace("Start of", "").Substring(0, linea.Length - 9).Trim();
            //            //                    }
            //            //                    else
            //            //                    {
            //            //                        string started = linea.Substring(0, 8).Trim();
            //            //                        linea = linea.Substring(8, linea.Length - 8);
            //            //                        string server = linea.Substring(0, 18).Trim();
            //            //                        linea = linea.Substring(18, linea.Length - 18);
            //            //                        string transaction = linea.Substring(0, 21).Trim();
            //            //                        linea = linea.Substring(21, linea.Length - 21);
            //            //                        string program = linea.Substring(0, 41).Trim();
            //            //                        linea = linea.Substring(41, linea.Length - 41);
            //            //                        string TScreen = linea.Substring(0, 7).Trim();
            //            //                        linea = linea.Substring(7, linea.Length - 7);
            //            //                        string WP = linea.Substring(0, 2).Trim();
            //            //                        linea = linea.Substring(2, linea.Length - 2);
            //            //                        string User = linea.Substring(0, 13).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(13, linea.Length - 13);
            //            //                        string ResponseTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string TimeInWPS = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string WaitTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string CPUTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string DBReqTime = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string VMCelapsed = linea.Substring(0, 12).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(12, linea.Length - 12);
            //            //                        string MemoryUsed = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string TransferedKBytes = linea.Substring(0, 11).Replace("|", "").Trim();
            //            //                        linea = linea.Substring(11, linea.Length - 11);
            //            //                        string Client = "300";
            //            //                        monitoreo.Insert_STAD(fecha, started, server, transaction, program, TScreen, WP, User, ResponseTime, TimeInWPS, WaitTime, CPUTime, DBReqTime, VMCelapsed, MemoryUsed, TransferedKBytes, Client);
            //            //                    }
            //            //                }
            //            //                catch { }
            //            //                break;
            //            //        }
            //            //    }
            //            //    sr.Close();
            //            //    //Console.WriteLine("Fin de carga...");
            //            //    //Console.WriteLine("");
            //            //    //Console.WriteLine("Generando Reporte");
            //            //}
            //            //catch
            //            //{
            //            //}
            //            //#endregion

            //            #region Reporte
            //            try
            //            {
            //                ClMonitoreo monitoreo = new ClMonitoreo();
            //                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //                if (xlApp == null)
            //                {
            //                    //Console.WriteLine("No se pudo iniciar EXCEL");
            //                }
            //                xlApp.Visible = false;
            //                xlApp.DisplayAlerts = false;
            //                Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            //                Worksheet ws = (Worksheet)wb.Worksheets[1];
            //                if (ws == null)
            //                {
            //                    //Console.WriteLine("No se pudo crear el Worksheet");
            //                }

            //                int CuentaFilas = 0, CuentaFilas2 = 0;
            //                System.Data.DataTable Datos = new System.Data.DataTable();

            //                ws.Name = "Transacciones";

            //                ws.Columns[2].ColumnWidth = 15;
            //                ws.Columns[3].ColumnWidth = 15;
            //                ws.Columns[4].ColumnWidth = 25;
            //                ws.Columns[5].ColumnWidth = 15;
            //                ws.Columns[6].ColumnWidth = 15;
            //                ws.Columns[7].ColumnWidth = 15;
            //                ws.Columns[8].ColumnWidth = 15;
            //                ws.Columns[9].ColumnWidth = 25;
            //                ws.Columns[10].ColumnWidth = 15;

            //                Datos = monitoreo.Select_STAD("Select * from Encabezado_TODOS_STAD");

            //                ws.Cells[2, 2].Font.Bold = true;
            //                ws.Cells[2, 2] = "Fecha:";

            //                for (int i = 0; i < Datos.Rows.Count; i++)
            //                {
            //                    ws.Cells[2, 3] = Datos.Rows[i]["Fecha"].ToString();
            //                }

            //                #region DB13_SM51
            //                CuentaFilas = 4;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "DB13";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from DB13_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                CuentaFilas2 = 4;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SM51";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SM51_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region SM59_SMLG
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "SM59";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SM59_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SMLG";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SMLG_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region RZ10_SCC4
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "RZ10";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from RZ10_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SCC4";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SCC4_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region STMS_SE11_OLD
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "STMS";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from STMS_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SE11_OLD";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SE11_OLD_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region SNOTE_SE14
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "SNOTE";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SNOTE_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SE14";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SE14_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region SE16_UASE16
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "SE16N";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SE16_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "UASE16N";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from UASE16_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region LSMW_SU10
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "LSMW";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from LSMW_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SU10";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SU10_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region SU01_SE38
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "SU01";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SU01_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "SE38";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SE38_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region SM66_FS10N
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "SM66";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SM66_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7] = "Transacción:";
            //                ws.Cells[CuentaFilas2, 8] = "FS10N";

            //                CuentaFilas2 = CuentaFilas2 + 2;

            //                ws.Cells[CuentaFilas2, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 7] = "Hora";
            //                ws.Cells[CuentaFilas2, 8].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 8].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 8].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 8] = "Transacción";
            //                ws.Cells[CuentaFilas2, 9].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 9].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 9].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 9] = "Programa";
            //                ws.Cells[CuentaFilas2, 10].Font.Bold = true;
            //                ws.Cells[CuentaFilas2, 10].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas2, 10].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas2, 10] = "Usuario";

            //                CuentaFilas2 = CuentaFilas2 + 1;

            //                Datos = monitoreo.Select_STAD("Select * from FS10N_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 7] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 8] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 9] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas2, 10] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas2 = CuentaFilas2 + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas2, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 7] = "---------------";
            //                    ws.Cells[CuentaFilas2, 8].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 8] = "---------------";
            //                    ws.Cells[CuentaFilas2, 9].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 9] = "-------------------------------";
            //                    ws.Cells[CuentaFilas2, 10].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas2, 10] = "---------------";
            //                    CuentaFilas2 = CuentaFilas2 + 2;
            //                }

            //                if (CuentaFilas > CuentaFilas2)
            //                {
            //                    CuentaFilas2 = CuentaFilas;
            //                }
            //                else
            //                {
            //                    CuentaFilas = CuentaFilas2;
            //                }
            //                #endregion

            //                #region SE03
            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2] = "Transacción:";
            //                ws.Cells[CuentaFilas, 3] = "SE03";

            //                CuentaFilas = CuentaFilas + 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Hora";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Transacción";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Programa";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select * from SE03_STAD Order by Hora");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Hora"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Transacción"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "-------------------------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }
            //                #endregion

            //                #region Transacciones_AUDITOREXT
            //                Datos = monitoreo.Select_STAD("Select * from Transacciones_AUDITOREXT_STAD");
            //                if (Datos.Rows.Count > 0)
            //                {
            //                    wb.Worksheets.Add();
            //                    ws = (Worksheet)wb.Worksheets[1];

            //                    ws.Name = "Transacciones AUDITOREXT";

            //                    ws.Columns[2].ColumnWidth = 30;
            //                    ws.Columns[3].ColumnWidth = 35;
            //                    ws.Columns[4].ColumnWidth = 10;
            //                    ws.Columns[5].ColumnWidth = 30;

            //                    ws.Cells[2, 2].Font.Bold = true;
            //                    ws.Cells[2, 2] = "Usuario:";
            //                    ws.Cells[3, 2].Font.Bold = true;
            //                    ws.Cells[3, 2] = "Fecha:";
            //                    ws.Cells[4, 2].Font.Bold = true;
            //                    ws.Cells[4, 2] = "Horario:";

            //                    Datos = monitoreo.Select_STAD("Select * from Encabezado_AUDITOREXT_STAD");

            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
            //                        ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            //                        ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
            //                        ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
            //                    }

            //                    ws.Cells[6, 2].Font.Bold = true;
            //                    ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
            //                    ws.Cells[6, 2] = "Código de la Transacción";
            //                    ws.Cells[6, 3].Font.Bold = true;
            //                    ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
            //                    ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[6, 3] = "Descripción de la Transacción";
            //                    ws.Cells[6, 4].Font.Bold = true;
            //                    ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
            //                    ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[6, 4] = "Módulo";
            //                    ws.Cells[6, 5].Font.Bold = true;
            //                    ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
            //                    ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[6, 5] = "Descripción del Módulo";

            //                    Datos = monitoreo.Select_STAD("Select * from Transacciones_AUDITOREXT_STAD");

            //                    if (Datos.Rows.Count > 0)
            //                    {
            //                        for (int i = 0; i < Datos.Rows.Count; i++)
            //                        {
            //                            ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                            ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
            //                            ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                            ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
            //                            ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                            ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
            //                            ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                            ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
            //                        }
            //                    }
            //                    else
            //                    {
            //                        ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[7, 2] = "-------------------------------------";
            //                        ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[7, 3] = "-----------------------------------------------";
            //                        ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[7, 4] = "-------------";
            //                        ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[7, 5] = "-----------------------------------------";
            //                    }
            //                }
            //                #endregion

            //                #region Transacciones_DDIC
            //                wb.Worksheets.Add();
            //                ws = (Worksheet)wb.Worksheets[1];

            //                ws.Name = "Transacciones DDIC";

            //                ws.Columns[2].ColumnWidth = 30;
            //                ws.Columns[3].ColumnWidth = 35;
            //                ws.Columns[4].ColumnWidth = 10;
            //                ws.Columns[5].ColumnWidth = 30;

            //                ws.Cells[2, 2].Font.Bold = true;
            //                ws.Cells[2, 2] = "Usuario:";
            //                ws.Cells[3, 2].Font.Bold = true;
            //                ws.Cells[3, 2] = "Fecha:";
            //                ws.Cells[4, 2].Font.Bold = true;
            //                ws.Cells[4, 2] = "Horario:";

            //                Datos = monitoreo.Select_STAD("Select * from Encabezado_DDIC_STAD");

            //                for (int i = 0; i < Datos.Rows.Count; i++)
            //                {
            //                    ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
            //                    ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            //                    ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
            //                    ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
            //                }

            //                ws.Cells[6, 2].Font.Bold = true;
            //                ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 2] = "Código de la Transacción";
            //                ws.Cells[6, 3].Font.Bold = true;
            //                ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 3] = "Descripción de la Transacción";
            //                ws.Cells[6, 4].Font.Bold = true;
            //                ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 4] = "Módulo";
            //                ws.Cells[6, 5].Font.Bold = true;
            //                ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 5] = "Descripción del Módulo";

            //                Datos = monitoreo.Select_STAD("Select * from Transacciones_DDIC_STAD");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
            //                        ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
            //                        ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
            //                        ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
            //                    }
            //                }
            //                else
            //                {
            //                    ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 2] = "-------------------------------------";
            //                    ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 3] = "-----------------------------------------------";
            //                    ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 4] = "-------------";
            //                    ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 5] = "-----------------------------------------";
            //                }

            //                #endregion

            //                #region Transacciones_GRCLATAM
            //                wb.Worksheets.Add();
            //                ws = (Worksheet)wb.Worksheets[1];

            //                ws.Name = "Transacciones GRCLATAM";

            //                ws.Columns[2].ColumnWidth = 30;
            //                ws.Columns[3].ColumnWidth = 35;
            //                ws.Columns[4].ColumnWidth = 10;
            //                ws.Columns[5].ColumnWidth = 30;

            //                ws.Cells[2, 2].Font.Bold = true;
            //                ws.Cells[2, 2] = "Usuario:";
            //                ws.Cells[3, 2].Font.Bold = true;
            //                ws.Cells[3, 2] = "Fecha:";
            //                ws.Cells[4, 2].Font.Bold = true;
            //                ws.Cells[4, 2] = "Horario:";

            //                Datos = monitoreo.Select_STAD("Select * from Encabezado_GRCLATAM_STAD");

            //                for (int i = 0; i < Datos.Rows.Count; i++)
            //                {
            //                    ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
            //                    ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            //                    ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
            //                    ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
            //                }

            //                ws.Cells[6, 2].Font.Bold = true;
            //                ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 2] = "Código de la Transacción";
            //                ws.Cells[6, 3].Font.Bold = true;
            //                ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 3] = "Descripción de la Transacción";
            //                ws.Cells[6, 4].Font.Bold = true;
            //                ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 4] = "Módulo";
            //                ws.Cells[6, 5].Font.Bold = true;
            //                ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 5] = "Descripción del Módulo";

            //                Datos = monitoreo.Select_STAD("Select * from Transacciones_GRCLATAM_STAD");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
            //                        ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
            //                        ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
            //                        ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
            //                    }
            //                }
            //                else
            //                {
            //                    ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 2] = "-------------------------------------";
            //                    ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 3] = "-----------------------------------------------";
            //                    ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 4] = "-------------";
            //                    ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 5] = "-----------------------------------------";
            //                }

            //                #endregion

            //                #region Transacciones_YNAVA
            //                wb.Worksheets.Add();
            //                ws = (Worksheet)wb.Worksheets[1];

            //                ws.Name = "Transacciones YNAVA";

            //                ws.Columns[2].ColumnWidth = 30;
            //                ws.Columns[3].ColumnWidth = 35;
            //                ws.Columns[4].ColumnWidth = 10;
            //                ws.Columns[5].ColumnWidth = 30;

            //                ws.Cells[2, 2].Font.Bold = true;
            //                ws.Cells[2, 2] = "Usuario:";
            //                ws.Cells[3, 2].Font.Bold = true;
            //                ws.Cells[3, 2] = "Fecha:";
            //                ws.Cells[4, 2].Font.Bold = true;
            //                ws.Cells[4, 2] = "Horario:";

            //                Datos = monitoreo.Select_STAD("Select * from Encabezado_YNAVA_STAD");

            //                for (int i = 0; i < Datos.Rows.Count; i++)
            //                {
            //                    ws.Cells[2, 3] = Datos.Rows[i]["Usuario"].ToString();
            //                    ws.Cells[3, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            //                    ws.Cells[3, 3] = Datos.Rows[i]["Fecha"].ToString();
            //                    ws.Cells[4, 3] = Datos.Rows[i]["Hora"].ToString();
            //                }

            //                ws.Cells[6, 2].Font.Bold = true;
            //                ws.Cells[6, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 2] = "Código de la Transacción";
            //                ws.Cells[6, 3].Font.Bold = true;
            //                ws.Cells[6, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 3] = "Descripción de la Transacción";
            //                ws.Cells[6, 4].Font.Bold = true;
            //                ws.Cells[6, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 4] = "Módulo";
            //                ws.Cells[6, 5].Font.Bold = true;
            //                ws.Cells[6, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[6, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[6, 5] = "Descripción del Módulo";

            //                Datos = monitoreo.Select_STAD("Select * from Transacciones_YNAVA_STAD");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + 7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 2] = Datos.Rows[i]["Transaccion"].ToString();
            //                        ws.Cells[i + 7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 3] = Datos.Rows[i]["Descripcion"].ToString();
            //                        ws.Cells[i + 7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 4] = Datos.Rows[i]["Modulo"].ToString();
            //                        ws.Cells[i + 7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + 7, 5] = Datos.Rows[i]["DescripcionModulo"].ToString();
            //                    }
            //                }
            //                else
            //                {
            //                    ws.Cells[7, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 2] = "-------------------------------------";
            //                    ws.Cells[7, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 3] = "-----------------------------------------------";
            //                    ws.Cells[7, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 4] = "-------------";
            //                    ws.Cells[7, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[7, 5] = "-----------------------------------------";
            //                }

            //                #endregion

            //                wb.SaveAs(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\STAD.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            //                wb.Close();
            //                xlApp.Quit();
            //                //Console.WriteLine("Reporte generado exitosamente");
            //            }
            //            catch { }
            //            #endregion
            //            #region Monitoreo
            //            try
            //            {
            //                ClMonitoreo monitoreo = new ClMonitoreo();
            //                ClUsoFirmas usofirmas = new ClUsoFirmas();
            //                monitoreo.Insert_Stad_Usuarios();
            //                System.Data.DataTable Usuarios = monitoreo.Select_STAD("Select Distinct Usuario from STAD_USUARIOS Where Nombre = ''");
            //                for (int i = 0; i < Usuarios.Rows.Count; i++)
            //                {
            //                    System.Data.DataTable DataUsuario = usofirmas.Select("Select Top 1 NombreTrabajador, Unidad, Departamento, Gerencia, Empresa from TiemposUsoSAP Where Usuario = '" + Usuarios.Rows[i]["Usuario"].ToString().Trim() + "'");
            //                    for (int j = 0; j < DataUsuario.Rows.Count; j++)
            //                    {
            //                        monitoreo.Update_Stad_Usuarios(DataUsuario.Rows[j]["NombreTrabajador"].ToString().Trim(), DataUsuario.Rows[j]["Unidad"].ToString().Trim(), DataUsuario.Rows[j]["Departamento"].ToString().Trim(), DataUsuario.Rows[j]["Gerencia"].ToString().Trim(), DataUsuario.Rows[j]["Empresa"].ToString().Trim(), Usuarios.Rows[i]["Usuario"].ToString().Trim());
            //                    }
            //                }
            //                Usuarios = monitoreo.Select_STAD("Select Distinct Usuario from STAD");
            //                for (int i = 0; i < Usuarios.Rows.Count; i++)
            //                {
            //                    //Comparar
            //                    monitoreo.Insert_Stad_Log(Usuarios.Rows[i]["Usuario"].ToString().Trim());
            //                }
            //                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //                if (xlApp == null)
            //                {
            //                    //Console.WriteLine("No se pudo iniciar EXCEL");
            //                }
            //                #region Proceso
            //                xlApp.Visible = false;
            //                xlApp.DisplayAlerts = false;
            //                Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            //                Worksheet ws = (Worksheet)wb.Worksheets[1];
            //                if (ws == null)
            //                {
            //                    //Console.WriteLine("No se pudo crear el Worksheet");
            //                }
            //                System.Data.DataTable Datos = new System.Data.DataTable();
            //                ws.Name = "Prohibido";

            //                int CuentaFilas = 2;

            //                ws.Cells[CuentaFilas, 2].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 2].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 2].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 2] = "Fecha";
            //                ws.Cells[CuentaFilas, 3].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 3].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 3].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 3] = "Servidor";
            //                ws.Cells[CuentaFilas, 4].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 4].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 4].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 4] = "Transaccion";
            //                ws.Cells[CuentaFilas, 5].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 5].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 5].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 5] = "Programa";
            //                ws.Cells[CuentaFilas, 6].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 6].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 6].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 6] = "TPantalla";
            //                ws.Cells[CuentaFilas, 7].Font.Bold = true;
            //                ws.Cells[CuentaFilas, 7].Interior.Color = XlRgbColor.rgbSteelBlue;
            //                ws.Cells[CuentaFilas, 7].Font.Color = XlRgbColor.rgbWhite;
            //                ws.Cells[CuentaFilas, 7] = "Usuario";

            //                CuentaFilas = CuentaFilas + 1;

            //                Datos = monitoreo.Select_STAD("Select Fecha, Servidor, Transaccion, Programa, TPantalla, Usuario from STAD_LOG Where Fecha Between '" + Fecha.Year.ToString() + Fecha.Month.ToString().PadLeft(2, '0') + Fecha.Day.ToString().PadLeft(2, '0') + " 00:00:00.000' And '" + Fecha.Year.ToString() + Fecha.Month.ToString().PadLeft(2, '0') + Fecha.Day.ToString().PadLeft(2, '0') + " 23:59:59.000'");

            //                if (Datos.Rows.Count > 0)
            //                {
            //                    for (int i = 0; i < Datos.Rows.Count; i++)
            //                    {
            //                        ws.Cells[i + CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 2] = Datos.Rows[i]["Fecha"].ToString();
            //                        ws.Cells[i + CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 3] = Datos.Rows[i]["Servidor"].ToString();
            //                        ws.Cells[i + CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 4] = Datos.Rows[i]["Transaccion"].ToString();
            //                        ws.Cells[i + CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 5] = Datos.Rows[i]["Programa"].ToString();
            //                        ws.Cells[i + CuentaFilas, 6].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 6] = Datos.Rows[i]["TPantalla"].ToString();
            //                        ws.Cells[i + CuentaFilas, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                        ws.Cells[i + CuentaFilas, 7] = Datos.Rows[i]["Usuario"].ToString();
            //                    }
            //                    CuentaFilas = CuentaFilas + Datos.Rows.Count + 1;
            //                }
            //                else
            //                {
            //                    ws.Cells[CuentaFilas, 2].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 2] = "---------------";
            //                    ws.Cells[CuentaFilas, 3].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 3] = "---------------";
            //                    ws.Cells[CuentaFilas, 4].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 4] = "---------------";
            //                    ws.Cells[CuentaFilas, 5].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 5] = "---------------";
            //                    ws.Cells[CuentaFilas, 6].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 6] = "---------------";
            //                    ws.Cells[CuentaFilas, 7].Borders.Color = XlRgbColor.rgbSteelBlue;
            //                    ws.Cells[CuentaFilas, 7] = "---------------";
            //                    CuentaFilas = CuentaFilas + 2;
            //                }

            //                wb.SaveAs(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\PROHIBIDO.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            //                wb.Close();
            //                xlApp.Quit();
            //                #endregion
            //            }
            //            catch
            //            { }
            //            #endregion

            //            //#region Copia Archivos
            //            //try 
            //            //{
            //            //    string Directorio = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\";
            //            //    string Archivo1 = "STAD-" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + "-00-12.txt";
            //            //    string Archivo2 = "STAD-" + Fecha.Year.ToString() + Fecha.Month.ToString().Trim().PadLeft(2, '0') + Fecha.Day.ToString().Trim().PadLeft(2, '0') + "-12-24.txt";
            //            //    File.Copy(Directorio + Archivo1, @"\\servfile\stad\" + Archivo1);
            //            //    File.Copy(Directorio + Archivo2, @"\\servfile\stad\" + Archivo2);
            //            //}
            //            //catch { }
            //            //#endregion
            //        }
            //        break;
            //    //case DayOfWeek.Saturday:
            //    //case DayOfWeek.Sunday:
            //    //    TimeSpan Ahorawe = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, 0);
            //    //    TimeSpan HoraProgramadawe = new TimeSpan(9, 0, 0);
            //    //    if (Ahorawe == HoraProgramadawe)
            //    //    {
            //    //        #region Copia Archivos Weekend
            //    //        try 
            //    //        {
            //    //            DateTime Fechawe = DateTime.Now;
            //    //            string Directoriowe = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\";
            //    //            string Archivo1we = "STAD-" + Fechawe.Year.ToString() + Fechawe.Month.ToString().Trim().PadLeft(2, '0') + Fechawe.Day.ToString().Trim().PadLeft(2, '0') + "-00-12.txt";
            //    //            string Archivo2we = "STAD-" + Fechawe.Year.ToString() + Fechawe.Month.ToString().Trim().PadLeft(2, '0') + Fechawe.Day.ToString().Trim().PadLeft(2, '0') + "-12-24.txt";
            //    //            File.Copy(Directoriowe + Archivo1we, @"\\servfile\stad\" + Archivo1we);
            //    //            File.Copy(Directoriowe + Archivo2we, @"\\servfile\stad\" + Archivo2we);
            //    //        }
            //    //        catch { }
            //    //        #endregion
            //    //    }
            //    //    break;
            //}
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.Stop();
        }
    }
}
