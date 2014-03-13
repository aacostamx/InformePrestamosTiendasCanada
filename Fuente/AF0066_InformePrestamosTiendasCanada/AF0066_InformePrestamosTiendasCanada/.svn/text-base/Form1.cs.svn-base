using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Globalization;
using System.Net;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//AUTOR:    Antonio Acosta Murillo                                                                          //
//OBJETIVO: Generar Informe de Prestamos Tiendas Canada                                                     //        
//FECHA:    27 Septiembre 2013                                                                              //
//PAIS:     Mexico                                                                                          //
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


namespace AF0066_InformePrestamosTiendasCanada
{
    public partial class Form1 : Form
    {
        SqlCommand cm = new SqlCommand();
        SqlConnection Cn = new SqlConnection();
        SqlDataAdapter Da = new SqlDataAdapter();
        CrystalDecisions.CrystalReports.Engine.ReportDocument Rpt1 = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        CrystalDecisions.CrystalReports.Engine.ReportDocument Rpt2 = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        ConexionSQL Con = new ConexionSQL();

        public Form1()
        {
            InitializeComponent();
        }

        private void GeneraInforme()
        {                
            try
            {
                Cn.ConnectionString = Con.LeeArchivo(@"C:\Sys\Exe\Conexion\ConexionControltiendas.txt");
                Cn.Open();
                cm.CommandTimeout = 0;
                cm.Connection = Cn;
                cm.CommandType = CommandType.Text;

                cm.CommandText = "Exec AF0066_InformePrestamosTiendasCanada";
                cm.ExecuteNonQuery();

                Cn.Close();

                //Reporte de Prestamos Tiendas Canada PDF
                Rpt1.FileName = @"rassdk://C:\sys\Crystal\AF0066_InformePrestamosTiendasCanadaPDF.rpt";
                Rpt1.DataSourceConnections[0].SetConnection(Con.IP, Con.DB, Con.USER, Con.PASS);
                Rpt1.Refresh();
                Rpt1.SetDatabaseLogon(Con.USER, Con.PASS);
                Rpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperLetter;
                Rpt1.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape;
                Rpt1.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "AF0066_InformePrestamosTiendasCanada" + ".pdf");
                Rpt1.Refresh();

                //Reporte de Prestamos Tiendas Canada EXCEL
                Rpt2.FileName = @"rassdk://C:\sys\Crystal\AF0066_InformePrestamosTiendasCanadaEXCEL.rpt";
                Rpt2.DataSourceConnections[0].SetConnection(Con.IP, Con.DB, Con.USER, Con.PASS);
                Rpt2.Refresh();
                Rpt2.SetDatabaseLogon(Con.USER, Con.PASS);
                Rpt2.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperLetter;
                Rpt2.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape;
                Rpt2.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, "AF0066_InformePrestamosTiendasCanada" + ".xls");
                Rpt2.Refresh();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error: " + Ex.Message.ToString() + "\nSource: " + Ex.Source.ToString() + "\nMetodo: " + Ex.TargetSite.ToString());
                Cn.Close();
            }
            
            label4.Text = DateTime.Now.ToLongTimeString();
            MessageBox.Show("Reporte generado", "Notificación");
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label3.Text = "";
            label4.Text = "";
            button1.Enabled = false;
            label3.Text = DateTime.Now.ToLongTimeString();
            
            GeneraInforme();
        }

        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            Ayuda obj = new Ayuda();
            this.Hide();
            obj.ShowDialog();
            this.Show();
        }

        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
    }
}