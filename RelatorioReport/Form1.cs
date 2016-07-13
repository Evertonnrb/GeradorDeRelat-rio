using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// adicionando as dll's para trabalhar com o report
using Microsoft.Reporting.WebForms;
using System.IO;//pegar as pastas temp
using System.Diagnostics;//abrir pdf

namespace RelatorioReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void t_Click(object sender, EventArgs e)
        {
            //Implementando o botão gerar utilizando a classe para visualizar relatórios
            ReportViewer reportViewer = new ReportViewer();
            //processando com enum ProcessingMode
            reportViewer.ProcessingMode = ProcessingMode.Local;
            //especificando caminho para encontrar o relatório 
            reportViewer.LocalReport.ReportEmbeddedResource = "RelatorioReport.Relatorio.rdlc";
            //Parametros do relatorio
            List<ReportParameter> listReportParameter = new List<ReportParameter>();
            //Criar novos parametros para enviar !!!
            listReportParameter.Add(new ReportParameter("Nome", Nome.Text));
            reportViewer.LocalReport.SetParameters(listReportParameter);
            // Renderizando o relatório
            //output método render
            Warning[] warnings;
            string[] streamids;
            string mineType;
            string encoding;
            string extension;

            //add combo para esclolha de como deseja renderizar

            //pdf
            byte[] bytePDF = reportViewer.LocalReport.Render(
                "Pdf",null, out mineType, out encoding,out extension, out streamids, out warnings
                );
            FileStream fileStreamPDF = null;
            string nomeArquivoPdf = Path.GetTempPath() + "Relatório" + DateTime.Now.ToString("dd_MM_yyyy-HH_mm_ss") + ".pdf";
            fileStreamPDF = new FileStream(nomeArquivoPdf, FileMode.Create);
            fileStreamPDF.Write(bytePDF, 0, bytePDF.Length);
            fileStreamPDF.Close();
            Process.Start(nomeArquivoPdf);

            //Exel
            byte[] byteExcel = reportViewer.LocalReport.Render(
                "Excel", null, out mineType, out encoding, out extension, out streamids, out warnings
                );

            FileStream fileStreamExcel = null;

            string nomeArquivoExcel = Path.GetTempPath() + "Relatório" + DateTime.Now.ToString("dd_MM_yyyy-HH_mm_ss") + ".xls";
            fileStreamExcel = new FileStream(nomeArquivoExcel, FileMode.Create);
            fileStreamExcel.Write(byteExcel, 0, byteExcel.Length);
            fileStreamExcel.Close();
            Process.Start(nomeArquivoExcel);

        }
    }
}
