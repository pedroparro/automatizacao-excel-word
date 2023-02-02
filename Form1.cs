using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace fv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //create word
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //carregar documento
            Microsoft.Office.Interop.Word.Document doc = null;
            Object fileName = @"C:\Users\pedro.parro\Desktop\Convert\MODELO.docx";
            Object missing = Type.Missing;
            for (int i = 0; i < 500; i++)
            {
                doc = app.Documents.Open(fileName, missing, missing);
                app.Selection.Find.ClearFormatting();
                app.Selection.Find.Replacement.ClearFormatting();

                //read file excel
                String[] tmp = new string[2];
                tmp = readExcel(i);

                

                //fill data
                app.Selection.Find.Execute("<Coluna-A>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[0], 2);
                app.Selection.Find.Execute("<Coluna-B>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[1], 2);
                app.Selection.Find.Execute("<Coluna-C>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[2], 2);
                app.Selection.Find.Execute("<Coluna-D>", missing, missing, missing, missing, missing, missing, missing, missing, formatString(tmp[3]), 2);
                app.Selection.Find.Execute("<Coluna-E>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[4], 2);
                app.Selection.Find.Execute("<Coluna-F>", missing, missing, missing, missing, missing, missing, missing, missing, uppercaseString(removerString(tmp[5])), 2);
                app.Selection.Find.Execute("<Coluna-G>", missing, missing, missing, missing, missing, missing, missing, missing, removeString(tmp[6]), 2);



                //save file
                object SaveAsFile = (object)@"C:\Users\pedro.parro\Desktop\Convert\" + tmp[0] + "-" + tmp[2] + ".docx";
                doc.SaveAs2(SaveAsFile, missing, missing, missing);
            }
            //mensagem
            MessageBox.Show("Arquivo criado com sucesso...", "MENSAGEM", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();

            doc.Close(false, missing, missing);
            app.Quit(false, false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        //REGEX
        private static string formatString(string number)
        {
            Regex regex = new Regex(@"[^\d]");
            number = regex.Replace(number, "");
            //string format = @"#######\-##\.####\.#\.##\.####";
            string format = @"0000000\-00\.0000\.0\.00\.0000";
            number = double.Parse(number).ToString(format);
            return number;
        }

        //REMOVE STRING GERAL
        private static string removeString(string str = "")
        {
            string result = str.Replace("(", "").Replace(")", "").Replace("GERAL", "").Replace("-", "").Replace("SP", "").Replace("BA", "").Replace("MG", "").Replace("PR", "").Replace("SC", "").Replace("RJ", "");
            return result;

        }

        //REMOVE STRING VARA
        private static string removerString(string str = ".")
        {
            string result = str.Replace(".", "");
            return result;

        }

        //UPPERCASE
        private static string uppercaseString(string upc)
        {
            return upc.ToUpper();
        }

        private string[] readExcel(int index)
        {
            string res = @"C:\Users\pedro.parro\Desktop\Convert\PLANILHA.xlsx";

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(res, 0, true, 5, "", "", true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            index += 2;

            string[] data = new string[7]; //array[id,nome,age]
            data[0] = xlWorkSheet.get_Range("A" + index.ToString()).Text; //number = TEXT
            data[1] = xlWorkSheet.get_Range("B" + index.ToString()).Value; //text = VALUE
            data[2] = xlWorkSheet.get_Range("C" + index.ToString()).Value;
            data[3] = xlWorkSheet.get_Range("D" + index.ToString()).Text;
            data[4] = xlWorkSheet.get_Range("E" + index.ToString()).Value;
            data[5] = xlWorkSheet.get_Range("F" + index.ToString()).Value;
            data[6] = xlWorkSheet.get_Range("G" + index.ToString()).Value;

            xlWorkBook.Close(false);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            //return data[]
            return data;

        }
    }
}
