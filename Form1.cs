using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace AdaptEtu
{
    public partial class Form1 : Form
    {
        string source;
        string destination;
        Microsoft.Office.Interop.Excel.Workbook wb;
        Microsoft.Office.Interop.Excel.Worksheet ws;
        public Form1()
        {
            source = "";
            destination = "";
            InitializeComponent();
        }


        private string get_file_name(DialogResult result)
        {
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;

                Excel excel = new Excel(file, 1);

                return file;
            }
            return "";
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            source = get_file_name(result);
            Console.WriteLine(source);            
        }


        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = ofd.FileName;
                    using (var stream = System.IO.File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader;

                        reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                        //// reader.IsFirstRowAsColumnNames
                        var conf = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        };

                        var dataSet = reader.AsDataSet(conf);

                        // Now you can get data from each sheet by its index or its "name"
                        var dataTableEtu = dataSet.Tables[0];
                        var dataTableTuteur = dataSet.Tables[1];

                        var cellStr = "AB2"; // var cellStr = "A1";
                        var match = Regex.Match(cellStr, @"(?<col>[A-Z]+)(?<row>\d+)");
                        var colStr = match.Groups["col"].ToString();
                        var col = colStr.Select((t, i) => (colStr[i] - 64) * Math.Pow(26, colStr.Length - i - 1)).Sum();
                        var row = int.Parse(match.Groups["row"].ToString());

                        for (var i = 0; i < dataTableEtu.Rows.Count; i++)
                        {
                            for (var j = 0; j < dataTableEtu.Columns.Count; j++)
                            {
                                var data = dataTableEtu.Rows[i][j];
                                Console.WriteLine("data: i: "+i+" j: "+j+" "+ data);
                            }
                        }

                        for (var i = 0; i < dataTableTuteur.Rows.Count; i++)
                        {
                            for (var j = 0; j < dataTableTuteur.Columns.Count; j++)
                            {
                                var data = dataTableTuteur.Rows[i][j];
                                Console.WriteLine("data Tuteur: i: " + i + " j: " + j +" "+ data);
                            }
                        }
                    }


                }
            }
        }

    }
}
