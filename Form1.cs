using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using OfficeOpenXml;
using OfficeOpenXml.Drawing.Style.ThreeD;

namespace AdaptEtu
{
    public partial class Form1 : Form
    {
        string destination;
        List<string> years;
        List<string> cursuss;
        DataTable dataTableEtu;
        DataTable dataTableTuteur;
        int id_year_etu;
        int id_cursus_etu;
        int id_year_tuteur;
        int id_cursus_tuteur;

        ExcelPackage excel;

        public Form1()
        {
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            cursuss = new List<string>();
            years = new List<string>();
            destination = "";
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = ofd.FileName;
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader;

                        reader = ExcelReaderFactory.CreateReader(stream);

                        //// reader.IsFirstRowAsColumnNames
                        var conf = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        };

                        var dataSet = reader.AsDataSet(conf);
                        
                        dataTableEtu = dataSet.Tables[0];
                        dataTableEtu.Columns.Add("role");
                        for (int i = 0; i <= dataTableEtu.Rows.Count-1; i++)
                        {
                            dataTableEtu.Rows[i]["role"] = "ETUDIANT";
                        }
                        reader.Close();
                    }


                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = ofd.FileName;
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader;

                        reader = ExcelReaderFactory.CreateReader(stream);

                        //// reader.IsFirstRowAsColumnNames
                        var conf = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        };

                        var dataSet = reader.AsDataSet(conf);

                        dataTableTuteur = dataSet.Tables[0];
                        Console.WriteLine(dataTableTuteur.Columns.Count);
                        dataTableTuteur.Columns.Add("role");
                        for (int i = 0; i <= dataTableTuteur.Rows.Count - 1; i++)
                        {
                            dataTableTuteur.Rows[i]["role"] = "TUTEUR";
                        }
                        reader.Close();
                    }


                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog result = new SaveFileDialog())
            {
                if (result.ShowDialog() == DialogResult.OK)
                {
                    txtFilename2.Text = result.FileName;
                    destination = result.FileName;

                    excel = new ExcelPackage();
                }
            }
        }

        public void instantiate_years_cursuss_list()
        {
            DataColumn cursus_etu = dataTableEtu.Columns["cursus"];
            DataColumn year_etu = dataTableEtu.Columns["Année"];
            if (year_etu == null) year_etu = dataTableEtu.Columns["Annee"];

            DataColumn cursus_tuteur = dataTableTuteur.Columns["cursus"];
            DataColumn year_tuteur = dataTableTuteur.Columns["Année"];
            if (year_tuteur == null) year_tuteur = dataTableTuteur.Columns["Annee"];


            id_year_etu = year_etu.Ordinal;
            id_cursus_etu = cursus_etu.Ordinal;
            id_year_tuteur = year_etu.Ordinal;
            id_cursus_tuteur = cursus_etu.Ordinal;


            for (int i = 0; i <= dataTableEtu.Rows.Count - 1; i++)
            {
                string dis = dataTableEtu.Rows[i][id_cursus_etu].ToString();
                if (cursuss.Where(item => item.Contains(dis)).FirstOrDefault() == null)
                    cursuss.Add(dis);
                string yea = dataTableEtu.Rows[i][id_year_etu].ToString();
                if (years.Where(item => item.Contains(yea)).FirstOrDefault() == null)
                    years.Add(yea);

            }

            for (int i = 0; i <= dataTableTuteur.Rows.Count - 1; i++)
            {
                string dis = dataTableTuteur.Rows[i][id_cursus_tuteur].ToString();
                if (cursuss.Where(item => item.Contains(dis)).FirstOrDefault() == null)
                    cursuss.Add(dis);

                string yea = dataTableTuteur.Rows[i][id_year_tuteur].ToString();
                if (years.Where(item => item.Contains(yea)).FirstOrDefault() == null)
                    years.Add(yea);
            }

        }

        private string ConvertObjectToString(object obj)
        {
            return obj?.ToString() ?? string.Empty;
        }

        private void create_new_excel(DataRow[] result, DataRow tuteur)
        {
            //DataTable dt = new DataTable(result[0][id_year_etu].ToString());

            if (result.Length != 0)
            {

                var res = result.AsEnumerable().CopyToDataTable();
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[result[0][id_year_etu].ToString()];
                if (worksheet == null)
                {
                    worksheet = excel.Workbook.Worksheets.Add(result[0][id_year_etu].ToString());
                }
                var headerRow = new List<string[]>()
                {
                    Array.ConvertAll<object, string>(tuteur.ItemArray, ConvertObjectToString)
               
                };

                int idx = 1;
                if(worksheet.Dimension != null) idx = worksheet.Dimension.End.Row+2;
                string hearderRange = "A" + idx + ":" + "F" + idx;
                worksheet.Cells[hearderRange].LoadFromArrays(headerRow);
                idx++;
                for (int i = 0; i <= result.Length - 1; i++)
                {
                    hearderRange = "A" + idx + ":" + "F" + idx;
                    headerRow = new List<string[]>()
                {
                   result[i].ItemArray.Cast<string>().ToArray()
                };
                    Console.WriteLine(String.Join(" : ", result[i].ItemArray) + "\n");
                    worksheet.Cells[hearderRange].LoadFromArrays(headerRow);
                    idx++;
                    
                }
                hearderRange = "A" + idx + ":" + "F" + idx;
                worksheet.Cells[hearderRange].LoadFromArrays(new List<string[]>());
                idx++;
                hearderRange = "A" + idx + ":" + "F" + idx;
                worksheet.Cells[hearderRange].LoadFromArrays(new List<string[]>());
                idx++;

                FileInfo excelFile = new FileInfo(@destination);
                excel.SaveAs(excelFile);
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            instantiate_years_cursuss_list();

            foreach (string year in years)
            {
                var rows_tuteur = dataTableTuteur.AsEnumerable().CopyToDataTable();
                var rows_etu = dataTableEtu.AsEnumerable().CopyToDataTable();

                DataRow[] results_tuteur = rows_tuteur.Select("année like '%" + year + "%'");
                DataRow[] results_etu = rows_etu.Select("année like '%" + year + "%'");

                foreach (DataRow row in results_tuteur)
                {
                    var rows_final = results_etu.AsEnumerable().CopyToDataTable();
                    DataRow[] final_results = rows_final.Select("cursus like '%" + row[id_cursus_tuteur] + "%'");

                    create_new_excel(final_results, row);

                }
            }
            Application.Exit();
        }

        
    }
}
