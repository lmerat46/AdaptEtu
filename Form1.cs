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
        List<string> disciplines;
        DataTable dataTableEtu;
        DataTable dataTableTuteur;
        int id_year_etu;
        int id_discipline_etu;
        int id_year_tuteur;
        int id_discipline_tuteur;

        ExcelPackage excel;

        public Form1()
        {
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            disciplines = new List<string>();
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

                        // Now you can get data from each sheet by its index or its "name"
                        
                        dataTableEtu = dataSet.Tables[1];
                        dataTableEtu.Columns.Add("role");
                        for (int i = 0; i <= dataTableEtu.Rows.Count-1; i++)
                        {
                            dataTableEtu.Rows[i]["role"] = "ETUDIANT";
                        }
                        
                        dataTableTuteur = dataSet.Tables[0];
                        dataTableTuteur.Columns.Add("role");
                        for (int i = 0; i <= dataTableTuteur.Rows.Count - 1; i++)
                        {
                            dataTableTuteur.Rows[i]["role"] = "TUTEUR";
                        }
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
                    destination = result.FileName;

                    excel = new ExcelPackage();
                }
            }
        }

        public void instantiate_years_disciplines_list()
        {
            DataColumn discipline_etu = dataTableEtu.Columns["Discipline"];
            DataColumn year_etu = dataTableEtu.Columns["Année"];
            if (year_etu == null) year_etu = dataTableEtu.Columns["Annee"];

            DataColumn discipline_tuteur = dataTableTuteur.Columns["Discipline"];
            DataColumn year_tuteur = dataTableTuteur.Columns["Année"];
            if (year_tuteur == null) year_tuteur = dataTableTuteur.Columns["Annee"];


            id_year_etu = year_etu.Ordinal;
            id_discipline_etu = discipline_etu.Ordinal;
            id_year_tuteur = year_etu.Ordinal;
            id_discipline_tuteur = discipline_etu.Ordinal;


            for (int i = 0; i <= dataTableEtu.Rows.Count - 1; i++)
            {
                string dis = dataTableEtu.Rows[i][id_discipline_etu].ToString();
                if (disciplines.Where(item => item.Contains(dis)).FirstOrDefault() == null)
                    disciplines.Add(dis);
                string yea = dataTableEtu.Rows[i][id_year_etu].ToString();
                if (years.Where(item => item.Contains(yea)).FirstOrDefault() == null)
                    years.Add(yea);

            }

            for (int i = 0; i <= dataTableTuteur.Rows.Count - 1; i++)
            {
                string dis = dataTableTuteur.Rows[i][id_discipline_tuteur].ToString();
                if (disciplines.Where(item => item.Contains(dis)).FirstOrDefault() == null)
                    disciplines.Add(dis);

                string yea = dataTableTuteur.Rows[i][id_year_tuteur].ToString();
                if (years.Where(item => item.Contains(yea)).FirstOrDefault() == null)
                    years.Add(yea);
            }

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
                   tuteur.ItemArray.Cast<string>().ToArray()
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
            foreach (DataColumn dc in dataTableEtu.Columns)
            {
                Console.WriteLine("columnname: " + dc.ColumnName);
            }
            instantiate_years_disciplines_list();

            foreach (string year in years)
            {
                var rows_tuteur = dataTableTuteur.AsEnumerable().CopyToDataTable();
                var rows_etu = dataTableEtu.AsEnumerable().CopyToDataTable();
                DataRow[] results_tuteur = rows_tuteur.Select("année like '%" + year + "%'");
                DataRow[] results_etu = rows_etu.Select("année like '%" + year + "%'");

                foreach (DataRow row in results_tuteur)
                {
                    var rows_final = results_etu.AsEnumerable().CopyToDataTable();
                    DataRow[] final_results = rows_final.Select("discipline like '%" + row[id_discipline_tuteur] + "%'");

                    create_new_excel(final_results, row);

                }
            }
        }

    }
}
