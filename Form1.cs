using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

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

        public Form1()
        {
            disciplines = new List<string>();
            years = new List<string>();
            destination = "";
            InitializeComponent();
        }


        private string get_file_name(DialogResult result)
        {
            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;

                return file;
            }
            return "";
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
                        dataTableTuteur = dataSet.Tables[0];
                    }


                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            destination = get_file_name(result);
            Console.WriteLine(destination);            
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


            foreach (string elt in disciplines)
            {
                Console.WriteLine("elt: " + elt);
            }
            foreach (string elt in years)
            {
                Console.WriteLine("elt: " + elt);
            }
        }


        

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataColumn dc in dataTableEtu.Columns)
            {
                Console.WriteLine("columnname: " + dc.ColumnName);
            }
            instantiate_years_disciplines_list();

            foreach(string year in years)
            {
                var rows_tuteur = dataTableTuteur.AsEnumerable().CopyToDataTable();
                var rows_etu = dataTableEtu.AsEnumerable().CopyToDataTable();
                DataRow[] results_tuteur = rows_tuteur.Select("année like '%" + year + "%'");
                DataRow[] results_etu = rows_etu.Select("année like '%" + year + "%'");
                
                foreach(DataRow row in results_tuteur)
                {
                    var rows_final = results_etu.AsEnumerable().CopyToDataTable();
                    DataRow[] final_results = rows_final.Select("discipline like '%" + row[id_discipline_tuteur] + "%'");

                    Console.WriteLine(String.Join(" : ", row.ItemArray) + "\n");
                    foreach(DataRow r in final_results)
                    {
                        Console.WriteLine(String.Join(" : ", r.ItemArray));
                    }
                    Console.WriteLine("STOOOOOOOOP\n\n\n\n\n");
                }
            }


            /**
            foreach(DataRow row in dataTableEtu.Select())
            {
                Console.WriteLine(row[id_year_etu]);
                DataRow[] results_tuteur = dataTableTuteur.Select("Année = "+row[id_year_etu].ToString());
                foreach(DataRow rows in results_tuteur)
                {
                    Console.WriteLine(rows[id_year_etu]);
                }
                
            }
            **/

            /**
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
            }**/
        }

    }
}
