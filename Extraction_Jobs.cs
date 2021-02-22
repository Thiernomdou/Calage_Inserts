using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Oracle.DataAccess.Client;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Data.SQLite;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using Microsoft.VisualBasic;

namespace Calage_Inserts
{
    public partial class Extraction_Jobs : Form
    {

        string connString = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=fra2exa01-sxdir1-vip.europe.essilor.group)(PORT=1561)))
                             (CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=PUE1)));User Id=combo;Password=combo;";
        public Extraction_Jobs()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accueil r = new Accueil("");
            r.Show();
        }
        
        
        private void button2_Click(object sender, EventArgs e)
        {
            // création d'entrées TNS  
            
            

            /*
                        PUE1.WORLD =
             (DESCRIPTION =
               (ADDRESS_LIST =
                 (ADDRESS = (COMMUNITY = tcpip.world)(PROTOCOL = TCP)(Host = fra2exa01 - sxdir1 - vip.europe.essilor.group)(Port = 1561))

                        */

            OracleConnection conn = new OracleConnection(connString);
            try
            {
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = conn;
                cmd.CommandText = "Select * FROM COMBO.COMBO_JOB_LINES";
                OracleDataReader reader = cmd.ExecuteReader();

                dataGridView1.Rows.Clear();
                while (reader.Read())
                {
                    dataGridView1.Rows.Add(reader[1], reader[2], reader[3], reader[4], reader[5], reader[6], reader[7], reader[8], reader[9]);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
             
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }

        private void Extraction_Jobs_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        
        private void button3_Click(object sender, EventArgs e)
        {

            var shots_nb = 0;
            var cd_press = "";
            var cd_mold = "";
            var routing_name = "";

            OracleConnection conn = new OracleConnection(connString); ;
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT distinct (SHOTS_NB), CD_PRESS, CD_MOLD, ROUTING_NAME FROM COMBO_JOB_HEADER_TRACKING where JOB_NB = " + dataGridView1.CurrentRow.Cells["JOB_NB"].Value;
            OracleDataReader reader = cmd.ExecuteReader();
            while (reader.Read() == true)
            {
                shots_nb = reader.GetInt32(0);
                cd_press = reader.GetString(1);
                cd_mold = reader.GetString(2);
                routing_name = reader.GetString(3);
            }

            conn.Close();





            object misValue = System.Reflection.Missing.Value;

            Excel._Application xlsp = new Microsoft.Office.Interop.Excel.Application();

            if(xlsp == null)
            {
                MessageBox.Show("Excel n'est pas corectement installé");
            }

            Excel.Workbook xlworkbook = xlsp.Workbooks.Add(misValue);
            //Enregistrer les données dans la feuille 1
            Excel.Worksheet xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
            xlworkSheet.get_Range("A1", "I30").Borders.Weight = Excel.XlBorderWeight.xlThin;           

            xlworkSheet.Cells[1, 1] = "Date";
            xlworkSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/MM/dd");
            xlworkSheet.Cells[1, 5] = "Semaine";
            xlworkSheet.Cells[1, 6] = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            xlworkSheet.Cells[1, 7] = "Validation préparation";
            xlworkSheet.get_Range("A1", "I1").Font.Size = 10;
            xlworkSheet.get_Range("B1", "D1").Merge(false);
            xlworkSheet.get_Range("G1", "H1").Merge(false);

            xlworkSheet.Cells[2, 7] = "VALIDATION RECEPTION REGLEUR";
            xlworkSheet.get_Range("A2", "I2").Font.Bold = 5;
            xlworkSheet.get_Range("A2:A3", "F2:F3").Merge(false);
            xlworkSheet.get_Range("G3", "I3").Merge(false);

            xlworkSheet.get_Range("A4:A30", "I4:I30").Font.Bold = 5;
            xlworkSheet.Cells[4, 1] = "Job";
            xlworkSheet.Cells[4, 5] = cd_mold;
            xlworkSheet.Cells[4, 2] = dataGridView1.CurrentRow.Cells["JOB_NB"].Value;
            xlworkSheet.get_Range("B4", "C4").Merge(false);
            xlworkSheet.Cells[4, 4] = "Moule";
            xlworkSheet.get_Range("E4", "F4").Merge(false);
            xlworkSheet.Cells[4, 7] = "Presse";
            xlworkSheet.Cells[4, 8] = cd_press;
            xlworkSheet.get_Range("H4", "I4").Merge(false);

            xlworkSheet.get_Range("A5", "I5").Merge(false);

            xlworkSheet.Cells[6, 1] = "Produit";
            xlworkSheet.Cells[6, 2] = routing_name;
            xlworkSheet.get_Range("B6", "E6").Merge(false);
            xlworkSheet.Cells[6, 6] = "Num Shot";
            xlworkSheet.Cells[6, 7] = shots_nb;
            xlworkSheet.get_Range("G6", "I6").Merge(false);

            xlworkSheet.get_Range("A7", "I7").Merge(false);

            xlworkSheet.Cells[8, 1] = "Inserts convexes";
            xlworkSheet.get_Range("A8", "I8").Merge(false);

            xlworkSheet.get_Range("A9:A10", "I9:I10").Merge(false);

            xlworkSheet.Cells[111, 1] = "Cavités";
            xlworkSheet.get_Range("A11", "B11").Merge(false);
            xlworkSheet.get_Range("C11", "I11").NumberFormat = "@";
            xlworkSheet.Cells[11, 3] = "01";
            xlworkSheet.Cells[11, 4] = "02";
            xlworkSheet.Cells[11, 5] = "03";
            xlworkSheet.Cells[11, 6] = "04";
            xlworkSheet.Cells[11, 7] = "06";
            xlworkSheet.Cells[11, 8] = "07";
            xlworkSheet.Cells[11, 9] = "08";
            xlworkSheet.Cells[12, 1] = "Bases/Sphère";
            xlworkSheet.get_Range("A12", "B12").Merge(false);
            xlworkSheet.Cells[13, 1] = "Ins CC";
            xlworkSheet.get_Range("A13", "B13").Merge(false);
            xlworkSheet.Cells[14, 1] = "Epais. Centre";
            xlworkSheet.get_Range("A14", "B14").Merge(false);
            xlworkSheet.Cells[15, 1] = "Cales CC";
            xlworkSheet.get_Range("A15", "B15").Merge(false);

            xlworkSheet.get_Range("A16:A17", "I16:I17").Merge(false);

            xlworkSheet.Cells[18, 1] = "CYL";
            xlworkSheet.get_Range("A18", "B18").Merge(false);
            xlworkSheet.Cells[19, 1] = "ins CX";
            xlworkSheet.get_Range("A19", "B19").Merge(false);
            xlworkSheet.Cells[20, 1] = "Epais. Ctr. CX";
            xlworkSheet.get_Range("A20", "B20").Merge(false);
            xlworkSheet.Cells[21, 1] = "Epais verre";
            xlworkSheet.get_Range("A21", "B21").Merge(false);
            xlworkSheet.Cells[22, 1] = "Cales CX";
            xlworkSheet.get_Range("A22", "B22").Merge(false);

            xlworkSheet.get_Range("A23:A24", "I23:I24").Merge(false);

            xlworkSheet.Cells[25, 1] = "Changement d'inserts";
            xlworkSheet.get_Range("A25", "I25").Merge(false);
            xlworkSheet.Cells[26, 1] = "Cavités";
            xlworkSheet.get_Range("A26", "B26").Merge(false);
            xlworkSheet.get_Range("C26", "I26").NumberFormat = "@";
            xlworkSheet.Cells[26, 3] = "01";
            xlworkSheet.Cells[26, 4] = "02";
            xlworkSheet.Cells[26, 5] = "03";
            xlworkSheet.get_Range("A29", "B29").Merge(false);
            xlworkSheet.Cells[26, 6] = "04";
            xlworkSheet.Cells[26, 7] = "06";
            xlworkSheet.Cells[26, 8] = "07";
            xlworkSheet.Cells[26, 9] = "08";
            xlworkSheet.get_Range("A27", "B27").Merge(false);
            xlworkSheet.get_Range("A28", "B28").Merge(false);
            xlworkSheet.get_Range("A29", "B29").Merge(false);
            xlworkSheet.get_Range("A30", "B30").Merge(false);

            string destination = @"R:\COMMUN\ACI\Data\Jobs_Combo\Job_" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value;

            if (!File.Exists(destination))
            {
                xlworkbook.SaveAs(destination, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue);
                MessageBox.Show("Vous avez extrait le Job" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + " dans " + "R:\\COMMUN\\ACI\\Data\\Jobs_Combo");
                xlworkbook.Close(true, misValue, misValue);
                xlsp.Quit();
            }
            else
            {
                MessageBox.Show("Le Job" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + " existe déjà");

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string searchValue = textBox1.Text;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(searchValue))
                    {
                        row.Selected = true;
                        break;
                    }
                    
                }
                MessageBox.Show("Ce Job n'existe pas dans la liste");
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }
    }
}