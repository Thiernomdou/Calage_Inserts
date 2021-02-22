using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Calage_Inserts
{
    public partial class Retirer_un_insert : Form
    {
        
        
        

        public Retirer_un_insert()
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
            //Retirer un insert
            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);

            try
            {
                myConnection.Open();
                string strRequete = "DELETE FROM [ACI].[dbo].[Inserts] WHERE Numero = '" + textBox1.Text + "'";
                SqlCommand myCommand = new SqlCommand(strRequete, myConnection);
                SqlDataReader mySqDataReader = myCommand.ExecuteReader();

                if (textBox1.Text == "")
                {
                    MessageBox.Show("Veuillez entrez le numéro de l'insert");
                    
                }
                else
                {
                    MessageBox.Show("Vous avez supprimé l'insert " + textBox1.Text);
                    textBox1.Text = "";
                }
            }
            catch (Exception eMsg1)
            {
                Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
            }
            finally
            {
                myConnection.Close();
            }
        }
    }
}
