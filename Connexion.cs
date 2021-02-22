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
    public partial class Connexion : Form
    {
        
       
        public Connexion()
        {
            InitializeComponent();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        

       

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox1.BackColor = Color.White;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //Connexion utilisateurs
            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);
           
            try
            {
                //conditions de remplissage du champ login
                if (textBox2.Text.Trim() == "" && textBox1.Text.Trim() == "")
                {
                    MessageBox.Show("Veuillez renseigner le champ login et password");
                    textBox2.BackColor = Color.Red;
                    textBox1.BackColor = Color.Red;
                    return;
                }
                else
                {
                    myConnection.Open();
                    string strRequete = "SELECT * FROM [ACI].[dbo].[User_ACI] WHERE [Login_User_ACI] = @user AND [Mdp] = @pass";
                    SqlCommand myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataAdapter da = new SqlDataAdapter(myCommand);

                    
                    myCommand.Parameters.AddWithValue("@user", textBox1.Text);
                    myCommand.Parameters.AddWithValue("@pass", textBox2.Text);
                    DataTable dt = new DataTable();
                    da.Fill(dt);


                    if (dt.Rows.Count > 0)
                    {
                        this.Hide();
                        if (dt.Rows[0][0].ToString() == "1")
                        {
                            Accueil a = new Accueil(dt.Rows[0][1].ToString());
                            a.Show();
                        }
                        else if (dt.Rows[0][0].ToString() == "2")
                        {
                            Accueil1 a = new Accueil1(dt.Rows[0][1].ToString());
                            a.Show();
                        }
                        else if (dt.Rows[0][0].ToString() == "3")
                        {
                            Accueil2 a = new Accueil2(dt.Rows[0][1].ToString());
                            a.Show();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Login ou Password incorrect");
                        textBox1.BackColor = Color.Red;
                        textBox2.BackColor = Color.Red;
                        return;
                    }
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

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            textBox2.BackColor = Color.White;
        }
    }
}
