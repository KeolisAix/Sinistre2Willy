using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sinitres2Willy
{
    public partial class Form1 : Form
    {
        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string connstring = String.Format("Server={0};Port={1};" +
                    "User Id={2};Password={3};Database={4};",
                    "mdmkpa", "5432", "postgres",
                    "postgres", "vehicules_sinistre");
                NpgsqlConnection conn = new NpgsqlConnection(connstring);
                conn.Open();
                int all = 0;
                string sql;
                if (all == 1){
                    sql = "SELECT * FROM vehicules";
                }
                else { 
                    sql = "SELECT * FROM vehicules WHERE constat = '0' AND archive = '0'";
                }
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                dataGridView1.DataSource = dt;
                conn.Close();
                button2.PerformClick();
            }
            catch (Exception msg)
            {
                MessageBox.Show(msg.ToString());
                throw;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            SmtpClient Smtp_Server = new SmtpClient();
            MailMessage e_mail = new MailMessage();
            Smtp_Server.EnableSsl = true;
            Smtp_Server.UseDefaultCredentials = true;
            Smtp_Server.Credentials = new System.Net.NetworkCredential("pmaldi", "Ilete1fois");
            Smtp_Server.DeliveryMethod = SmtpDeliveryMethod.Network;
            Smtp_Server.Port = 25;
            Smtp_Server.Host = "webmail.keolis.com";

            e_mail = new MailMessage();
            e_mail.From = new MailAddress("patrice.maldi@keolis.com");
            e_mail.To.Add("patrice.maldi@keolis.com");
            e_mail.To.Add("sylvain.mennillo@keolis.com");
            e_mail.To.Add("willy.boisfer@keolis.com");
            e_mail.To.Add("alison.cavelier@keolis.com");
            //e_mail.To.Add("philippe.boulet@keolis.com");
            e_mail.Subject = "Sinistralité : " + System.DateTime.Now + "";
            e_mail.IsBodyHtml = true;
            if (dataGridView1.Rows.Count != 0)
            {
                sb.AppendLine("<center><img src='K:\\Commun\\Keolis_logo.png'><br><br>Date du rapport : <b>" + System.DateTime.Now + "</b>");
                sb.AppendLine("<center><b>Liste des sinistres sans Constat & Numéro Keorisk : </b></center><br>");
                sb.AppendLine("<style type='text/css'>table.border_1_solid_black td { border : 1px solid black; }</style>");
                sb.AppendLine("<br><br><table class='border_1_solid_black'><tr><td align='center' valign='middle'>Numero du Sinitre :</td><td align='center' valign='middle'>Numero de Bus :</td><td align='center' valign='middle'>Motif :</td><td align='center' valign='middle'>Controleur :</td><td align='center' valign='middle'>Date :</td><td align='center' valign='middle'>Heure :</td></tr>");
                foreach (DataGridViewRow rowView in dataGridView1.Rows)
                {
                    string num = rowView.Cells["id"].Value.ToString();
                    string bus = rowView.Cells["vehicule"].Value.ToString();
                    string motif = rowView.Cells["motif"].Value.ToString();
                    string controleur = rowView.Cells["controleur"].Value.ToString();
                    string date = rowView.Cells["date"].Value.ToString();
                    string Heure = rowView.Cells["Heure"].Value.ToString();
                    byte[] data = Convert.FromBase64String(motif);

                    string connstring = String.Format("Server={0};Port={1};" +
                    "User Id={2};Password={3};Database={4};",
                    "mdmkpa", "5432", "postgres",
                    "postgres", "vehicules");
                    NpgsqlConnection conn = new NpgsqlConnection(connstring);
                    conn.Open();
                    string sql = "SELECT modele FROM vehicules.vehicule WHERE parc_keolis = '" + bus + "'";
                    NpgsqlCommand command = new NpgsqlCommand(sql, conn);
                    string modele = command.ExecuteScalar().ToString();
                    string decodedString = Encoding.UTF8.GetString(data);
                    string decodedStringHTML = decodedString.Replace("%20", " ").Replace("%E9", "é").Replace("%2C", ",").Replace("%29", ")").Replace("%28", "(");
                    sb.AppendLine("<tr><td align='center' valign='middle'>" + num + "</td><td align='center' valign='middle'><a href='http://mdmkpa/sinistre/index.php?bus="+bus+"&modele="+modele+"&controleur=Willy%20Boisfer'>" + bus + "</a></td><td align='center' valign='middle'>" + decodedStringHTML + "</td><td align='center' valign='middle'>" + controleur + "</td><td align='center' valign='middle'>" + date + "</td><td align='center' valign='middle'>" + Heure + "</td></tr>");
                }
                sb.AppendLine("</table><br><br><i>Note : Attention le nom du contrôleur n'est pas encore validé, par conséquent le contrôleur par défaut est Willy BOISFER.");
            }
            else
            {
                sb.AppendLine("<center><img src='K:\\Commun\\Keolis_logo.png'><br><br>Date du rapport : <b>" + System.DateTime.Now + "</b>");
                sb.AppendLine("<br><br><h1><center><b>Il n'y a pas de nouveau Sinistre !</b></center><br><br>");
                sb.AppendLine("<center><img src='K:\\Commun\\pouce.png'>");
            }                
            e_mail.Body = sb.ToString();
            Smtp_Server.Send(e_mail);
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1.PerformClick();
        }
    }
}
