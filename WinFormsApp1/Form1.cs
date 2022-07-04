using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Reflection;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public static bool Bearbeiten_Mode = true;
        public static bool Druck_Mode = false;
        public static int letzte_Zeile = 1;
        public int Schriftgröße = -1;
        public static string z1, z2, z3, z4, z5, z6, h1, h2, h3, h4, h5, h6, 
                             s1, s2, s3, s4, s5, s6, stzNr, BdNr, brte, AVA, srft, frb, gedd, Dtm;
        Dictionary<TextBox, string> GoNull = new Dictionary<TextBox, string> { };
        public string AktuellDruck;
        public string AktuellDatei = "Druckdatei";


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;
            BearbeitenMode();
            UngedruckteZeilenBerechnen();
            druckenToolStripMenuItem.Enabled = false;
            automSuchenToolStripMenuItem.Enabled = false;
            referenzfahrtToolStripMenuItem.Enabled = false;
            schriftplattWechselnToolStripMenuItem.Enabled = false;
            dauersuchenToolStripMenuItem.Enabled = false;
            AktuellDatei = "Druckdatei";
            this.ActiveControl = SchriftGröße;
            label3.Text = AktuellDatei;
        } 

        // Wechseln zwischen Druckmode und Bearbeitenmode
        private void FensterWechseln_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode == true)
            {
                Bearbeiten_Mode = false;
                Druck_Mode = true;
                DruckMode();
            }
            else if(Bearbeiten_Mode == false)
            {
                Bearbeiten_Mode = true;
                Druck_Mode = false;
                BearbeitenMode();
            }
        }

        // Zum Druckmode wechslen
        private void DruckMode ()
        {
            this.BackColor = Color.FromArgb(38, 87, 166);
            StopButton.Visible = true;
            AutoSuchen.Visible = true;
            AutoSuchen.Text = "Autom. suchen";
            SchriftWechseln.Visible = true;
            ReferenzFahrt.Visible = true;
            FensterWechseln.Text = "Textfenster";
            TextSpeichernOrDrucken.Text = "Drucken";
            panel7.Enabled = false;
            panel8.Enabled = false;
            druckenToolStripMenuItem.Enabled = true;
            automSuchenToolStripMenuItem.Enabled = true;
            referenzfahrtToolStripMenuItem.Enabled = true;
            schriftplattWechselnToolStripMenuItem.Enabled = true;
            dauersuchenToolStripMenuItem.Enabled = true;
        }

        // Zum Druckmode Bearbeitenmode
        private void BearbeitenMode()
        {
            this.BackColor = SystemColors.Control;
            StopButton.Visible = false;
            AutoSuchen.Visible = false;
            AutoSuchen.Text = "Titel ändern";
            SchriftWechseln.Visible = false;
            ReferenzFahrt.Visible = false;
            FensterWechseln.Text = "Druckfenster";
            TextSpeichernOrDrucken.Text = "Text speichern";
            panel7.Enabled = true;
            panel8.Enabled = true;
            druckenToolStripMenuItem.Enabled = false;
            automSuchenToolStripMenuItem.Enabled = false;
            referenzfahrtToolStripMenuItem.Enabled = false;
            schriftplattWechselnToolStripMenuItem.Enabled = false;
            dauersuchenToolStripMenuItem.Enabled = false;
        }

        // Wenn man auf eine Zeile ist und auf Enter kliclt, dann fokusieren auf nächsten Zeile 
        private void Zeile1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile2.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile3.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile4.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile5.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile6.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        // Überprüfen welche Zeile die aktuelle ist, um aus "Vorschläge" einen Vorschlag zu bekommen
        private void Zeile1_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 1;
        }

        private void Zeile2_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 2;
        }

        private void Zeile3_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 3;
        }

        private void Zeile4_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 4;
        }

        private void Zeile5_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 5;
        }

        private void Zeile6_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 6;
        }

        // Vorschläge zu den Zeilen schicken + Auf eine Zeile ohne Schriftgröße zu schreiben verhindern
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (letzte_Zeile == 1 && SchriftGröße.SelectedIndex != -1)
            {
                Zeile1.Text = Vorschläge.SelectedItem.ToString();
            }
            else if (letzte_Zeile == 2)
            {
                Zeile2.Text = Vorschläge.SelectedItem.ToString();
            }
            else if (letzte_Zeile == 3)
            {
                Zeile3.Text = Vorschläge.SelectedItem.ToString();
            }
            else if (letzte_Zeile == 4)
            {
                Zeile4.Text = Vorschläge.SelectedItem.ToString();
            }
            else if (letzte_Zeile == 5)
            {
                Zeile5.Text = Vorschläge.SelectedItem.ToString();
            }
            else if (letzte_Zeile == 6)
            {
                Zeile6.Text = Vorschläge.SelectedItem.ToString();
            }
        }

        // Nur Nummer in "Länge" erlauben
        private void LaengeZeile1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void LaengeZeile2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void LaengeZeile3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void LaengeZeile4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void LaengeZeile5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void LaengeZeile6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        // Nur Nummer in "Höhe" erlauben
        private void HoeheZeile1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void HoeheZeile2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void HoeheZeile3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void HoeheZeile4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void HoeheZeile5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void HoeheZeile6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        // Nur Nummer in "Sperren" erlauben
        private void TextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TextBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TextBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TextBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TextBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TextBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        // Das Programm verlassen
        private void ProgrammBeendenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        
        // Hot keys
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F12)
            {
                druckSuchenToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F7)
            {
                drToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.S && e.Control == true)
            {
                druckSpeichernToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.Y && e.Control == true)
            {
                druckLöschenToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F8)
            {
                druckNeuToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F6 && e.Control == true)
            {
                archivdruckToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.A && e.Control == true)
            {
                archivdruckSpeichernToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F7 && e.Control == true)
            {
                letzterDruckArchivToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.D && e.Control == true)
            {
                ausdruckToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F1)
            {
                druckenToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F1 && e.Control == true)
            {
                seite2ToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F2)
            {
                automSuchenToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F3)
            {
                textDruckfensterToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F4)
            {
                referenzfahrtToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F5)
            {
                schriftplattWechselnToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F9)
            {
                bandbreiteToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F11)
            {
                teildruckToolStripMenuItem.PerformClick();
            }

            if (e.KeyCode == Keys.F2 && e.Control == true)
            {
                dauersuchenToolStripMenuItem.PerformClick();
            }

        } 

        // Ändern "Schritgröße"
        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Schriftgröße = SchriftGröße.SelectedIndex + 1;
            SchritGrößeÄndernFürLänge(Zeile1, LaengeZeile1);
            SchritGrößeÄndernFürLänge(Zeile2, LaengeZeile2);
            SchritGrößeÄndernFürLänge(Zeile3, LaengeZeile3);
            SchritGrößeÄndernFürLänge(Zeile4, LaengeZeile4);
            SchritGrößeÄndernFürLänge(Zeile5, LaengeZeile5);
            SchritGrößeÄndernFürLänge(Zeile6, LaengeZeile6);
        }

        // "Länge" berechnen
        private void SchritGrößeÄndernFürLänge (TextBox Zeile, TextBox Laenge)
        {
            if (Zeile.Text == "")
            {
                Laenge.Text = "";
                return;
            }
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string x = Zeile.Text;
                double y = 0;
                foreach (char c in x)
                {
                    string sql = "select Breite from Tabelle" + Schriftgröße + " where Zeichen = '" + c + "' COLLATE utf8mb4_bin;";
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    MySqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        y += Int32.Parse(rdr[0].ToString());
                    }
                    rdr.Close();
                }
                y = y / 10;
                Laenge.Text = y.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            conn.Close();
        }

        // Schicken die Zeielen um ihre Länge zu berechnen
        private void Zeile1_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile1, LaengeZeile1);
        }

        private void Zeile2_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile2, LaengeZeile2);
        }

        private void Zeile3_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile3, LaengeZeile3);
        }

        private void Zeile4_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile4, LaengeZeile4);
        }

        private void Zeile5_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile5, LaengeZeile5);
        }

        private void Zeile6_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile6, LaengeZeile6);
        }
        
        private void TextSpeichernOrDrucken_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode)
            {
                if (CheckEveryZeileHasHöhe() == false) { return; }
                CreateNullForDatenbank();
                try
                {
                    if (ArchivCheckBox.Checked && !DruckCheckBox.Checked)
                    {
                        DruckSpeichern("Archiv");
                        MessageBox.Show("Der Druck wurde in Archiv gespeichert");
                    }
                    else if (!ArchivCheckBox.Checked)
                    {
                        DruckSpeichern("Druckdatei");
                        MessageBox.Show("Der Druck wurde in Druckdatei gespeichert");
                    }
                    else if (ArchivCheckBox.Checked && DruckCheckBox.Checked)
                    {
                        DruckSpeichern("Archiv");
                        DruckSpeichern("Druckdatei");
                        MessageBox.Show("Der Druck wurde in Druckdatei und Archiv gespeichert");
                    }
                    EmptyTheFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void ArchivdruckToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DruckSuchen frm2 = new DruckSuchen("Archiv");
            frm2.ShowDialog();
            AktuellDatei = "Archiv";
            Satz_Nr.Text = stzNr;
            AktuellDruck = stzNr;
            Zeile1.Text = z1;
            Zeile2.Text = z2;
            Zeile3.Text = z3;
            Zeile4.Text = z4;
            Zeile5.Text = z5;
            Zeile6.Text = z6;
            HoeheZeile1.Text = h1;
            HoeheZeile2.Text = h2;
            HoeheZeile3.Text = h3;
            HoeheZeile4.Text = h4;
            HoeheZeile5.Text = h5;
            HoeheZeile6.Text = h6;
            Sperren1.Text = s1;
            Sperren2.Text = s2;
            Sperren3.Text = s3;
            Sperren4.Text = s4;
            Sperren5.Text = s5;
            Sperren6.Text = s6;
            BandNr.Text = BdNr;
            try { SchriftGröße.SelectedIndex = Int32.Parse(srft) - 1; } catch { SchriftGröße.SelectedIndex = -1; }
            try { Breite.SelectedIndex = Int32.Parse(brte); } catch { Breite.SelectedIndex = -1; }
            try { AbstandVonAussen.Value = Int32.Parse(AVA); } catch { AbstandVonAussen.Value = 100; }
            try { FarbeEingabe.SelectedIndex = Int32.Parse(frb) - 1; } catch { FarbeEingabe.SelectedIndex = -1; }
        }

        private void DruckSpeichernToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextSpeichernOrDrucken.PerformClick();
        }

        private void druckdateiNeuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("Den Druckdatei wirklich leeren?", "", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
                MySqlConnection conn = new MySqlConnection(connStr);

                try
                {
                    conn.Open();
                    string sql = "CREATE DATABASE IF NOT EXISTS movedb;USE movedb;DROP TABLE IF EXISTS Druckdatei;" +
                        "CREATE TABLE Druckdatei(nr INTEGER AUTO_INCREMENT,Zeile1 VARCHAR(100),Zeile2 VARCHAR(100)," +
                        "Zeile3 VARCHAR(100),Zeile4 VARCHAR(50),Zeile5 VARCHAR(50),Zeile6 VARCHAR(50),Zeile7 VARCHAR(50)," +
                        "Zeile8 VARCHAR(50),Zeile9 VARCHAR(50),Zeile10 VARCHAR(50),Zeile11 VARCHAR(50),Zeile12 VARCHAR(50)," +
                        "Zeile13 VARCHAR(50),Zeile14 VARCHAR(50), Zeile15 VARCHAR(50),Schrift VARCHAR(2), Höhe1 VARCHAR(5)," +
                        "Höhe2 VARCHAR(5), Höhe3 VARCHAR(5),Höhe4 VARCHAR(5),Höhe5 VARCHAR(5),Höhe6 VARCHAR(5),Höhe7 VARCHAR(5)," +
                        "Höhe8 VARCHAR(5),Höhe9 VARCHAR(5),Höhe10 VARCHAR(5),Höhe11 VARCHAR(5), Höhe12 VARCHAR(5), Höhe13 VARCHAR(5)," +
                        "Höhe14 VARCHAR(5),Höhe15 VARCHAR(5),Sperren1 VARCHAR(3),Sperren2 VARCHAR(3), Sperren3 VARCHAR(3)," +
                        "Sperren4 VARCHAR(3),Sperren5 VARCHAR(3), Sperren6 VARCHAR(3), Sperren7 VARCHAR(3), Sperren8 VARCHAR(3)," +
                        "Sperren9 VARCHAR(3), Sperren10 VARCHAR(3),Sperren11 VARCHAR(3), Sperren12 VARCHAR(3), Sperren13 VARCHAR(3)," +
                        "Sperren14 VARCHAR(3), Sperren15 VARCHAR(3), AbstvU VARCHAR(4),Farbe VARCHAR(10),BandNr VARCHAR(5)," +
                        " BandBr VARCHAR(5), Ged VARCHAR(20), Datum VARCHAR(50),INDEX(nr) ) ENGINE = myisam DEFAULT CHARSET = utf8;" +
                        "SET autocommit = 1;";

                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Druckdatei wurde geleert!");
                    EmptyTheFields();
                    UngedruckteZeilenBerechnen();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                conn.Close();
            } else
            {
                return;
            }
        }

       

        // Überprüfen ob jede Zeile ihre "Höhe" hat
        private bool CheckEveryZeileHasHöhe()
        {
            if(!string.IsNullOrEmpty(Zeile1.Text) && string.IsNullOrEmpty(HoeheZeile1.Text)) {
                MessageBox.Show("Höhe Zeile 1 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile2.Text) && string.IsNullOrEmpty(HoeheZeile2.Text))
            {
                MessageBox.Show("Höhe Zeile 2 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile3.Text) && string.IsNullOrEmpty(HoeheZeile3.Text))
            {
                MessageBox.Show("Höhe Zeile 3 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile4.Text) && string.IsNullOrEmpty(HoeheZeile4.Text))
            {
                MessageBox.Show("Höhe Zeile 4 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile5.Text) && string.IsNullOrEmpty(HoeheZeile5.Text))
            {
                MessageBox.Show("Höhe Zeile 5 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile6.Text) && string.IsNullOrEmpty(HoeheZeile6.Text))
            {
                MessageBox.Show("Höhe Zeile 6 fehlt!");
                return false;
            }
            else
            {
                return true;
            }
        }

        // "Breite" zum Index Nummer konvertieren (So ist es in der MS-Datenbank)
        private int BreiteToDatenBank(int Breite)
        {
            if (Breite == 55) return 0;
            else if (Breite == 75) return 1;
            else if (Breite == 100) return 2;
            else if (Breite == 125) return 3;
            else if (Breite == 150) return 4;
            else if (Breite == 175) return 5;
            else if (Breite == 200) return 6;
            else if (Breite == 225) return 7;
            else return -1;
        }

        // 'Ged'-Spalte erstellen (muss man mit 'x' erweitern, wenn es gedruckt ist)
        private string CreateGedLangDruck()
        {
            string Ged = "L";
            int SchriftInt = SchriftGröße.SelectedIndex + 1;
            string SchriftChar = SchriftInt.ToString();

            Ged = (Zeile1.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile2.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile3.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile4.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile5.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile6.Text == "") ? Ged + "0" : Ged + SchriftChar;

            return Ged;
        }

        // Das Problem lösen, dass keine Null an die Datenbank zu senden und keine Leerzeilen zuzulassen
        private void CreateNullForDatenbank()
        {
            List<TextBox> textBoxes = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5, Zeile6,
                HoeheZeile1, HoeheZeile2, HoeheZeile3, HoeheZeile4, HoeheZeile5, HoeheZeile6,
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5, Sperren6, BandNr
            };
            foreach (TextBox t in textBoxes)
            {
                if (t.Text == "")
                {
                    try
                    {
                        GoNull.Add(t, "null");
                    }
                    catch {GoNull[t] = "null";}
                }
                else
                {
                    try
                    {
                        GoNull.Add(t, "\"" + t.Text + "\"");
                    }
                    catch {GoNull[t] = "\"" + t.Text + "\""; }
                }
            }
        }

        // Leeren alle Felder
        private void EmptyTheFields()
        {
            List<TextBox> textBoxes = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5, Zeile6,
                HoeheZeile1, HoeheZeile2, HoeheZeile3, HoeheZeile4, HoeheZeile5, HoeheZeile6,
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5, Sperren6, BandNr, Satz_Nr
            };
            foreach (TextBox t in textBoxes)
            {
                t.Text = "";
            }
            AbstandVonAussen.Value = 100;
            Breite.SelectedIndex = -1;
            FarbeEingabe.SelectedIndex = -1;
        }

        // Speichern von Druck in Druckdatei oder Archiv oder beides
        private void DruckSpeichern(string SaveIn)
        {
            DateTime dt1 = DateTime.Now;
            string Datum = "\"" + dt1.ToString() + "\"";
            int breite = (Breite.Text == "") ? -1 : BreiteToDatenBank(Int32.Parse(Breite.Text));
            int farbe = (FarbeEingabe.Text == "") ? -1 : FarbeEingabe.SelectedIndex + 1;
            string Ged = "\"" + CreateGedLangDruck() + "\"";
            string schrift = (SchriftGröße.SelectedIndex + 1).ToString();
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "INSERT INTO " + SaveIn + "(Zeile1, Zeile2, Zeile3, Zeile4, Zeile5, Zeile6, Schrift," +
                    "Höhe1, Höhe2, Höhe3, Höhe4, Höhe5, Höhe6, Sperren1, Sperren2, Sperren3, Sperren4, Sperren5, Sperren6, " +
                    "AbstvU, Farbe, BandNr, BandBr, Ged, Datum) VALUES " + "(" + GoNull[Zeile1] + ", " + GoNull[Zeile2] + ", " +
                    GoNull[Zeile3] + ", " + GoNull[Zeile4] + ", " + GoNull[Zeile5] + ", " + GoNull[Zeile6] + ", " + schrift + ", " +
                    GoNull[HoeheZeile1] + ", " + GoNull[HoeheZeile2] + ", " + GoNull[HoeheZeile3] + ", " + GoNull[HoeheZeile4] + ", " +
                    GoNull[HoeheZeile5] + ", " + GoNull[HoeheZeile6] + ", " + GoNull[Sperren1] + ", " + GoNull[Sperren2] + ", " +
                    GoNull[Sperren3] + ", " + GoNull[Sperren4] + ", " + GoNull[Sperren5] + ", " + GoNull[Sperren6] + ", " +
                    AbstandVonAussen.Value + ", " + farbe + ", " + GoNull[BandNr] + ", " + breite + ", " + Ged + ", " + Datum + ")";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            conn.Close();
            UngedruckteZeilenBerechnen();
        }

        // Die aktuellen Druck von Druckdatei oder Archiv löschen
        private void DruckLöschen(string AktuellDatei)
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "delete from " + AktuellDatei + " where nr = " +  Satz_Nr.Text + ";";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("den Druck wurde von " + AktuellDatei + " gelöscht");
                EmptyTheFields();
            }
            catch
            {
                MessageBox.Show("Kein Druck zum Löschen!");
            }
            conn.Close();
            UngedruckteZeilenBerechnen();
        }

        // Druck aus Datenbank aufrufen
        private void DruckAufrufen(string SelectDruck)
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "select * from " + AktuellDatei + " where nr = " + SelectDruck;
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    Satz_Nr.Text = rdr["nr"].ToString();
                    Zeile1.Text = rdr["Zeile1"].ToString();
                    Zeile2.Text = rdr["Zeile2"].ToString();
                    Zeile3.Text = rdr["Zeile3"].ToString();
                    Zeile4.Text = rdr["Zeile4"].ToString();
                    Zeile5.Text = rdr["Zeile5"].ToString();
                    Zeile6.Text = rdr["Zeile6"].ToString();
                    HoeheZeile1.Text = rdr["Höhe1"].ToString();
                    HoeheZeile2.Text = rdr["Höhe2"].ToString();
                    HoeheZeile3.Text = rdr["Höhe3"].ToString();
                    HoeheZeile4.Text = rdr["Höhe4"].ToString();
                    HoeheZeile5.Text = rdr["Höhe5"].ToString();
                    HoeheZeile6.Text = rdr["Höhe6"].ToString();
                    Sperren1.Text = rdr["Sperren1"].ToString();
                    Sperren2.Text = rdr["Sperren2"].ToString();
                    Sperren3.Text = rdr["Sperren3"].ToString();
                    Sperren4.Text = rdr["Sperren4"].ToString();
                    Sperren5.Text = rdr["Sperren5"].ToString();
                    Sperren6.Text = rdr["Sperren6"].ToString();
                    Breite.SelectedIndex = Int32.Parse(rdr["BandBr"].ToString());
                    BandNr.Text = rdr["BandNr"].ToString();
                    AbstandVonAussen.Value = Int32.Parse(rdr["AbstvU"].ToString());
                    if (Int32.Parse(rdr["Schrift"].ToString()) != -1)
                    {
                        SchriftGröße.SelectedIndex = Int32.Parse(rdr["Schrift"].ToString()) - 1;
                    }
                    else { SchriftGröße.SelectedIndex = -1; }

                    if (Int32.Parse(rdr["Farbe"].ToString()) != -1)
                    {
                        FarbeEingabe.SelectedIndex = Int32.Parse(rdr["Farbe"].ToString()) - 1;
                    }
                    else { FarbeEingabe.SelectedIndex = -1; }
                }
                rdr.Close();
            }

            catch(Exception e) 
            { 
                MessageBox.Show(e.ToString());
                EmptyTheFields();
            }
        }

        private void DrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AktuellDatei = "Druckdatei";
            DruckAufrufen("(select max(nr) from druckdatei);");
            AktuellDruck = Satz_Nr.Text;
            label3.Text = AktuellDatei;
        }

        private void LetzterDruckArchivToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AktuellDatei = "Archiv";
            DruckAufrufen("(select max(nr) from Archiv);");
            AktuellDruck = Satz_Nr.Text;
            label3.Text = AktuellDatei;
        }

        private void BandToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BandNr.Focus();
        }

        private void DruckfarbeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FarbeEingabe.Focus();
        }

        private void VorherigerDruck_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(AktuellDruck))
            {
                drToolStripMenuItem.PerformClick();
            }
            else
            {
                try
                {
                    DruckAufrufen("(select max(nr) from "+ AktuellDatei + " where nr < " + AktuellDruck + ");");
                    AktuellDruck = Satz_Nr.Text;
                }
                catch { return; }
            }
        }

        private void NächsterDruck_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(AktuellDruck))
            {
                DruckAufrufen("(select min(nr) from druckdatei);");
                AktuellDruck = Satz_Nr.Text;
            }
            else
            {
                try
                {
                    DruckAufrufen("(select min(nr) from " + AktuellDatei + " where nr > " + AktuellDruck + ");");
                    AktuellDruck = Satz_Nr.Text;
                }
                catch { return; }
            }
        }

        // Versuch
        private void PumpeAus_Click(object sender, EventArgs e)
        {
            
        }

        private void DruckLöschenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("den Druck wirklich löschen?", "", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    if (AktuellDatei == "Archiv")
                    {
                        DruckLöschen("Archiv");
                    }
                    else
                    {
                        DruckLöschen("Druckdatei");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else if (dr == DialogResult.No)
            {
                return;
            }
        }

        // Änderungen speichern
        private void DruckÄndern(string AktuellDatei)
        {
            if (CheckEveryZeileHasHöhe() == false) { return; }
            CreateNullForDatenbank();

            DateTime dt1 = DateTime.Now;
            string Datum = "\"" + dt1.ToString() + "\"";
            int breite = (Breite.Text == "") ? -1 : BreiteToDatenBank(Int32.Parse(Breite.Text));
            int farbe = (FarbeEingabe.Text == "") ? -1 : FarbeEingabe.SelectedIndex + 1;
            string Ged = "\"" + CreateGedLangDruck() + "\"";
            string schrift = (SchriftGröße.SelectedIndex + 1).ToString();
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "update " + AktuellDatei + " set Zeile1 = " + GoNull[Zeile1] +
                    ", Zeile2 = " + GoNull[Zeile2] +
                    ", Zeile3 = " + GoNull[Zeile3] +
                    ", Zeile4 = " + GoNull[Zeile4] +
                    ", Zeile5 = " + GoNull[Zeile5] +
                    ", Zeile6 = " + GoNull[Zeile6] +
                    ", Höhe1 = " + GoNull[HoeheZeile1] +
                    ", Höhe2 = " + GoNull[HoeheZeile2] +
                    ", Höhe3 = " + GoNull[HoeheZeile3] +
                    ", Höhe4 = " + GoNull[HoeheZeile4] +
                    ", Höhe5 = " + GoNull[HoeheZeile5] +
                    ", Höhe6 = " + GoNull[HoeheZeile6] +
                    ", Sperren1 = " + GoNull[Sperren1] +
                    ", Sperren2 = " + GoNull[Sperren2] +
                    ", Sperren3 = " + GoNull[Sperren3] +
                    ", Sperren4 = " + GoNull[Sperren4] +
                    ", Sperren5 = " + GoNull[Sperren5] +
                    ", Sperren6 = " + GoNull[Sperren6] +
                    ", BandNr = " + GoNull[BandNr] +
                    ", schrift = " + schrift +
                    ", Datum = " + Datum +
                    ", Ged = " + Ged +
                    ", Farbe = " + farbe +
                    ", BandBr = " + breite +
                    ", AbstvU = " + AbstandVonAussen.Value +
                    " where nr = " + Satz_Nr.Text + ";";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Änderungen wurden in "+ AktuellDatei +" gespeichert");
            }
            catch
            {
                MessageBox.Show("Kein Druck zum Ändern!");
            }
            conn.Close();
        }

        private void DruckToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutoSuchen.PerformClick();
        }

        private void ArchivdruckSpeichernToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (CheckEveryZeileHasHöhe() == false) { return; }
                CreateNullForDatenbank();
                DruckSpeichern("Archiv");
                MessageBox.Show("Der Druck wurde in Archiv gespeichert");
                EmptyTheFields();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void DruckNeuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EmptyTheFields();
        }

        private void BandbreiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Breite.Focus();
        }

        // Ermöglichen Titel ändern Knopf
        private void Satz_Nr_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Satz_Nr.Text)) { AutoSuchen.Visible = true; }
            else if (Bearbeiten_Mode == true && string.IsNullOrEmpty(Satz_Nr.Text)) { AutoSuchen.Visible = false; }
        }

        private void AutoSuchen_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode)
            {
                DialogResult dr;
                dr = MessageBox.Show("Änderungen speichern?", "", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    DruckÄndern(AktuellDatei);
                }
                else if (dr == DialogResult.No)
                {
                    return;
                }
            }
        }

        private void DruckSuchenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DruckSuchen frm2 = new DruckSuchen("Druckdatei");
            frm2.ShowDialog();
            AktuellDatei = "Druckdatei";
            Satz_Nr.Text = stzNr;
            AktuellDruck = stzNr;
            Zeile1.Text = z1;
            Zeile2.Text = z2;
            Zeile3.Text = z3;
            Zeile4.Text = z4;
            Zeile5.Text = z5;
            Zeile6.Text = z6;
            HoeheZeile1.Text = h1;
            HoeheZeile2.Text = h2;
            HoeheZeile3.Text = h3;
            HoeheZeile4.Text = h4;
            HoeheZeile5.Text = h5;
            HoeheZeile6.Text = h6;
            Sperren1.Text = s1;
            Sperren2.Text = s2;
            Sperren3.Text = s3;
            Sperren4.Text = s4;
            Sperren5.Text = s5;
            Sperren6.Text = s6;
            BandNr.Text = BdNr;
            try { SchriftGröße.SelectedIndex = Int32.Parse(srft) - 1; } catch { SchriftGröße.SelectedIndex = -1; }
            try { Breite.SelectedIndex = Int32.Parse(brte); } catch { Breite.SelectedIndex = -1; }
            try { AbstandVonAussen.Value = Int32.Parse(AVA); } catch { AbstandVonAussen.Value = 100; }
            try { FarbeEingabe.SelectedIndex = Int32.Parse(frb) - 1; } catch { FarbeEingabe.SelectedIndex = -1; }
        }

        // die Anzahl der ungedruckten Zeilen berechnen
        private void UngedruckteZeilenBerechnen()
        {
            int p1 = 0, p2 = 0, p3 = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql1 = "select sum(Length(ged) - length(replace(ged, \"1\", \"\"))) from druckdatei;";
                MySqlCommand cmd1 = new MySqlCommand(sql1, conn);
                object result1 = cmd1.ExecuteScalar();
                if (result1 != null)
                {
                    p1 = Convert.ToInt32(result1);
                }
                string sql2 = "select sum(Length(ged) - length(replace(ged, \"2\", \"\"))) from druckdatei;";
                MySqlCommand cmd2 = new MySqlCommand(sql2, conn);
                object result2 = cmd2.ExecuteScalar();
                if (result2 != null)
                {
                    p2 = Convert.ToInt32(result2);
                }
                string sql3 = "select sum(Length(ged) - length(replace(ged, \"3\", \"\"))) from druckdatei;";
                MySqlCommand cmd3 = new MySqlCommand(sql3, conn);
                object result3 = cmd3.ExecuteScalar();
                if (result3 != null)
                {
                    p3 = Convert.ToInt32(result3);
                }
                InfoZeile.Text = "Ungedruckt = " + (p1 + p2 + p3).ToString() + " Zeilen     P1= " + p1.ToString() +
                 "  P2= " + p2.ToString() + "  P3= " + p3.ToString();
            }
            catch
            {
                InfoZeile.Text = "Ungedruckt = " + (p1 + p2 + p3).ToString() + " Zeilen    P1= " + p1.ToString() +
                 "  P2= " + p2.ToString() + "  P3= " + p3.ToString();
            }
            conn.Close();
        }

        // Automatische Zeilenhöhe 
        private void Automatisch_Zeilenhöhe_Click(object sender, EventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen!");
                return;
            }
            else
            {
                HoheSeite1();
                HoheSeite2();
                if (Zeile1.Text == "") { HoeheZeile1.Text = ""; }
                if (Zeile2.Text == "") { HoeheZeile2.Text = ""; }
                if (Zeile3.Text == "") { HoeheZeile3.Text = ""; }
                if (Zeile4.Text == "") { HoeheZeile4.Text = ""; }
                if (Zeile5.Text == "") { HoeheZeile5.Text = ""; }
                if (Zeile6.Text == "") { HoeheZeile6.Text = ""; }
            }
        }

        // Die hohe von 2 Zeilen automatisch berechnen
        private void HoheZweiZeilen(string value,TextBox x, TextBox y)
        {
            string [] splited = value.Split(',');
            x.Text = splited[0];
            y.Text = splited[1];
        }

        // Die hohe von 3 Zeilen automatisch berechnen
        private void HoheDreiZeilen(string value, TextBox x, TextBox y, TextBox z)
        {
            string[] splited = value.Split(',');
            x.Text = splited[0];
            y.Text = splited[1];
            z.Text = splited[2];
        }

        // Erste Seite Höhe
        private void HoheSeite1 ()
        {
            string scrift, band, breite, _zeile = "";
            if (SchriftGröße.SelectedIndex == 0) { scrift = "s1"; }
            else if (SchriftGröße.SelectedIndex == 1) { scrift = "s2"; }
            else { scrift = "s3"; }

            if (BandNr.Text == "992") { band = "n"; }
            else if (BandNr.Text == "993") { band = "m"; }
            else { band = "z"; }

            breite = (!string.IsNullOrEmpty(Breite.Text)) ? Breite.Text : "55";

            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            if (Zeile1.Text != "" && Zeile2.Text == "" && Zeile3.Text == "") { _zeile = "1"; }
            else if (Zeile1.Text != "" && Zeile2.Text != "" && Zeile3.Text == "") { _zeile = "2"; }
            else if (Zeile1.Text != "" && Zeile2.Text != "" && Zeile3.Text != "") { _zeile = "3"; }
            else { return; }
            try
            {
                conn.Open();
                string sql1 = "Select " + scrift + band + _zeile + " from Zeilenhohe where Breite = " + breite + ";";
                MySqlCommand cmd1 = new MySqlCommand(sql1, conn);
                MySqlDataReader rdr1 = cmd1.ExecuteReader();
                while (rdr1.Read())
                {
                    string rd = rdr1[0].ToString();
                    if (string.IsNullOrEmpty(rd)) { HoheSeite1Leeren(); }
                    if (_zeile == "1" && !string.IsNullOrEmpty(rd)) { HoeheZeile1.Text = rd; }
                    else if (_zeile == "2" && !string.IsNullOrEmpty(rd)) { HoheZweiZeilen(rd, HoeheZeile1, HoeheZeile2); }
                    else if (_zeile == "3" && !string.IsNullOrEmpty(rd)) { HoheDreiZeilen(rd, HoeheZeile1, HoeheZeile2, HoeheZeile3); }
                }
                rdr1.Close();
                conn.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        // Zweite Seite Höhe
        private void HoheSeite2()
        {
            string scrift, band, breite, Zeile_ = "";
            if (SchriftGröße.SelectedIndex == 0) { scrift = "s1"; }
            else if (SchriftGröße.SelectedIndex == 1) { scrift = "s2"; }
            else { scrift = "s3"; }

            if (BandNr.Text == "992") { band = "n"; }
            else if (BandNr.Text == "993") { band = "m"; }
            else { band = "z"; }

            breite = (!string.IsNullOrEmpty(Breite.Text)) ? Breite.Text : "55";

            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            if (Zeile4.Text != "" && Zeile5.Text == "" && Zeile6.Text == "") { Zeile_ = "1"; }
            else if (Zeile4.Text != "" && Zeile5.Text != "" && Zeile6.Text == "") { Zeile_ = "2"; }
            else if (Zeile4.Text != "" && Zeile5.Text != "" && Zeile6.Text != "") { Zeile_ = "3"; }
            else { return; }

            try
            {
                conn.Open();
                string sql2 = "Select " + scrift + band + Zeile_ + " from Zeilenhohe where Breite = " + breite + ";"; ;
                MySqlCommand cmd2 = new MySqlCommand(sql2, conn);
                MySqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    string rd2 = rdr2[0].ToString();
                    if (string.IsNullOrEmpty(rd2)) { HoheSeite2Leeren(); }
                    if (Zeile_ == "1" && !string.IsNullOrEmpty(rd2)) { HoeheZeile4.Text = rd2; }
                    else if (Zeile_ == "2" && !string.IsNullOrEmpty(rd2)) { HoheZweiZeilen(rd2, HoeheZeile4, HoeheZeile5); }
                    else if (Zeile_ == "3" && !string.IsNullOrEmpty(rd2)) { HoheDreiZeilen(rd2, HoeheZeile4, HoeheZeile5, HoeheZeile6); }
                }
                rdr2.Close();
                conn.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //Hohe erste 3 Zeilen leeren, Falls es keine automatischen Höhe gibt
        private void HoheSeite1Leeren()
        {
            HoeheZeile1.Text = "";
            HoeheZeile2.Text = "";
            HoeheZeile3.Text = "";
        }

        //Hohe zweite 3 Zeilen leeren, Falls es keine automatischen Höhe gibt
        private void HoheSeite2Leeren()
        {
            HoeheZeile4.Text = "";
            HoeheZeile5.Text = "";
            HoeheZeile6.Text = "";
        }
    }
}
