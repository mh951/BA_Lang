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
using System.IO.Ports;
using System.Threading;


namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public static bool Bearbeiten_Mode = true;
        public static int letzte_Zeile = 1;
        public int Schriftgröße = -1, Tisch_init = 0;
        public static string z1, z2, z3, z4, z5, z6, h1, h2, h3, h4, h5, h6,
                             s1, s2, s3, s4, s5, s6, stzNr, BdNr, brte, AVA, srft, frb, gedd, Dtm;
        Dictionary<TextBox, string> GoNull = new Dictionary<TextBox, string> { };
        Dictionary<TextBox, TextBox> Zeilen_Höhe = new Dictionary<TextBox, TextBox> { };
        public string AktuellDruck;
        public string AktuellDatei = "Druckdatei";
        public static bool Referencefahrt_done = false;
        public static bool weiter = true;
        List<char> SonderZeichen = new List<char> {  'Ä', 'Ü', 'Ö', ':', ';', 'é', 'è', 'á', 'à'};
        List<char> Abstände = new List<char> { ' ', '²', '³', '|', '@', 'µ' };
        List<TextBox> Zeilen = new List<TextBox> { };
        private SerialPort myport;
        private SerialPort myport2;


        public Form1()
        {
            InitializeComponent();
            Paaren();
            Zeilen_füllen();
        }

        // Jede Zeile mit entsprechender Höhe verknüpfen
        private void Paaren()
        {
            Zeilen_Höhe.Add(Zeile1, HoeheZeile1);
            Zeilen_Höhe.Add(Zeile2, HoeheZeile2);
            Zeilen_Höhe.Add(Zeile3, HoeheZeile3);
            Zeilen_Höhe.Add(Zeile4, HoeheZeile4);
            Zeilen_Höhe.Add(Zeile5, HoeheZeile5);
            Zeilen_Höhe.Add(Zeile6, HoeheZeile6);
        }

        // Die Zeilen in einer Liste legen
        private void Zeilen_füllen()
        {
            Zeilen.Add(Zeile1);
            Zeilen.Add(Zeile2);
            Zeilen.Add(Zeile3);
            Zeilen.Add(Zeile4);
            Zeilen.Add(Zeile5);
            Zeilen.Add(Zeile6);
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
            if (Bearbeiten_Mode)
            {
                Bearbeiten_Mode = false;
                OpenPorts();
                if (!Stempel_prüfen())
                {
                    MessageBox.Show("Ventil machen!");
                    ClosePorts();
                    return;
                }
                if (!Bandhalter_prüfen())
                {
                    MessageBox.Show("Bandhalter schließen!");
                    ClosePorts();
                    return;
                }
                DruckMode();
            }
            else
            {

                Bearbeiten_Mode = true;
                BearbeitenMode();
                ClosePorts();
            }
        }

        // Zum Druckmode wechslen
        private void DruckMode()
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

            myport.WriteLine("I1"); //Platte prüfen
            Thread.Sleep(50);
            string platte = myport.ReadExisting();
            Platte.Text = platte;

        }

        private void OpenPorts()
        {
            try
            {
                myport = new SerialPort();
                myport.BaudRate = 9600;
                myport.PortName = "COM7";
                myport.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Error Arduino");
            }
            try
            {
                myport2 = new SerialPort();
                myport2.BaudRate = 19200;
                myport2.PortName = "COM8";
                myport2.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Error Motor");
            }
        }

        private void ClosePorts()
        {
            myport.Close();
            myport2.Close();
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
                return;
            }
            if (e.KeyCode == Keys.D2 && e.Control == true)
            {
                string newchar = "²";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D3 && e.Control == true)
            {
                string newchar = "³";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D4 && e.Control == true)
            {
                string newchar = "@";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D5 && e.Control == true)
            {
                string newchar = "µ";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.Oemcomma && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.OemPeriod && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.F12)
            {
                string newchar = "@-@";
                int selectedIndex = Zeile1.SelectionStart;
                Zeile1.Text = Zeile1.Text.Insert(selectedIndex, newchar);
                Zeile1.SelectionStart = selectedIndex + 3;
            }
        }

        private void Zeile2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile3.Focus();
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.D2 && e.Control == true)
            {
                string newchar = "²";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D3 && e.Control == true)
            {
                string newchar = "³";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D4 && e.Control == true)
            {
                string newchar = "@";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D5 && e.Control == true)
            {
                string newchar = "µ";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.Oemcomma && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.OemPeriod && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.F12)
            {
                string newchar = "@-@";
                int selectedIndex = Zeile2.SelectionStart;
                Zeile2.Text = Zeile2.Text.Insert(selectedIndex, newchar);
                Zeile2.SelectionStart = selectedIndex + 3;
            }
        }

        private void Zeile3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile4.Focus();
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.D2 && e.Control == true)
            {
                string newchar = "²";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D3 && e.Control == true)
            {
                string newchar = "³";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D4 && e.Control == true)
            {
                string newchar = "@";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D5 && e.Control == true)
            {
                string newchar = "µ";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.Oemcomma && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.OemPeriod && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.F12)
            {
                string newchar = "@-@";
                int selectedIndex = Zeile3.SelectionStart;
                Zeile3.Text = Zeile3.Text.Insert(selectedIndex, newchar);
                Zeile3.SelectionStart = selectedIndex + 3;
            }
        }

        private void Zeile4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile5.Focus();
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.D2 && e.Control == true)
            {
                string newchar = "²";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D3 && e.Control == true)
            {
                string newchar = "³";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D4 && e.Control == true)
            {
                string newchar = "@";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D5 && e.Control == true)
            {
                string newchar = "µ";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.Oemcomma && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.OemPeriod && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.F12)
            {
                string newchar = "@-@";
                int selectedIndex = Zeile4.SelectionStart;
                Zeile4.Text = Zeile4.Text.Insert(selectedIndex, newchar);
                Zeile4.SelectionStart = selectedIndex + 3;
            }
        }

        private void Zeile5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile6.Focus();
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.D2 && e.Control == true)
            {
                string newchar = "²";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D3 && e.Control == true)
            {
                string newchar = "³";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D4 && e.Control == true)
            {
                string newchar = "@";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D5 && e.Control == true)
            {
                string newchar = "µ";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.Oemcomma && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.OemPeriod && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.F12)
            {
                string newchar = "@-@";
                int selectedIndex = Zeile5.SelectionStart;
                Zeile5.Text = Zeile5.Text.Insert(selectedIndex, newchar);
                Zeile5.SelectionStart = selectedIndex + 3;
            }
        }

        private void Zeile6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.D2 && e.Control == true)
            {
                string newchar = "²";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D3 && e.Control == true)
            {
                string newchar = "³";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D4 && e.Control == true)
            {
                string newchar = "@";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.D5 && e.Control == true)
            {
                string newchar = "µ";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.Oemcomma && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.OemPeriod && e.Control == true)
            {
                string newchar = "|";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 1;
            }
            if (e.KeyCode == Keys.F12)
            {
                string newchar = "@-@";
                int selectedIndex = Zeile6.SelectionStart;
                Zeile6.Text = Zeile6.Text.Insert(selectedIndex, newchar);
                Zeile6.SelectionStart = selectedIndex + 3;
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
            if (e.KeyCode == Keys.F12 && e.Control == true)
            {
                dauersuchenToolStripMenuItem.PerformClick();
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
        private void SchritGrößeÄndernFürLänge(TextBox Zeile, TextBox Laenge)
        {
            if (Zeile.Text == "")
            {
                Laenge.Text = "";
                return;
            }
            TextBox ActualSperren;
            if (Zeile == Zeile1) { ActualSperren = Sperren1; }
            else if (Zeile == Zeile2) { ActualSperren = Sperren2; }
            else if (Zeile == Zeile3) { ActualSperren = Sperren3; }
            else if (Zeile == Zeile4) { ActualSperren = Sperren4; }
            else if (Zeile == Zeile5) { ActualSperren = Sperren5; }
            else { ActualSperren = Sperren6; }
            bool nichtleersperren = (!string.IsNullOrEmpty(ActualSperren.Text));
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string x = Zeile.Text.Trim();
                double y = 0;
                string strC = "";
                foreach (char c in x)
                {
                    if (c == '\\') { strC = "\\\\"; }
                    else if (c == '\'') { strC = "\\\'"; }
                    else { strC = c.ToString(); }
                    string sql = "select Breite from Tabelle" + Schriftgröße + " where Zeichen = '" + strC + "' COLLATE utf8mb4_bin;";
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    MySqlDataReader rdr = cmd.ExecuteReader();
                    if (nichtleersperren)
                    {
                        while (rdr.Read())
                        {
                            y += Int32.Parse(rdr[0].ToString()) + (Int32.Parse(ActualSperren.Text) * 5);
                        }
                    }
                    else
                    {
                        while (rdr.Read())
                        {
                            y += Int32.Parse(rdr[0].ToString());
                        }
                    }
                    rdr.Close();
                }
                if (nichtleersperren) y -= Int32.Parse(ActualSperren.Text);

                y /= 10;
                Laenge.Text = y.ToString();
                CheckLängsteZeile(Laenge);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            conn.Close();
        }

        // kontroliere welche Zeile die längsten ist, um ihre Länge mit Rot zu färben
        private void CheckLängsteZeile(TextBox Laenge)
        {
            if (string.IsNullOrEmpty(Laenge.Text)) return;
            List<TextBox> textBoxes = new List<TextBox>
            { LaengeZeile1, LaengeZeile2, LaengeZeile3, LaengeZeile4,
              LaengeZeile5, LaengeZeile6};

            TextBox Höhste = Laenge;
            foreach (TextBox tb in textBoxes)
            {
                if (tb == Höhste) continue;
                else if (string.IsNullOrEmpty(tb.Text))
                {
                    tb.BackColor = Color.White;
                    continue;
                }
                if (Convert.ToDouble(tb.Text) > Convert.ToDouble(Höhste.Text)) Höhste = tb;
                else { tb.BackColor = Color.White; }
            }
            Höhste.BackColor = Color.LightPink;
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
                    }
                    else if (!ArchivCheckBox.Checked)
                    {
                        DruckSpeichern("Druckdatei");
                    }
                    else if (ArchivCheckBox.Checked && DruckCheckBox.Checked)
                    {
                        DruckSpeichern("Archiv");
                        DruckSpeichern("Druckdatei");
                    }
                    EmptyTheFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                TextSpeichernOrDrucken.Enabled = false;
                this.ActiveControl = null;

                if (!Referencefahrt_done)
                {
                    MessageBox.Show("Eine Referenzfahrt ist vorher nötig!");
                    return;
                }
                if (Platte_prüfen().Trim() != (SchriftGröße.SelectedIndex + 1).ToString())
                {
                    MessageBox.Show("Bitte die richtige Platte benutzen!");
                    return;
                }
                if (!Bandhalter_prüfen())
                {
                    MessageBox.Show("Bandhalter schließen!");
                    return;
                }
                if (Schriftgröße == -1)
                {
                    MessageBox.Show("Bitte Schriftgröße wählen!");
                    return;
                }
                if (string.IsNullOrEmpty(HoeheZeile1.Text))
                {
                    MessageBox.Show("Bitte Höhe wählen!");
                    return;
                }

                string Zum_Drucken = Zeile1.Text.Trim();
                if (string.IsNullOrEmpty(Zum_Drucken)) return;
                else
                {
                    myport.WriteLine("O31"); // Pumpe einschalten
                    int Sprr = (!string.IsNullOrEmpty(Sperren1.Text)) ? Int32.Parse(Sperren1.Text) : 0;
                    Tisch_init = 8380 - (Int32.Parse(HoeheZeile1.Text) * 40);
                    int ABStVA = (int)((AbstandVonAussen.Value - 34) * 40);
                    
                    int M3 = 0;
                    int AnzahlVonABSt = 0;
                    for (int x = Zum_Drucken.Length; x > 0; x--)
                    {
                        int M5;
                        try
                        {
                            if (x == Zum_Drucken.Length)
                            {
                                if (SonderZeichen.Contains(Zum_Drucken[x - 1]))
                                {
                                    SonderZeichen_Drucken(Zum_Drucken[x - 1],"0", "20", ABStVA.ToString());
                                    M3 = 0;
                                    continue;
                                }
                                else if (Abstände.Contains(Zum_Drucken[x - 1]))
                                {
                                    continue;
                                }
                                Motoren_1A_2A_3R_4A_5R(Zum_Drucken[x - 1].ToString(), M3.ToString(), "20", "1", ABStVA);
                            }
                            else
                            {
                                if (Abstände.Contains(Zum_Drucken[x - 1]))
                                {
                                    M3 += Zeichen_Breite(Zum_Drucken[x - 1]) * 4;
                                    AnzahlVonABSt += 1;
                                    continue;
                                }
                                if (!Abstände.Contains(Zum_Drucken[x]))
                                {
                                    M3 += Convert.ToInt32((Zeichen_Breite(Zum_Drucken[x]) + Zeichen_Breite(Zum_Drucken[x - 1])) * 2 + 20);
                                }
                                else
                                {
                                    M3 += Convert.ToInt32((Zeichen_Breite(Zum_Drucken[x + AnzahlVonABSt]) + Zeichen_Breite(Zum_Drucken[x - 1])) * 2 + 20);
                                }
                                
                                M3 += Sprr * 20;
                                M5 = Convert.ToInt32((Zeichen_Breite(Zum_Drucken[x + AnzahlVonABSt]) + Zeichen_Breite(Zum_Drucken[x - 1])) * 0.12);
                                

                                if (SonderZeichen.Contains(Zum_Drucken[x - 1]))
                                {
                                    SonderZeichen_Drucken(Zum_Drucken[x - 1], M3.ToString(), M5.ToString());
                                    M3 = 0;
                                    continue;
                                }
                                Motoren_1A_2A_3R_4A_5R(Zum_Drucken[x - 1].ToString(), M3.ToString(), M5.ToString());
                            }
                            Motoren_stehen();
                            Stempel_ab(Zum_Drucken[x - 1]);
                            Stempel_auf();
                            bool Stmp = true;
                            while (Stmp)
                            {
                                Stmp = !Stempel_prüfen();
                            }
                            Trennung();
                            M3 = 0;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                    Motore_5_Drehen_Relativ("20");
                    StartPunkt();
                    myport.WriteLine("O30"); // Pumpe aus
                    Tisch_init = 0;
                }
                TextSpeichernOrDrucken.Enabled = true;
                this.ActiveControl = null;

            }
        }



        // Die Sonderzeichen Drucken
        private void SonderZeichen_Drucken(char v, string M3 = "0", string M5 = "0", string abst = "0")
        {
            char b = ' ';
            char Sonder = ' ';
            
            if (v == 'Ä')
            {
                b = 'A';
                Sonder = '*';
            }
            else if (v == 'Ö')
            {
                b = 'O';
                Sonder = '*';
            }
            else if (v == 'Ü')
            {
                b = 'U';
                Sonder = '*';
            }
            else if (v == ':')
            {
                b = '.';
                Sonder = '.';
            }
            else if (v == ';')
            {
                b = ',';
                Sonder = '.';
            }
            else if (v == 'é')
            {
                b = 'e';
                Sonder = '´';
            }
            else if (v == 'è')
            {
                b = 'e';
                Sonder = '`';
            }
            else if (v == 'á')
            {
                b = 'a';
                Sonder = '´';
            }
            else if (v == 'à')
            {
                b = 'a';
                Sonder = '`';
            }
            //MotorenDrehen(b);
            Motoren_1A_2A_3R_4A_5R(b.ToString(), M3, M5, "1", Int32.Parse(abst));
            Motoren_stehen();
            Stempel_ab(b);
            Stempel_auf();
            bool Stmp = true;
            while (Stmp)
            {
                Stmp = !Stempel_prüfen();
            }
            Trennung();
            string sdr = TIsch_Pos_bringen(v.ToString());
            Motoren_1A_2A_3R_4A_5R(Sonder.ToString(), "0", M5, sdr);
            Motoren_stehen();
            Stempel_ab(Sonder);
            Stempel_auf();
            Stmp = true;
            while (Stmp)
            {
                Stmp = !Stempel_prüfen();
            }
            Trennung();
        }

        private string TIsch_Pos_bringen(string x)
        {
            string TischPos = "";
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select TischPos from tabelle" + Schriftgröße + " where Zeichen = '" + x + "' COLLATE utf8mb4_bin";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    TischPos = rdr["TischPos"].ToString();
                }
                rdr.Close();
            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
            return TischPos;
        }

        // Trennt die Foile und Band (ein)
        private void Trennung()
        {
            myport.WriteLine("O41"); // Trennt die Foile und Band
            Thread.Sleep(300);
            myport.WriteLine("O40");
        }

        // Stempel ab (nach unten bewegen)
        private void Stempel_ab(char c)
        {
            int druck = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select Druck from Tabelle" + Schriftgröße + " where Zeichen = '" + c + "' COLLATE utf8mb4_bin;";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    druck = Int32.Parse(rdr[0].ToString());
                }
                conn.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Double dk = druck * Decimal.ToDouble(Druckstärke.Value) / 550;
            druck = Convert.ToInt32(dk);
            myport.WriteLine("A" + druck.ToString()); // Schick PWM-signal von Arduino an Ventil 1
            myport.WriteLine("O11"); // Ventil ab (nach unten)
            Thread.Sleep(1000);
        }

        // Stempel auf (nach oben bewegen)
        private void Stempel_auf()
        {
            myport.WriteLine("O10"); // Ventil 1  abschalten
            Thread.Sleep(100);
            myport.WriteLine("A130"); // Ventil 2 PWM-signal bekommen
            myport.WriteLine("O21"); // Ventil 2 einschalten (nach oben)
            Thread.Sleep(500);
            myport.WriteLine("O20"); // Ventil 2 ausschalten
        }

        // Richtige Platte prüfen
        private string Platte_prüfen()
        {
            myport.WriteLine("I1");
            Thread.Sleep(100);
            return myport.ReadExisting();
        }

        // Bandhalter prüfen
        private bool Bandhalter_prüfen()
        {
            myport.WriteLine("I3"); //Bandhalter prüfen
            Thread.Sleep(50);
            string halter = myport.ReadExisting();
            int X = Int32.Parse(halter);
            return X != 0;
        }

        // Stempel prüfen
        private bool Stempel_prüfen()
        {
            myport.WriteLine("I2");
            Thread.Sleep(50);
            string X = myport.ReadExisting().Trim();
            return (X == "1");
        }

        // Motoren 1,2 und 4 mit Drehen Anfangen
        private void MotorenDrehen(char x, int y = 1)
        {
            int XPos = 0, YPos = 0, TischPos = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select XPos, YPos, TischPos from tabelle" + Schriftgröße + " where Zeichen = '" + x + "' COLLATE utf8mb4_bin";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    XPos = Int32.Parse(rdr["XPos"].ToString());
                    YPos = Int32.Parse(rdr["YPos"].ToString());
                    TischPos = Int32.Parse(rdr["TischPos"].ToString());
                }
                rdr.Close();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            if(XPos != 0 && YPos != 0)
            {
                // Motor 1
                string message = "#1p2\r";  //Positionierart Absolut Motor 1
                myport2.WriteLine(message);
                Thread.Sleep(50);
                string auslesen = myport2.ReadExisting();

                message = "#1s" + XPos.ToString() + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Thread.Sleep(100);

                message = "#1A\r";
                myport2.WriteLine(message);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();

                // Motor 2
                message = "#" + (char)2 + "p2\r";  //Positionierart Absolut Motor 2
                myport2.WriteLine(message);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();

                message = "#" + (char)2 + "d1\r"; //Drehrichtung links
                myport2.WriteLine(message);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();

                message = "#" + (char)2 + "s" + YPos.ToString() + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();

                message = "#" + (char)2 + "A\r";
                myport2.WriteLine(message);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
            }
            // Motoren_stehen();
            // Motor 4
            if (TischPos != 0 && y == 1)
            {
                int Show = (TischPos + Tisch_init);
                Motore_4_Drehen_Absolut(Show.ToString());
            }
        }

        // Motor 3 Absolut Drehen
        private void Motore_3_Drehen(string x)
        {
            // Motor 3
            string message = "#" + (char)3 + "p2\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            string auslesen = myport2.ReadExisting();

            message = "#" + (char)3 + "d2\r"; //Drehrichtung rechts
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)3 + "s" + x + "\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)3 + "A\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();
        }

        // Motor 3 Relativ Drehen
        private void Motore_3_Drehen_Relativ(string x)
        {
            // Motor 3
            string message = "#" + (char)3 + "p1\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            string auslesen = myport2.ReadExisting();

            message = "#" + (char)3 + "d1\r"; //Drehrichtung links
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)3 + "s" + x + "\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)3 + "A\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();
        }

        // Motor 4 Absolut Drehen
        private void Motore_4_Drehen_Absolut(string x)
        {
            // Motor 4
            string message = "#" + (char)4 + "p2\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            string auslesen = myport2.ReadExisting();

            /*message = "#" + (char)4 + "d1\r"; //Drehrichtung links
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();*/

            message = "#" + (char)4 + "s" + x + "\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)4 + "A\r";
            myport2.WriteLine(message);
            Thread.Sleep(50);
            auslesen = myport2.ReadExisting();
        }

        // Motor 4 Relativ Drehen
        private void Motore_4_Drehen_Relativ(string x)
        {
            
            // Motor 4
            string message = "#" + (char)4 + "p1\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);
            string auslesen = myport2.ReadExisting();

            message = "#" + (char)4 + "d0\r"; //Drehrichtung links
            myport2.WriteLine(message);
            Thread.Sleep(20);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)4 + "s" + x + "\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);
            auslesen = myport2.ReadExisting();

            message = "#" + (char)4 + "A\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);
            auslesen = myport2.ReadExisting();
            //Tisch_init = Int32.Parse(x);
        }


        private void Motoren_3_A_4_A(string x, string y)
        {
            string message = "#" + (char)3 + "p2\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);

            message = "#" + (char)4 + "p1\r";
            myport2.WriteLine(message);

            Thread.Sleep(50);

            message = "#" + (char)4 + "d0\r"; //Drehrichtung links
            myport2.WriteLine(message);
            Thread.Sleep(20);

            message = "#" + (char)3 + "s" + x + "\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);

            message = "#" + (char)4 + "s" + y + "\r";
            myport2.WriteLine(message);

            Thread.Sleep(50);

            message = "#" + (char)3 + "A\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);

            message = "#" + (char)4 + "A\r";
            myport2.WriteLine(message);

            Thread.Sleep(50);
        }

        private void Motoren_1A_2A_3R_4A_5R (string _124, string _3, string _5, string t = "1", int ab = 0)
        {
            Tisch_init = (Int32.Parse(HoeheZeile1.Text) * 40);
            int XPos = 0, YPos = 0, TischPos = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select XPos, YPos, TischPos from tabelle" + Schriftgröße + " where Zeichen = '" + _124 + "' COLLATE utf8mb4_bin";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    XPos = Int32.Parse(rdr["XPos"].ToString());
                    YPos = Int32.Parse(rdr["YPos"].ToString());
                    TischPos = Int32.Parse(rdr["TischPos"].ToString());
                }
                rdr.Close();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            string message;
            if (XPos != 0 && YPos != 0)
            {
                message = "#1p2\r";  //Positionierart Absolut Motor 1
                myport2.WriteLine(message);
                Thread.Sleep(20);

                message = "#" + (char)2 + "p2\r";  //Positionierart Absolut Motor 2
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            message = "#" + (char)4 + "p2\r";  //Positionierart Absolut Motor 2
            myport2.WriteLine(message);
            Thread.Sleep(20);
            
            if (ab == 0 && _3 != "0")
            {
                message = "#" + (char)3 + "p1\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }
            else if (ab != 0)
            {
                message = "#" + (char)3 + "p2\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            if (_5 != "0")
            {
                message = "#" + (char)5 + "p1\r";
                myport2.WriteLine(message);
            }
                
            Thread.Sleep(50);

            if (ab == 0 && _3 != "0")
            {
                message = "#" + (char)3 + "d1\r"; //Drehrichtung rechts
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            if (_5 != "0")
            {
                message = "#" + (char)5 + "d0\r"; //Drehrichtung rechts
                myport2.WriteLine(message);
            }

            Thread.Sleep(50);


            if (XPos != 0 && YPos != 0)
            {
                message = "#1s" + XPos.ToString() + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);

                message = "#" + (char)2 + "s" + YPos.ToString() + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            if (t == "1")
            {
                int Vier = (TischPos + Tisch_init);
                message = "#" + (char)4 + "s" + Vier.ToString() + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            else
            {
                int Vier = (Int32.Parse(t)+ Tisch_init);
                message = "#" + (char)4 + "s" + Vier + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }


            if (ab == 0 && _3 != "0")
            {
                message = "#" + (char)3 + "s" + _3 + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            else if (ab != 0)
            {
                message = "#" + (char)3 + "s" + ab.ToString() + "\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }

            if (_5 != "0")
            {
                message = "#" + (char)5 + "s" + _5 + "\r";
                myport2.WriteLine(message);
            }


            Thread.Sleep(50);

            if (XPos != 0 && YPos != 0)
            {
                message = "#1A\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);

                message = "#" + (char)2 + "A\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }
            
            message = "#" + (char)4 + "A\r";
            myport2.WriteLine(message);
            Thread.Sleep(20);

            if (ab !=0 || _3 != "0")
            {
                message = "#" + (char)3 + "A\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }
            

            if (_5 != "0")
            {
                message = "#" + (char)5 + "A\r";
                myport2.WriteLine(message);
            }
            Thread.Sleep(50);
        }

        // Motor 5 Relativ Drehen
        private void Motore_5_Drehen_Relativ(string x)
        {
            // Motor 5
            int Schritt = Int32.Parse(x);
            string Daten = "#" + (char)5 + "!1\r";      //Betriebsart
            myport2.WriteLine(Daten);
            Daten = "#" + (char)5 + "p1\r";        //Positionierart setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "s" + Schritt + "\r";     //Schritte setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "u250\r";     //Startfrequenz setzten
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "o300\r";     //Max Frequenz setzten
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "b3\r";       //Rampe setzten
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "do\r";       //Richtung setzten
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "W1\r";       //Durchgänge setzten
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            Daten = "#" + (char)5 + "A\r";       //starten
            myport2.WriteLine(Daten);
        } 

        private void ReferenzFahrt_Click(object sender, EventArgs e)
        {
            if (!Bearbeiten_Mode)
            {
                if (!Bandhalter_prüfen())
                {
                    MessageBox.Show("Bandhalter schließen!");
                    return;
                }
                string PArt = "4";
                string FStart = "400";
                string FMax = "1000";
                string Rampe = "4";
                string Rtg = "0";
                int St1 = 0;
                int St2 = 0;
                int St3 = 0;
                int St4 = 0;
                int Status = 0;
                //Bandhalter prüfen
                myport2.ReadExisting();
                string Daten = "#W1\r";       //Wiederholung 1x
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                string auslesen = myport2.ReadExisting();
                Daten = "#" + (char)2 + "W1\r";
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                //TbDaten.Text += auslesen;
                Daten = "#" + (char)3 + "W1\r";
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                //TbDaten.Text += auslesen;
                Daten = "#" + (char)4 + "W1\r";
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)5 + "W1\r";
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                //TbDaten.Text += auslesen;  
                Daten = "#" + (char)2 + "leb1\r";    //Endschalter setzten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "leb1\r";    //Endschalter setzten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "leb1\r";    //Endschalter setzten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                Daten = "#1p" + PArt + "\r";        //Positionierart setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                Daten = "#" + (char)2 + "p" + PArt + "\r";        //Positionierart setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "p" + PArt + "\r";        //Positionierart setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "p" + PArt + "\r";        //Positionierart setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#1u400\r";        //Startfrequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)2 + "u" + FStart + "\r";        //Startfrequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "u" + FStart + "\r";        //Startfrequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "u" + FStart + "\r";        //Startfrequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#1o" + FMax + "\r";        //Max-Frequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                Daten = "#" + (char)2 + "o" + FMax + "\r";        //Max-Frequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "o" + FMax + "\r";        //Max-Frequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "o" + FMax + "\r";        //Max-Frequenz setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                //Daten = "#1b" + Rampe + "\r";        //Rampe setzen
                //myport2.WriteLine(Daten);
                Daten = "#" + (char)2 + "b" + Rampe + "\r";        //Rampe setzen
                myport2.WriteLine(Daten);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "b" + Rampe + "\r";        //Rampe setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "b" + Rampe + "\r";        //Rampe setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#1d" + Rtg + "\r";        // Richtung setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)2 + "d" + Rtg + "\r";        // Richtung setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "d" + Rtg + "\r";        // Richtung setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "d" + Rtg + "\r";        // Richtung setzen
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#1A\r";        // starten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)2 + "A" + "\r";        // starten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)3 + "A" + "\r";        // starten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Daten = "#" + (char)4 + "A" + "\r";        // starten
                myport2.WriteLine(Daten);
                Thread.Sleep(50);
                auslesen = myport2.ReadExisting();
                Thread.Sleep(100);
                int pause = 0;
                do
                {
                    pause += 1;
                    if (St1 < 1)
                    {
                        Daten = "#1$\r";        // Status abfragen
                        myport2.WriteLine(Daten);
                        Thread.Sleep(100);
                        auslesen = myport2.ReadExisting();
                        Thread.Sleep(100);
                        Thread.Sleep(100);
                        if (auslesen.EndsWith("163\r") || auslesen.EndsWith("161\r"))
                        {
                            St1 = 1;
                        }
                    }
                    if (St2 < 1)
                    {
                        Daten = "#" + (char)2 + "$" + "\r";        // Status abfragen
                        myport2.WriteLine(Daten);
                        Thread.Sleep(100);
                        auslesen = myport2.ReadExisting();
                        if (auslesen.EndsWith("\r") || auslesen.EndsWith("\r"))
                        {
                            St2 = 1;
                        }
                    }
                    if (St3 < 1)
                    {
                        Daten = "#" + (char)3 + "$" + "\r";        // Status abfragen
                        myport2.WriteLine(Daten);
                        Thread.Sleep(100);
                        auslesen = myport2.ReadExisting();
                        //TbDaten.Text += auslesen;
                        if (auslesen.EndsWith("\r") || auslesen.EndsWith("\r"))
                        {
                            St3 = 1;
                        }
                    }
                    if (St4 < 1)
                    {
                        Daten = "#" + (char)4 + "$" + "\r";        // Status abfragen
                        myport2.WriteLine(Daten);
                        Thread.Sleep(50);
                        auslesen = myport2.ReadExisting();
                        //TbDaten.Text += auslesen;
                        if (auslesen.EndsWith("\r") || auslesen.EndsWith("\r"))
                        {
                            St4 = 1;
                        }
                    }
                    Status = St1 + St2 + St3 + St4;
                }
                while (Status < 4);
                StartPunkt();
                Referencefahrt_done = true;
                TextSpeichernOrDrucken.Focus();
            }
        }
        // Zum Startpunkt fahren (nach Referenzfahrt)
        private void StartPunkt()
        {
            // Daten für Normalbetrieb
            myport2.ReadExisting();
            string Daten = "#1p2\r";        //Positionierart setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(100);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "p2\r";        //Positionierart setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "p2\r";         //Positionierart setzen
            myport2.WriteLine(Daten);
            myport2.ReadExisting();
            Daten = "#" + (char)4 + "p2\r";        //Positionierart setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#1u400\r";        //Startfrequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "u700\r";        //Startfrequenz setzen
            myport2.WriteLine(Daten);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "u700\r";        //Startfrequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)4 + "u700\r";          //Startfrequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#1o3200\r";        //Max-Frequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "o3000\r";        //Max-Frequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "o3000\r";        //Max-Frequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)4 + "o3200\r";        //Max-Frequenz setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#1b10000\r";        //Rampe setzen 19115
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "b13\r";        //Rampe setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "b15\r";        //Rampe setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)4 + "b14\r";        //Rampe setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#1s6000\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "s100\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "s4000\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)4 + "s8380\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Thread.Sleep(100);
            myport2.ReadExisting();
            Daten = "#1A\r";        // starten
            myport2.WriteLine(Daten);
            Thread.Sleep(100);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "A" + "\r";        // starten
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "A" + "\r";        // starten
            myport2.WriteLine(Daten);
            myport2.ReadExisting();
            Thread.Sleep(50);
            Daten = "#" + (char)4 + "A" + "\r";        // starten
            myport2.WriteLine(Daten);
            myport2.ReadExisting();
        }

        // Prüfen ob die Motoren stehen
        private void Motoren_stehen()
        {
            int St1 = 0, St2 = 0, St3 = 0, St4 = 0, Status = 0;
            string Daten, auslesen;
            do
            {
                if (St1 < 1)
                {
                    Daten = "#1$\r";        // Status abfragen
                    myport2.WriteLine(Daten);
                    Thread.Sleep(100);
                    auslesen = myport2.ReadExisting();
                    Thread.Sleep(100);
                    Thread.Sleep(100);
                    if (auslesen.EndsWith("163\r") || auslesen.EndsWith("161\r"))
                    {
                        St1 = 1;
                    }
                }
                if (St2 < 1)
                {
                    Daten = "#" + (char)2 + "$" + "\r";        // Status abfragen
                    myport2.WriteLine(Daten);
                    Thread.Sleep(100);
                    auslesen = myport2.ReadExisting();
                    if (auslesen.EndsWith("\r") || auslesen.EndsWith("\r"))
                    {
                        St2 = 1;
                    }
                }
                if (St3 < 1)
                {
                    Daten = "#" + (char)3 + "$" + "\r";        // Status abfragen
                    myport2.WriteLine(Daten);
                    Thread.Sleep(100);
                    auslesen = myport2.ReadExisting();
                    //TbDaten.Text += auslesen;
                    if (auslesen.EndsWith("\r") || auslesen.EndsWith("\r"))
                    {
                        St3 = 1;
                    }
                }
                if (St4 < 1)
                {
                    Daten = "#" + (char)4 + "$" + "\r";        // Status abfragen
                    myport2.WriteLine(Daten);
                    Thread.Sleep(50);
                    auslesen = myport2.ReadExisting();
                    //TbDaten.Text += auslesen;
                    if (auslesen.EndsWith("\r") || auslesen.EndsWith("\r"))
                    {
                        St4 = 1;
                    }
                }
                Status = St1 + St2 + St3 + St4;
            } while (Status < 4);
        }

        // Breite von Zeichen aufrufen
        private int Zeichen_Breite(char c)
        {
            int y = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select Breite from Tabelle" + Schriftgröße + " where Zeichen = '" + c + "' COLLATE utf8mb4_bin;";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    y += Int32.Parse(rdr[0].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            conn.Close();
            return y;
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

        private void textDruckfensterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FensterWechseln.PerformClick();
        }
        private void referenzfahrtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReferenzFahrt.PerformClick();
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
                EmptyTheFields();
            }
            catch
            {
                MessageBox.Show("Kein Druck zum Löschen!");
            }
            conn.Close();
            UngedruckteZeilenBerechnen();
        }

        private void Sperren1_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile1, LaengeZeile1);
        }

        private void Sperren2_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile2, LaengeZeile2);
        }

        private void Sperren3_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile3, LaengeZeile3);
        }

        private void Sperren4_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile4, LaengeZeile4);
        }

        private void Sperren5_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile5, LaengeZeile5);
        }

        private void Sperren6_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile6, LaengeZeile6);
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
                CheckLängsteZeile(LaengeZeile1);
            }

            catch(Exception e) 
            { 
                MessageBox.Show(e.ToString());
                EmptyTheFields();
            }
        }

        private void ausdruckToolStripMenuItem_Click(object sender, EventArgs e)
        {

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
                EmptyTheFields();
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
                if (Breite.SelectedIndex == 1) { AbstandVonAussen.Value = 42; }
                else if (Breite.SelectedIndex == 2) { AbstandVonAussen.Value = 45; }
                else if (Breite.SelectedIndex == 3) { AbstandVonAussen.Value = 60; }
                else { AbstandVonAussen.Value = 100; }
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

        private void druckenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextSpeichernOrDrucken.PerformClick();
        }

        private void teildruckToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        // Versuch
        private void PumpeAus_Click(object sender, EventArgs e)
        {
            Referencefahrt_done = true;
        }

    }
}
