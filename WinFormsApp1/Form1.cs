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
        Dictionary<TextBox, TextBox> Zeilen_Sperren = new Dictionary<TextBox, TextBox> { };
        Dictionary<TextBox, TextBox> Zeilen_Länge = new Dictionary<TextBox, TextBox> { };
        Dictionary<TextBox, TextBox> Zeilen_Aus_Länge = new Dictionary<TextBox, TextBox> { };
        Dictionary<int, TextBox> Nummer_Zeilen = new Dictionary<int, TextBox> { };
        public string AktuellDruck;
        public string AktuellDatei = "Druckdatei";
        public static bool Referencefahrt_done = false;
        public static bool weiter = true;
        public static bool Druck_done = false;
        public static bool USB_geht = false;
        public static bool Nexter_Druck = false;
        List<char> SonderZeichen = new List<char> {  'Ä', 'Ü', 'Ö', ':', ';', 'é', 'è', 'á', 'à'};
        List<char> Abstände = new List<char> { ' ', '²', '³', '|', '@', 'µ' };
        List<TextBox> Zeilen = new List<TextBox> { };
        List<TextBox> AlleZeilen = new List<TextBox> { };
        private SerialPort myport;
        private SerialPort myport2;
        private TextBox Längste_Zeile;


        /*************************************** Load, paaren, füllen, Modes und Ports ***************************************/


        public Form1()
        {
            InitializeComponent();
            Paaren_Höhe();
            Erste_Zeilen_füllen();
            Paaren_Sperren();
            Paaren_Länge();
            Paaren_Aus_Länge();
            Paaren_Nummer();
            AlleZeilen_füllen();
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
            seite2ToolStripMenuItem.Enabled = false;
            AktuellDatei = "Druckdatei";
            this.ActiveControl = SchriftGröße;
            label3.Text = AktuellDatei;
        }


        // Jede Zeile mit entsprechender Höhe verknüpfen
        private void Paaren_Höhe()
        {
            Zeilen_Höhe.Add(Zeile1, HoeheZeile1);
            Zeilen_Höhe.Add(Zeile2, HoeheZeile2);
            Zeilen_Höhe.Add(Zeile3, HoeheZeile3);
            Zeilen_Höhe.Add(Zeile4, HoeheZeile4);
            Zeilen_Höhe.Add(Zeile5, HoeheZeile5);
            Zeilen_Höhe.Add(Zeile6, HoeheZeile6);
        }

        // Jede Zeile mit entsprechender Sperren verknüpfen
        private void Paaren_Sperren()
        {
            Zeilen_Sperren.Add(Zeile1, Sperren1);
            Zeilen_Sperren.Add(Zeile2, Sperren2);
            Zeilen_Sperren.Add(Zeile3, Sperren3);
            Zeilen_Sperren.Add(Zeile4, Sperren4);
            Zeilen_Sperren.Add(Zeile5, Sperren5);
            Zeilen_Sperren.Add(Zeile6, Sperren6);
        }

        // Paaren Zeile-Länge 
        private void Paaren_Länge()
        {
            Zeilen_Länge.Add(Zeile1, LaengeZeile1);
            Zeilen_Länge.Add(Zeile2, LaengeZeile2);
            Zeilen_Länge.Add(Zeile3, LaengeZeile3);
            Zeilen_Länge.Add(Zeile4, LaengeZeile4);
            Zeilen_Länge.Add(Zeile5, LaengeZeile5);
            Zeilen_Länge.Add(Zeile6, LaengeZeile6);
        }

        // Paaren Länge-Zeile
        private void Paaren_Aus_Länge()
        {
            Zeilen_Aus_Länge.Add(LaengeZeile1,Zeile1);
            Zeilen_Aus_Länge.Add(LaengeZeile2,Zeile2);
            Zeilen_Aus_Länge.Add(LaengeZeile3,Zeile3);
            Zeilen_Aus_Länge.Add(LaengeZeile4,Zeile4);
            Zeilen_Aus_Länge.Add(LaengeZeile5,Zeile5);
            Zeilen_Aus_Länge.Add(LaengeZeile6,Zeile6);
        }
        
        // Jede Zeile mit entsprechende Nummer verknüpfen
        private void Paaren_Nummer()
        {
            Nummer_Zeilen.Add(1, Zeile1);
            Nummer_Zeilen.Add(2, Zeile2);
            Nummer_Zeilen.Add(3, Zeile3);
            Nummer_Zeilen.Add(4, Zeile4);
            Nummer_Zeilen.Add(5, Zeile5);
            Nummer_Zeilen.Add(6, Zeile6);
        }


        // Die Zeilen erste Seite in einer Liste legen
        private void Erste_Zeilen_füllen()
        {
            Zeilen.Add(Zeile1);
            Zeilen.Add(Zeile2);
            Zeilen.Add(Zeile3);
        }

        // Alle-Zeilen-List füllen
        private void AlleZeilen_füllen()
        {
            AlleZeilen.Add(Zeile1);
            AlleZeilen.Add(Zeile2);
            AlleZeilen.Add(Zeile3);
            AlleZeilen.Add(Zeile4);
            AlleZeilen.Add(Zeile5);
            AlleZeilen.Add(Zeile6);
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
            Platte.Enabled = false;
            druckenToolStripMenuItem.Enabled = true;
            automSuchenToolStripMenuItem.Enabled = true;
            referenzfahrtToolStripMenuItem.Enabled = true;
            schriftplattWechselnToolStripMenuItem.Enabled = true;
            dauersuchenToolStripMenuItem.Enabled = true;
            teildruckToolStripMenuItem.Enabled = false;
            seite2ToolStripMenuItem.Enabled = true;
            druckSpeichernToolStripMenuItem.Enabled = false;
            druckToolStripMenuItem.Enabled = false;
            druckLöschenToolStripMenuItem.Enabled = false;
            druckNeuToolStripMenuItem.Enabled= false;
            myport.WriteLine("I1"); //Platte prüfen
            Thread.Sleep(50);
            string platte = myport.ReadExisting();
            Platte.Text = platte;
        }

        // Zum Druckmode Bearbeitenmode
        private void BearbeitenMode()
        {
            this.BackColor = SystemColors.Control;
            StopButton.Visible = false;
            AutoSuchen.Visible = true;
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
            teildruckToolStripMenuItem.Enabled = true;
            seite2ToolStripMenuItem.Enabled = false;
            Platte.Enabled = true;
            Druckstärke.Enabled = true;
            tableLayoutPanel1.Enabled = true;
            druckSpeichernToolStripMenuItem.Enabled = true;
            druckToolStripMenuItem.Enabled = true;
            druckLöschenToolStripMenuItem.Enabled = true;
            druckNeuToolStripMenuItem.Enabled = true;
            Zeile1.Focus();
            Zeile1.Select(0, 0);
        }


        // Open Ports
        private void OpenPorts()
        {
            USB_geht = true;
            try
            {
                myport = new SerialPort();
                myport.BaudRate = 9600;
                myport.PortName = "COM4";
                myport.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Error Arduino");
                USB_geht = false;
            }
            try
            {
                myport2 = new SerialPort();
                myport2.BaudRate = 19200;
                myport2.PortName = "COM5";
                myport2.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Error Motor");
                USB_geht = false;
            }
        }

        // Close Ports
        private void ClosePorts()
        {
            myport.Close();
            myport2.Close();
        }


        /*************************************** Events ***************************************/


        // Wenn man auf eine Zeile ist und auf Enter klict, dann fokusieren auf nächsten Zeile 
        private void Zeile1_KeyDown(object sender, KeyEventArgs e)
        {
            if (!Bearbeiten_Mode)
            {
                return;
            }
            if (e.KeyCode == Keys.Down)
            {
                Zeile2.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 2;
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                Zeile2.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 2;
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
            if (!Bearbeiten_Mode)
            {
                return;
            }
            if (e.KeyCode == Keys.Up)
            {
                Zeile1.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 1;
                return;
            }
            if (e.KeyCode == Keys.Down)
            {
                Zeile3.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 3;
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                Zeile3.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 3;
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
            if (!Bearbeiten_Mode)
            {
                return;
            }
            if (e.KeyCode == Keys.Up)
            {
                Zeile2.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 2;
                return;
            }
            if (e.KeyCode == Keys.Down)
            {
                Zeile4.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 4;
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                Zeile4.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 4;
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
            if (!Bearbeiten_Mode)
            {
                return;
            }
            if (e.KeyCode == Keys.Up)
            {
                Zeile3.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 3;
                return;
            }
            if (e.KeyCode == Keys.Down)
            {
                Zeile5.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 5;
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                Zeile5.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 5;
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
            if (!Bearbeiten_Mode)
            {
                return;
            }
            if (e.KeyCode == Keys.Up)
            {
                Zeile4.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 4;
                return;
            }
            if (e.KeyCode == Keys.Down)
            {
                Zeile6.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 6;
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                Zeile6.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 6;
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
            if (!Bearbeiten_Mode)
            {
                return;
            }
            if (e.KeyCode == Keys.Up)
            {
                Zeile5.Focus();
                e.Handled = e.SuppressKeyPress = true;
                letzte_Zeile = 5;
                return;
            }
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
        

        // Länge nach Änderung in Sperren berechnen
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


        // Ermöglichen Titel ändern Knopf
        private void Satz_Nr_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Satz_Nr.Text)) { AutoSuchen.Visible = true; }
            else if (Bearbeiten_Mode == true && string.IsNullOrEmpty(Satz_Nr.Text)) { AutoSuchen.Visible = false; }
        }


        // Close Ports mit schließen
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (USB_geht) { ClosePorts(); }
        }


        /*************************************** Drucksverfahrenfucnktionen ***************************************/


        // Async Drucken
        private async Task LongRunningOperationAsync()
        {
            await Task.Run(() =>
            {
                Drucken(Zeilen);
            });
        }

        // Zeilen Drucken
        private void Drucken(List<TextBox> Zeilen)
        {
            this.BeginInvoke(new MethodInvoker(() =>
            {
                TextSpeichernOrDrucken.Enabled = false;
                AutoSuchen.Enabled = false;
                FensterWechseln.Enabled = false;
                ReferenzFahrt.Enabled = false;
                SchriftWechseln.Enabled = false;
                panel8.Enabled = false;
            }));
            this.ActiveControl = null;

            if (!Referencefahrt_done)
            {
                MessageBox.Show("Eine Referenzfahrt ist vorher nötig!");
                return;
            }
            this.BeginInvoke(new MethodInvoker(() =>
            {
                if (Platte_prüfen().Trim() != (SchriftGröße.SelectedIndex + 1).ToString())
                {
                    MessageBox.Show("Bitte die richtige Platte benutzen!");
                    return;
                }
            }));
            Thread.Sleep(100);
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
            if (!Check_Höhe_Hier())
            {
                MessageBox.Show("Bitte Höhe wählen!");
                return;
            }
            if (!Check_Zeilen_Länge())
            {
                return;
            }

            myport.WriteLine("O31"); // Pumpe einschalten
            foreach (TextBox Zeile in Zeilen)
            {
                if (string.IsNullOrEmpty(Zeile.Text)) continue;
                string Zum_Drucken = Zeile.Text.Trim();
                int Sprr = (!string.IsNullOrEmpty(Zeilen_Sperren[Zeile].Text)) ? Int32.Parse(Zeilen_Sperren[Zeile].Text) : 0;
                Tisch_init = 8380 - (Int32.Parse(Zeilen_Höhe[Zeile].Text) * 40);
                int Mitte = (int)((float)(AbstandVonAussen.Value - 32) + float.Parse(Zeilen_Länge[Längste_Zeile].Text) / 2);
                int ABStVA = (Mitte - ((int)(float.Parse(Zeilen_Länge[Zeile].Text) / 2))) * 40;
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
                                SonderZeichen_Drucken(Zeile, Zum_Drucken[x - 1], "0", "20", ABStVA.ToString());
                                M3 = 0;
                                continue;
                            }
                            else if (Abstände.Contains(Zum_Drucken[x - 1]))
                            {
                                continue;
                            }
                            int höhe = (Int32.Parse(Zeilen_Höhe[Zeile].Text) * 40);
                            Motoren_1A_2A_3R_4A_5R(höhe, Zum_Drucken[x - 1].ToString(), M3.ToString(), "20", "1", ABStVA);
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
                                SonderZeichen_Drucken(Zeile, Zum_Drucken[x - 1], M3.ToString(), M5.ToString());
                                M3 = 0;
                                AnzahlVonABSt = 0;
                                continue;
                            }
                            int höhe = (Int32.Parse(Zeilen_Höhe[Zeile].Text) * 40);
                            Motoren_1A_2A_3R_4A_5R(höhe, Zum_Drucken[x - 1].ToString(), M3.ToString(), M5.ToString());
                        }
                        Motoren_stehen();
                        if (weiter == false)
                        {
                            DialogResult dr = MessageBox.Show("weiter?", "", MessageBoxButtons.YesNo);
                            if (dr == DialogResult.Yes) { weiter = true; }
                            else
                            {
                                myport.WriteLine("O30");
                                StartPunkt();
                                weiter = true;
                                this.BeginInvoke(new MethodInvoker(() =>
                                {
                                    TextSpeichernOrDrucken.Enabled = true;
                                    AutoSuchen.Enabled = true;
                                    FensterWechseln.Enabled = true;
                                    ReferenzFahrt.Enabled = true;
                                    SchriftWechseln.Enabled = true;
                                    EmptyTheFields();
                                }));
                                this.ActiveControl = null;
                                return;
                            };
                        }
                        Stempel_ab(Zum_Drucken[x - 1]);
                        Stempel_auf();
                        bool Stmp = true;
                        while (Stmp)
                        {
                            Stmp = !Stempel_prüfen();
                        }
                        Trennung();
                        M3 = 0;
                        AnzahlVonABSt = 0;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
                Motore_5_Drehen_Relativ("20");
                Tisch_init = 0;
                if (Satz_Nr.Text != "") { DruckÄndern_Neu("Druckdatei", Zeile); }
                this.BeginInvoke(new MethodInvoker(() =>
                {
                    UngedruckteZeilenBerechnen();
                }));
            }
            StartPunkt();
            myport.WriteLine("O30"); // Pumpe aus

            this.BeginInvoke(new MethodInvoker(() =>
            {
                TextSpeichernOrDrucken.Enabled = true;
                AutoSuchen.Enabled = true;
                FensterWechseln.Enabled = true;
                ReferenzFahrt.Enabled = true;
                SchriftWechseln.Enabled = true;
                panel8.Enabled = true;
            }));

            this.ActiveControl = null;
            Druck_done = true;
        }

        // Die Sonderzeichen Drucken
        private void SonderZeichen_Drucken(TextBox Zeile, char v, string M3 = "0", string M5 = "0", string abst = "0")
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
            int höhe = (Int32.Parse(Zeilen_Höhe[Zeile].Text) * 40);
            Motoren_1A_2A_3R_4A_5R(höhe, b.ToString(), M3, M5, "1", Int32.Parse(abst));
            Motoren_stehen();
            Stempel_ab(b);
            Stempel_auf();
            bool Stmp = true;
            while (Stmp)
            {
                Stmp = !Stempel_prüfen();
            }
            Trennung();
            string sdr = Tisch_Pos_bringen(v.ToString());
            Motoren_1A_2A_3R_4A_5R(höhe, Sonder.ToString(), "0", M5, sdr);
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

        // Tisch Position aus Datenbank aufrunfen
        private string Tisch_Pos_bringen(string x)
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
            string strC;
            if (c == '\\') { strC = "\\\\"; }
            else if (c == '\'') { strC = "\\\'"; }
            else { strC = c.ToString(); }
            int druck = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select Druck from Tabelle" + Schriftgröße + " where Zeichen = '" + strC + "' COLLATE utf8mb4_bin;";
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
            Thread.Sleep((int)Druckzeit.Value * 20);
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
            Thread.Sleep(50);
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

        // Motoren Drehen
        private void Motoren_1A_2A_3R_4A_5R(int Hohe, string _124, string _3, string _5, string t = "1", int ab = 0)
        {
            Tisch_init = Hohe;
            int XPos = 0, YPos = 0, TischPos = 0;
            string strC;
            if (_124 == "\\") { strC = "\\\\"; }
            else if (_124 == "\'") { strC = "\\\'"; }
            else { strC = _124.ToString(); }
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select XPos, YPos, TischPos from tabelle" + Schriftgröße + " where Zeichen = '" + strC + "' COLLATE utf8mb4_bin";
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
                Thread.Sleep(20);
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
                int Vier = (Int32.Parse(t) + Tisch_init);
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
                Thread.Sleep(20);
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

            if (ab != 0 || _3 != "0")
            {
                message = "#" + (char)3 + "A\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
            }


            if (_5 != "0")
            {
                message = "#" + (char)5 + "A\r";
                myport2.WriteLine(message);
                Thread.Sleep(20);
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
            Thread.Sleep(50);
        }

        // Zum Startpunkt fahren (nach Referenzfahrt)
        private void StartPunkt(int wait = 100)
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
            Thread.Sleep(wait);
            myport2.ReadExisting();
            Daten = "#1A\r";        // starten
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


        /*************************************** Apps-Operationsfucnktionen ***************************************/


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
            AktuellDatei = SaveIn;
        }

        // Die aktuellen Druck von Druckdatei oder Archiv löschen
        private void DruckLöschen(string AktuellDatei)
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "delete from " + AktuellDatei + " where nr = " + Satz_Nr.Text + ";";
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

        // Änderungen nach dem Drucken speichern 
        private void DruckÄndern_Neu(string AktuellDatei, TextBox gedruckt = null)
        {
            if (CheckEveryZeileHasHöhe() == false) { return; }
            CreateNullForDatenbank();

            DateTime dt1 = DateTime.Now;
            string Datum = "\"" + dt1.ToString() + "\"";
            int breite = -1;
            int farbe = -1;
            string Satz_Nummer = "";
            string Ged = "";
            string schrift = "";
            this.BeginInvoke(new MethodInvoker(() =>
            {
                breite = (Breite.Text == "") ? -1 : BreiteToDatenBank(Int32.Parse(Breite.Text));
                farbe = (FarbeEingabe.Text == "") ? -1 : FarbeEingabe.SelectedIndex + 1;
                Satz_Nummer = (Satz_Nr.Text == "") ? Letzte_Satz_Nummer() : Satz_Nr.Text;
                Ged = "\"" + CreateGedLangDruck(gedruckt, Satz_Nummer) + "\"";
                schrift = (SchriftGröße.SelectedIndex + 1).ToString();
            }));
            Thread.Sleep(100);
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
                    " where nr = " + Satz_Nummer + ";";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            conn.Close();

            this.BeginInvoke(new MethodInvoker(() =>
            {
                UngedruckteZeilenBerechnen();
            }));
        }

        // Änderungen speichern
        private void DruckÄndern(string AktuellDatei, TextBox gedruckt = null)
        {
            if (CheckEveryZeileHasHöhe() == false) { return; }
            CreateNullForDatenbank();

            DateTime dt1 = DateTime.Now;
            string Datum = "\"" + dt1.ToString() + "\"";
            int breite = (Breite.Text == "") ? -1 : BreiteToDatenBank(Int32.Parse(Breite.Text));
            int farbe = (FarbeEingabe.Text == "") ? -1 : FarbeEingabe.SelectedIndex + 1;
            string Satz_Nummer = (Satz_Nr.Text == "") ? Letzte_Satz_Nummer() : Satz_Nr.Text;
            string Ged = "\"" + CreateGedLangDruck(gedruckt, Satz_Nummer) + "\"";
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
                    " where nr = " + Satz_Nummer + ";";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Kein Druck zum Ändern!");
            }
            conn.Close();
            UngedruckteZeilenBerechnen();
        }

        // Druck aus Datenbank aufrufen
        private void DruckAufrufen(string SelectDruck, string auto = " and Ged like \"L%\"")
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            string Ged;
            try
            {
                conn.Open();
                string sql = "select * from " + AktuellDatei + " where nr = " + SelectDruck + auto;
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
                    Ged = rdr["Ged"].ToString();
                    BackGroundFarben(Ged);
                }
                rdr.Close();
                CheckLängsteZeile(LaengeZeile1);
            }

            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                EmptyTheFields();
            }
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
                //int AnzahZeichen = x.Length - 1;
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

        // Breite von Zeichen aufrufen
        private int Zeichen_Breite(char c)
        {
            string strC;
            if (c == '\\') { strC = "\\\\"; }
            else if (c == '\'') { strC = "\\\'"; }
            else { strC = c.ToString(); }
            int y = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select Breite from Tabelle" + Schriftgröße + " where Zeichen = '" + strC + "' COLLATE utf8mb4_bin;";
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

        // die Nummer letztes Drucks geben
        private string Letzte_Satz_Nummer()
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select max(nr) from " + AktuellDatei + " where Ged like \"L%\"";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    return rdr[0].ToString();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            return null;
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

        // 'Ged'-Spalte erstellen
        private string CreateGedLangDruck(TextBox gedruckt = null, string satz_Nummer = "")
        {
            string Ged = "";
            string SchriftChar = (SchriftGröße.SelectedIndex + 1).ToString();
            if (gedruckt == null)
            {
                Ged = "L";
                Ged = (Zeile1.Text == "") ? Ged + "0" : Ged + SchriftChar;
                Ged = (Zeile2.Text == "") ? Ged + "0" : Ged + SchriftChar;
                Ged = (Zeile3.Text == "") ? Ged + "0" : Ged + SchriftChar;
                Ged = (Zeile4.Text == "") ? Ged + "0" : Ged + SchriftChar;
                Ged = (Zeile5.Text == "") ? Ged + "0" : Ged + SchriftChar;
                Ged = (Zeile6.Text == "") ? Ged + "0" : Ged + SchriftChar;
            }
            else 
            {
                string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
                MySqlConnection conn = new MySqlConnection(connStr);
                try
                {
                    conn.Open();
                    string sql = "select Ged from " + AktuellDatei + " where nr = " + satz_Nummer;
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    MySqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        Ged = rdr[0].ToString();
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                int key = Nummer_Zeilen.FirstOrDefault(x => x.Value == gedruckt).Key;
                char[] charArry = Ged.ToCharArray();
                charArry[key] = 'x';
                string new_Ged = new string(charArry);
                return new_Ged;
            }
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
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5, Sperren6, BandNr, Satz_Nr, LaengeZeile1
                , LaengeZeile2, LaengeZeile3, LaengeZeile4, LaengeZeile5, LaengeZeile6
            };
            foreach (TextBox t in textBoxes)
            {
                t.Text = "";
                t.BackColor = Color.White;
            }
            AbstandVonAussen.Value = 100;
            Breite.SelectedIndex = -1;
            FarbeEingabe.SelectedIndex = -1;
            letzte_Zeile = 1;
        }

        // Der Hintergrund mit Pink farbigen
        private void BackGroundFarben(String Ged)
        {
            if (!Ged.Contains("x"))
            {
                foreach (TextBox t in AlleZeilen)
                {
                    t.BackColor = Color.White;
                }
            }
            else
            {
                for (int i = 1; i < 7; i++)
                {
                    if (Ged[i] == 'x')
                    {
                        Nummer_Zeilen[i].BackColor= Color.LightPink;
                    }
                    else
                    {
                        Nummer_Zeilen[i].BackColor = Color.White;
                    }
                }
            }
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
                string sql1 = "select sum(Length(ged) - length(replace(ged, \"1\", \"\"))) from druckdatei WHERE ged LIKE \"L%\";";
                MySqlCommand cmd1 = new MySqlCommand(sql1, conn);
                object result1 = cmd1.ExecuteScalar();
                if (result1 != null)
                {
                    p1 = Convert.ToInt32(result1);
                }
                string sql2 = "select sum(Length(ged) - length(replace(ged, \"2\", \"\"))) from druckdatei WHERE ged LIKE \"L%\";";
                MySqlCommand cmd2 = new MySqlCommand(sql2, conn);
                object result2 = cmd2.ExecuteScalar();
                if (result2 != null)
                {
                    p2 = Convert.ToInt32(result2);
                }
                string sql3 = "select sum(Length(ged) - length(replace(ged, \"3\", \"\"))) from druckdatei WHERE ged LIKE \"L%\";";
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
        private bool HoheSeite1 ()
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
            else 
            {
                MessageBox.Show("Die Zeilen sind entweder leer oder falsch verteilt!");
                return false;
            }
            try
            {
                conn.Open();
                string sql1 = "Select " + scrift + band + _zeile + " from Zeilenhohe where Breite = " + breite + ";";
                MySqlCommand cmd1 = new MySqlCommand(sql1, conn);
                MySqlDataReader rdr1 = cmd1.ExecuteReader();
                while (rdr1.Read())
                {
                    string rd = rdr1[0].ToString();
                    if (string.IsNullOrEmpty(rd))
                    { 
                        HoheSeite1Leeren();
                        MessageBox.Show("Der Text passt nicht auf die gewählte Schleife!");
                        return false;
                    }
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
            return true;
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
                    if (string.IsNullOrEmpty(rd2)) { HoheSeite2Leeren(); MessageBox.Show("Der Text passt nicht auf die gewählte Schleife!"); }
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


        /*************************************** Checken ***************************************/


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
            Längste_Zeile = Zeilen_Aus_Länge[Höhste];
        }

        // kontroliere ob alle Zeilen eine Höhe haben
        private bool Check_Höhe_Hier()
        {
            foreach (TextBox i in Zeilen)
            {
                string Zum_Drucken = i.Text.Trim();
                if (!string.IsNullOrEmpty(Zum_Drucken) && string.IsNullOrEmpty(Zeilen_Höhe[i].Text)) return false;
            }
            return true;
        }

        // kontroliere ob alle Zeilen kurzer als 71 cm
        private bool Check_Zeilen_Länge()
        {
            foreach (TextBox t in AlleZeilen)
            {
                if (Zeilen_Länge[t].Text == "") continue;
                if (Convert.ToDouble(Zeilen_Länge[t].Text) >= 670)
                {
                    MessageBox.Show("Text wird zu lang!");
                    return false;
                }
            }
            return true;
        }

        // Überprüfen ob jede Zeile ihre "Höhe" hat
        private bool CheckEveryZeileHasHöhe()
        {
            if (!string.IsNullOrEmpty(Zeile1.Text) && string.IsNullOrEmpty(HoeheZeile1.Text))
            {
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


        /*************************************** Buttonsklicken ***************************************/


        private async void TextSpeichernOrDrucken_Click(object sender, EventArgs e)
        {
            if (!CheckEveryZeileHasHöhe()) { return; }
            if (!Check_Zeilen_Länge()) { return; }
            CreateNullForDatenbank();
            if (Bearbeiten_Mode)
            {
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
                    Zeile1.Focus();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
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
                if (!Check_Höhe_Hier())
                {
                    MessageBox.Show("Bitte Höhe wählen!");
                    return;
                }
                if (!Check_Zeilen_Länge())
                {
                    return;
                }
                //if (Satz_Nr.Text == "")
                //{
                //    try
                //    {
                //        DruckSpeichern("Druckdatei");
                //    }
                //    catch (Exception ex)
                //    {
                //        MessageBox.Show(ex.ToString());
                //    }
                //}
                Zeilen.Clear();
                Zeilen.AddRange(new List<TextBox> { Zeile1, Zeile2, Zeile3 });
                //Drucken(Zeilen);
                await LongRunningOperationAsync();

                if (!Druck_done) return;
                if ((Zeile4.Text.Length > 0) || (Zeile5.Text.Length > 0) || (Zeile6.Text.Length > 0))
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Zweiter Flügel einspannen?", "", MessageBoxButtons.OKCancel);
                    if (dr == DialogResult.OK)
                    {
                        Zeilen.Clear();
                        Zeilen.AddRange(new List<TextBox> { Zeile4, Zeile5, Zeile6 });
                        //Drucken(Zeilen);
                        await LongRunningOperationAsync();

                        Zeilen.Clear();
                        Zeilen.AddRange(new List<TextBox> { Zeile1, Zeile2, Zeile3 });
                    }
                    else
                    {
                        Druck_done = false;
                        EmptyTheFields();
                        return;
                    }
                }
                EmptyTheFields();
                Druck_done = false;
                Nexter_Druck = true;
            }
        }

        // Wechseln zwischen Druckmode und Bearbeitenmode
        private void FensterWechseln_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode)
            {
                OpenPorts();
                if (!USB_geht) return;
                if (!Stempel_prüfen())
                {
                    MessageBox.Show("Mit Motoren verbinden!");
                    ClosePorts();
                    return;
                }

                Bearbeiten_Mode = false;
                DruckMode();
            }
            else
            {
                Bearbeiten_Mode = true;
                BearbeitenMode();
                ClosePorts();
            }
        }

        // Das Programm verlassen
        private void ProgrammBeendenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
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
            }
            else
            {
                return;
            }
        }

        private void druckenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextSpeichernOrDrucken.PerformClick();
        }

        private void schriftplattWechselnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!Referencefahrt_done)
            {
                MessageBox.Show("Eine Referenzfahrt ist vorher nötig!");
                TextSpeichernOrDrucken.Enabled = true;
                return;
            }
            if (!Bandhalter_prüfen())
            {
                MessageBox.Show("Bandhalter schließen!");
                return;
            }
            MessageBox.Show("Heizung herunterdrehen!");

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

            Daten = "#1s2800\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)2 + "s16160\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            myport2.ReadExisting();
            Daten = "#" + (char)3 + "s4000\r";        //Startposition setzen
            myport2.WriteLine(Daten);
            Thread.Sleep(50);
            myport2.ReadExisting();
            Daten = "#" + (char)4 + "s400\r";        //Startposition setzen
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

            Motoren_stehen();

            DialogResult dr;
            dr = MessageBox.Show("Züruck?", "", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                StartPunkt(1000);
                myport.WriteLine("I1"); //Platte prüfen
                Thread.Sleep(50);
                string platte = myport.ReadExisting();
                Platte.Text = platte;
                if (platte.Trim() != "1" && platte.Trim() != "2" && platte.Trim() != "3")
                { 
                    MessageBox.Show("Achtung: Falsche Platte oder Falsche Richtung!");
                }
                MessageBox.Show("Die Heizung auf 120C° stellen!");
            }
            else
            {
                return;
            }
        }

        // letzter Druck in Druckmode aufrufen
        private void automSuchenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Farbe = "";
            if (Radio_gold.Checked) Farbe = "1";
            else if (Radio_Schwarz.Checked) Farbe = "2";
            else if (Radio_Silber.Checked) { Farbe = "3"; }

            string Add_Farbe = " and Farbe = " + Farbe;
            if (Radio_Alle.Checked) Add_Farbe = "";
            AktuellDatei = "druckdatei";
            SchriftGröße.SelectedIndex = Int32.Parse(Platte.Text) - 1;
            DruckAufrufen("(select max(nr) from druckdatei where Ged Like \"%" +
                (SchriftGröße.SelectedIndex + 1).ToString() + "%\"" + Add_Farbe + ")");
            AktuellDruck = Satz_Nr.Text;
        }

        // Dauersuchen
        private async void dauersuchenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EmptyTheFields();
            while (true)
            {
                automSuchenToolStripMenuItem.PerformClick();
                if (Satz_Nr.Text == "")
                {
                    MessageBox.Show("Alle Schleifen sind gedruckt!");
                    return;
                }
                while (!Nexter_Druck)
                {
                    await Task.Delay(500);
                }
                Nexter_Druck = false;
                DialogResult dr;
                //dr = MessageBox.Show("Nächster Druck drucken?", "", MessageBoxButtons.YesNo);
                //if (dr == DialogResult.Yes)
                //{
                //    continue;
                //}
                //else
                //{
                //    return;
                //}
            }

        }

        private void teildruckToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FensterWechseln.PerformClick();
            TextSpeichernOrDrucken.Enabled = false;
            this.ActiveControl = null;

            if (!Referencefahrt_done)
            {
                MessageBox.Show("Eine Referenzfahrt ist vorher nötig!");
                TextSpeichernOrDrucken.Enabled = true;

                return;
            }
            if (Platte_prüfen().Trim() != (SchriftGröße.SelectedIndex + 1).ToString())
            {
                MessageBox.Show("Bitte die richtige Platte benutzen!");
                TextSpeichernOrDrucken.Enabled = true;
                return;
            }
            if (!Bandhalter_prüfen())
            {
                MessageBox.Show("Bandhalter schließen!");
                TextSpeichernOrDrucken.Enabled = true;
                return;
            }
            if (Schriftgröße == -1)
            {
                MessageBox.Show("Bitte Schriftgröße wählen!");
                TextSpeichernOrDrucken.Enabled = true;
                return;
            }
            if (!Check_Höhe_Hier())
            {
                MessageBox.Show("Bitte Höhe wählen!");
                TextSpeichernOrDrucken.Enabled = true;
                return;
            }
            if (!Check_Zeilen_Länge())
            {
                return;
            }
            foreach (TextBox t in AlleZeilen)
            {
                if (t.SelectedText.Length < 1) continue;
                else
                {
                    myport.WriteLine("O31");
                    int Sprr = (!string.IsNullOrEmpty(Zeilen_Sperren[t].Text)) ? Int32.Parse(Zeilen_Sperren[t].Text) : 0;
                    Tisch_init = 8380 - (Int32.Parse(Zeilen_Höhe[t].Text) * 40);
                    int Mitte = (int)((float)(AbstandVonAussen.Value - 32) + float.Parse(Zeilen_Länge[Längste_Zeile].Text) / 2);
                    int ABStVA = (Mitte - ((int)(float.Parse(Zeilen_Länge[t].Text) / 2))) * 40;

                    int M3 = 0;
                    bool skip = false;
                    for (int x = t.TextLength; x > 0; x--)
                    {
                        if (x - 1 < t.SelectionStart)
                        {
                            StartPunkt();
                            myport.WriteLine("O30"); // Pumpe aus
                            TextSpeichernOrDrucken.Enabled = true;
                            this.ActiveControl = null;
                            EmptyTheFields();
                            return;
                        }
                        int M5;
                        try
                        {
                            if (x == t.Text.Length)
                            {
                                if (x > t.SelectionStart + t.SelectionLength)
                                {
                                    skip = true;
                                    continue;
                                }
                                if (SonderZeichen.Contains(t.Text[x - 1]))
                                {
                                    SonderZeichen_Drucken(t, t.Text[x - 1], "0", "20", ABStVA.ToString());
                                    M3 = 0;
                                    continue;
                                }
                                else if (Abstände.Contains(t.Text[x - 1]))
                                {
                                    continue;
                                }
                                int höhe = (Int32.Parse(Zeilen_Höhe[t].Text) * 40);
                                Motoren_1A_2A_3R_4A_5R(höhe, t.Text[x - 1].ToString(), M3.ToString(), "20", "1", ABStVA);
                            }
                            else
                            {
                                if (Abstände.Contains(t.Text[x - 1]))
                                {
                                    M3 += Convert.ToInt32((Zeichen_Breite(t.Text[x]) + Zeichen_Breite(t.Text[x - 1])) * 2);
                                    continue;
                                }
                                if (x > t.SelectionStart + t.SelectionLength)
                                {
                                    M3 += Convert.ToInt32((Zeichen_Breite(t.Text[x]) + Zeichen_Breite(t.Text[x - 1])) * 2 + 20);
                                    M3 += Sprr * 20;
                                    continue;
                                }
                                M3 += Convert.ToInt32((Zeichen_Breite(t.Text[x]) + Zeichen_Breite(t.Text[x - 1])) * 2 + 20);
                                M3 += Sprr * 20;
                                M5 = Convert.ToInt32((Zeichen_Breite(t.Text[x]) + Zeichen_Breite(t.Text[x - 1])) * 0.12);
                                if (SonderZeichen.Contains(t.Text[x - 1]))
                                {
                                    if (skip)
                                    {
                                        SonderZeichen_Drucken(t, t.Text[x - 1], M3.ToString(), M5.ToString(), (ABStVA + M3).ToString());
                                        skip = false;
                                    }
                                    else
                                    {
                                        SonderZeichen_Drucken(t, t.Text[x - 1], M3.ToString(), M5.ToString());
                                    }
                                    M3 = 0;
                                    continue;
                                }
                                if (skip == true)
                                {
                                    int höhe = (Int32.Parse(Zeilen_Höhe[t].Text) * 40);
                                    Motoren_1A_2A_3R_4A_5R(höhe, t.Text[x - 1].ToString(), M3.ToString(), M5.ToString(), "1", ABStVA + M3);
                                    skip = false;
                                }
                                else
                                {
                                    int höhe = (Int32.Parse(Zeilen_Höhe[t].Text) * 40);
                                    Motoren_1A_2A_3R_4A_5R(höhe, t.Text[x - 1].ToString(), M3.ToString(), M5.ToString());
                                }
                            }
                            Motoren_stehen();
                            Stempel_ab(t.Text[x - 1]);
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
                        Motore_5_Drehen_Relativ("20");
                    }
                    StartPunkt();
                    myport.WriteLine("O30"); // Pumpe aus
                    TextSpeichernOrDrucken.Enabled = true;
                    this.ActiveControl = null;
                    EmptyTheFields();
                    return;
                }
            }
            MessageBox.Show("Es wird Kein Text gewählt!");
            FensterWechseln.PerformClick();
        }

        private async void seite2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("Zweiter Flügel einspannen?", "", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                if ((Zeile4.Text.Length > 0) || (Zeile5.Text.Length > 0) || (Zeile6.Text.Length > 0))
                {
                    Zeilen.Clear();
                    Zeilen.AddRange(new List<TextBox> { Zeile4, Zeile5, Zeile6 });
                    await LongRunningOperationAsync();
                    Zeilen.Clear();
                    Zeilen.AddRange(new List<TextBox> { Zeile1, Zeile2, Zeile3 });
                    EmptyTheFields();
                    Nexter_Druck = true;
                }
            }
            else
            {
                return;
            }
        }

        private void DrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AktuellDatei = "Druckdatei";
            DruckAufrufen("(select max(nr) from druckdatei)");
            AktuellDruck = Satz_Nr.Text;
            label3.Text = AktuellDatei;
        }

        private void LetzterDruckArchivToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AktuellDatei = "Archiv";
            DruckAufrufen("(select max(nr) from Archiv)");
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
                    DruckAufrufen("(select max(nr) from " + AktuellDatei + " where nr < " + AktuellDruck + ")");
                    AktuellDruck = Satz_Nr.Text;
                }
                catch { return; }
            }
        }

        private void NächsterDruck_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(AktuellDruck))
            {
                DruckAufrufen("(select min(nr) from druckdatei)");
                AktuellDruck = Satz_Nr.Text;
            }
            else
            {
                try
                {
                    DruckAufrufen("(select min(nr) from " + AktuellDatei + " where nr > " + AktuellDruck + ")");
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

        private void AutoSuchen_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode)
            {
                if (Satz_Nr.Text == "") { MessageBox.Show("Keinen Druck zum Ändern!"); return; }
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
            else
            {
                SchriftGröße.SelectedIndex = Int32.Parse(Platte.Text) - 1;
                DruckAufrufen("(select max(nr) from druckdatei where Ged Like \"%" +
                    (SchriftGröße.SelectedIndex + 1).ToString() + "%\")");
                AktuellDruck = Satz_Nr.Text;
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
            if (string.IsNullOrEmpty(gedd)) return;
            BackGroundFarben(gedd);
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
                if (HoheSeite1()) HoheSeite2();
                if (Zeile1.Text == "") { HoeheZeile1.Text = ""; }
                if (Zeile2.Text == "") { HoeheZeile2.Text = ""; }
                if (Zeile3.Text == "") { HoeheZeile3.Text = ""; }
                if (Zeile4.Text == "") { HoeheZeile4.Text = ""; }
                if (Zeile5.Text == "") { HoeheZeile5.Text = ""; }
                if (Zeile6.Text == "") { HoeheZeile6.Text = ""; }
                if (Breite.SelectedIndex == 0) { AbstandVonAussen.Value = 42; }
                else if (Breite.SelectedIndex == 1) { AbstandVonAussen.Value = 42; }
                else if (Breite.SelectedIndex == 2) { AbstandVonAussen.Value = 45; }
                else if (Breite.SelectedIndex == 3) { AbstandVonAussen.Value = 60; }
                else { AbstandVonAussen.Value = 100; }
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            weiter = false;
        }

        // Versuch
        private void PumpeAus_Click(object sender, EventArgs e)
        {
            //Referencefahrt_done = true;
            myport.WriteLine("O30"); // Pumpe aus
        }
    }
}
