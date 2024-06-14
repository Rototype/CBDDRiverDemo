using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Rototype.CBD;
using System.IO.Ports;

namespace CBD1000_Demo
{
    public partial class Form1 : Form
    {
        readonly SerialPort rbport = new SerialPort();
        readonly Random rnd = new Random();

        readonly CbdDriver Cbd1000 = new CbdDriver();
        string BookletLayoutFile = "";

        int BookletStartNo = 0;
        int BCCounter = 0;

        public Form1()
        {
            InitializeComponent();
            toolStripComboBox4.SelectedIndex = 1;
            Cbd1000.CurrentModel = (CbdDriver.CbdModel)toolStripComboBox4.SelectedIndex;

            toolStripStatusLabel1.Text = "Close";
            toolStripStatusLabel2.Text = BookletLayoutFile;
            richTextBox1.BackColor = Color.Gray;

            Cbd1000.CompleteEvent += new EventHandler<CbdDriver.CompleteEventArgs>(Cbd1000_CompleteEvent);
            Cbd1000.ConnectedEvent += new EventHandler<EventArgs>(Cbd1000_ConnectedEvent);
            Cbd1000.DisconnectedEvent += new EventHandler<EventArgs>(Cbd1000_DisconnectedEvent);
            Cbd1000.ErrorEvent += new EventHandler<CbdDriver.ErrorEventArgs>(Cbd1000_ErrorEvent);
            Cbd1000.ForceDisconnectionEvent += new EventHandler<EventArgs>(Cbd1000_ForceDisconnectionEvent);

            char [] sep = { ':' };

            rbport.DataReceived += (o, i) =>
            {
                string rbstr = rbport.ReadLine();
                var waitedBC = BookletStartNo + BCCounter;
                BCCounter++;
                var ss = rbstr.Split(sep);
                Int32.TryParse(ss[1], out int n);
                if (ss.Length == 2)
                {
                    string reply = "";
                    if (waitedBC == n)
                    {
                        reply = string.Format("Verified BC:{0}", n);
                        richTextBox1.Invoke(new Action(() => LogText(reply, Color.Green)));
                    }
                    else
                    {
                        reply = string.Format("[!] Bad BC reading, read {0} instead of {1}", n, waitedBC);
                        richTextBox1.Invoke(new Action(() => LogText(reply, Color.Red)));
                        Cbd1000.Close();
                    }
                }
            };


        }

        private void C(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        CbdDriver.Booklet BuildBook(int start)
        {
            var Booklet = new CbdDriver.Booklet
            {

                // This option should not be enabled on documents having a height < 80mm
                NoBookletExpose = noBookletExposeToolStripMenuItem.Checked,
                NoRibbon = noRibbonToolStripMenuItem.Checked
            };


            int chequeno = int.Parse(toolStripComboBox2.Text);

            bool addcovers = enableCoversToolStripMenuItem.Checked;


            if (addcovers && chequeno > 6) 
                Booklet.Sheets.Add(new CbdDriver.Sheet(CbdDriver.SheetType.FrontCover));
            for (int f = 0; f != chequeno; f++)
            {
                CbdDriver.Sheet cheque = new CbdDriver.Sheet(CbdDriver.SheetType.Cheque);

                string number = (start + f).ToString("00000000");
                string DateString = DateTime.Now.ToString("dddd, dd MMMM yyyy - HH:mm");
                string TText = "";
                if (enableCodelinePrintingToolStripMenuItem.Checked)
                    TText = new string(' ', 30) + "1234567890 " + number;
                cheque.Barcode = new CbdDriver.BarcodeDefinition
                {
                    Number = 1,
                    Version = 10,
                    Zoom = 5,
                    Text = DateString + " #" + number
                };
                
                cheque.Fields.Add(new CbdDriver.Field
                    {
                     
                        Number = 0,
                        Text = TText
                    }
                );
                cheque.Fields.Add(new CbdDriver.Field
                    {
                        Number = 1,
                        Text = "number:" + number
                }
                );
                cheque.Fields.Add(new CbdDriver.Field
                    {
                        Number = 2,
                        Text = "Second variable field."
                }
                );
                cheque.Fields.Add(new CbdDriver.Field
                    {
                        Number = 3,
                    Text = "Third variable field."
                }
                );
                cheque.Fields.Add(new CbdDriver.Field
                    {
                        Number = 4,
                        Text = "Fourth variable field."
                    }
                );
                Booklet.Sheets.Add(cheque);

                if (addcovers && chequeno > 6)
                    if (f == chequeno - 3)
                        Booklet.Sheets.Add(new CbdDriver.Sheet(CbdDriver.SheetType.OrderForm));
            }
            if (addcovers && chequeno > 6)
                Booklet.Sheets.Add(new CbdDriver.Sheet(CbdDriver.SheetType.BackCover));

            CbdDriver.Booklet newbook = null;

            if ( CbdDriver.Booklet.Save(Booklet, "StoredBook.out"))
                MessageBox.Show("Error while saving booklet", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if ((newbook = CbdDriver.Booklet.Load("StoredBook.out")) == null)
                MessageBox.Show("Error while loading booklet", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            CbdDriver.Booklet.Save(Booklet, "StoredBook.out");

            return newbook;
        }


        private void ScanForSerialPort()
        {
            var ports =
                System.IO.Ports.SerialPort.GetPortNames()
                .OrderBy(l => l).Reverse();
            toolStripComboBox1.Items.Clear();
            foreach (var port in ports)
                toolStripComboBox1.Items.Add(port);
            if (toolStripComboBox1.Items.Count > 0)
                toolStripComboBox1.Text = toolStripComboBox1.Items[0].ToString();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

            ScanForSerialPort();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Cbd1000.CompleteEvent -= new EventHandler<CbdDriver.CompleteEventArgs>(Cbd1000_CompleteEvent);
            Cbd1000.ConnectedEvent -= new EventHandler<EventArgs>(Cbd1000_ConnectedEvent);
            Cbd1000.DisconnectedEvent -= new EventHandler<EventArgs>(Cbd1000_DisconnectedEvent);
            Cbd1000.ErrorEvent -= new EventHandler<CbdDriver.ErrorEventArgs>(Cbd1000_ErrorEvent);
            Cbd1000.ForceDisconnectionEvent -= new EventHandler<EventArgs>(Cbd1000_ForceDisconnectionEvent);
            Cbd1000.Close();
        }



        void RichTextLogText(string msg, Color color)
        {
            if( msg.StartsWith("Disconnected"))
                richTextBox1.BackColor = Color.Gray;
            if (msg.StartsWith("Connected"))
                richTextBox1.BackColor = Color.White;


            int start = richTextBox1.TextLength;
            richTextBox1.AppendText(msg);
            int end = richTextBox1.TextLength;
            richTextBox1.Select(start, end - start);
            richTextBox1.SelectionColor = color;
            richTextBox1.SelectionLength = 0;
            richTextBox1.ScrollToCaret();
        }

        void LogText(string msg, Color color)
        {

            msg += "\r\n";
            if (richTextBox1.InvokeRequired)
                Invoke(new Action(() => RichTextLogText(msg, color)));
            else
                RichTextLogText(msg, color);
        }

        void Cbd1000_ErrorEvent(object sender, CbdDriver.ErrorEventArgs e)
        {
            Color color = Color.Black;
            switch (e.Code[0])
            {
                case 'E':
                    color = Color.Red;
                    break;
                case 'F':
                    color = Color.Purple;
                    break;
                case 'J':
                    color = Color.Green;
                    break;
                case 'W':
                    color = Color.Black;
                    break;
            }
            LogText( e.Code + "->" + e.Description, color);
            //MessageBox.Show(e.Description, e.Code, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        void Cbd1000_ForceDisconnectionEvent(object sender, EventArgs e)
        {
            LogText("Disconnected by timeout", Color.Blue);
        }

        void Cbd1000_DisconnectedEvent(object sender, EventArgs e)
        {
            LogText("Disconnected", Color.Blue);
        }

        void Cbd1000_ConnectedEvent(object sender, EventArgs e)
        {
            LogText("Connected", Color.Blue);
        }

        void Cbd1000_CompleteEvent(object sender, CbdDriver.CompleteEventArgs e)
        {
            if ( e.Success)
                LogText("Complete successfully", Color.Brown);
            else
                LogText("Complete with errors", Color.Brown);
        }


       private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("A sample application to test the CBDDriver library", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void printBookletToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogText("PrintBooklet()", Color.Brown);

            BookletStartNo = (rnd.Next(5, 9000000));
            BCCounter = 0;
            CbdDriver.Booklet Booklet = BuildBook(BookletStartNo);
            if (Cbd1000.PrintBooklet(Booklet))
                MessageBox.Show( "PrintBooklet Failed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        void recoveraction()
        {
            LogText("Recover()", Color.Brown);
            if (Cbd1000.Recover())
                MessageBox.Show("Recover Failed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        void clearaction()
        {
            LogText("Clear()", Color.Brown);
            if (Cbd1000.Clear())
                MessageBox.Show("Clear Failed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        void flushaction()
        {
            LogText("Flush()", Color.Brown);
            if (Cbd1000.Flush())
                MessageBox.Show("Flush Failed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        // recover
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            recoveraction();
        }
        private void recoverToolStripMenuItem_Click(object sender, EventArgs e)
        {
            recoveraction();
        }
        // clear
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            clearaction();
        }
        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearaction();
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            flushaction();
        }
        private void testLoadPrintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flushaction();
        }


        private void ejectBookletToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogText("EjectBooklet()", Color.Brown);
            if (Cbd1000.EjectBooklet())
                MessageBox.Show("EjectBooklet Failed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void retractBookletToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogText("RetractBooklet()", Color.Brown);
            if (Cbd1000.RetractBooklet())
                MessageBox.Show("RetractBooklet Failed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void getStatusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CbdDriver.Status stat = Cbd1000.GetStatus();

            StringBuilder msg = new StringBuilder();
            msg.Append("Device is in " + stat.Device.ToString() + " state\n\n");
            msg.Append("Cheque feeder is " + stat.ChequeFeeder.ToString() + "\n");
            msg.Append("Cover feeder is " + stat.CoverFeeder.ToString() + "\n");
            msg.Append("Binder tape is " + stat.BinderTape.ToString() + "\n");
            msg.Append("Encoder ribbon is " + stat.EncoderRibbon.ToString() + "\n");
            msg.Append("Retract Bin is " + stat.RetractBin.ToString() + "\n");
            msg.Append("Sheet in path is " + stat.SheetInPath.ToString() + "\n");
            msg.Append("Booklet is exposed " + stat.BookletExposed.ToString() + "\n");
            msg.Append("Leaf In Path " + stat.LeafInPath.ToString() + "\n");

            MessageBox.Show( msg.ToString(), "Status:" + stat.RawStatus, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private bool OpenRbPort(string port)
        {
            try
            {
                if (!rbport.IsOpen)
                {

                    rbport.BaudRate = 38400;
                    rbport.DataBits = 8;
                    rbport.StopBits = StopBits.One;
                    rbport.Parity = Parity.None;
                    rbport.Handshake = Handshake.None;
                    rbport.NewLine = "\n";
                    rbport.PortName = port;
                    rbport.Open();
                }
                return false;
            }
            catch
            {
                return true;
            }
        }


        void OpenAction()
        {
            LogText("Open()", Color.Brown);
            if (Cbd1000.Open(toolStripComboBox1.Text))
                MessageBox.Show("Failed to open CBD port " + toolStripComboBox1.Text, "Open", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
            {
                toolStripStatusLabel1.Text = "Open on " + toolStripComboBox1.Text;
            }
        }
        void CloseAction()
        {
            LogText("Close()", Color.Brown);
            Cbd1000.Close();
            toolStripStatusLabel1.Text = "Close";
        }

        //open
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            OpenAction();
        }
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenAction();
        }

        // close
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            CloseAction();
        }
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAction();
        }



        private void loadBookletLayoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = null;
            if (Cbd1000.CurrentModel == CbdDriver.CbdModel.CBD1000)
            {
                ofd = new OpenFileDialog
                {
                    FileName = BookletLayoutFile,
                    Filter = "Layout Files (*.xml)|*.xml"
                };
            }
            else
            {
                ofd = new OpenFileDialog
                {
                    FileName = BookletLayoutFile,
                    Filter = "Layout Files (*.rbd)|*.rbd"
                };
            }

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                BookletLayoutFile = ofd.FileName;
                toolStripStatusLabel2.Text = System.IO.Path.GetFileName(BookletLayoutFile);
                if (Cbd1000.LoadBookletLayout(BookletLayoutFile))
                    MessageBox.Show("Cannot load booklet layout file:" + BookletLayoutFile, "Load", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }




        private void changeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UInt16[] max = Cbd1000.GetRetries();

            RetriesForm rf = new RetriesForm(max);
            if (rf.ShowDialog() == DialogResult.OK)
            {
                if ( Cbd1000.SetRetries(rf.m_Max))
                    MessageBox.Show("SetRetries() failed");
            }

        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            ScanForSerialPort();
        }

        private void toolStripComboBox3_Click(object sender, EventArgs e)
        {
            ScanForSerialPort();
        }

        private void toolStripComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cbd1000.CurrentModel = (CbdDriver.CbdModel)toolStripComboBox4.SelectedIndex;
        }
    }
}
