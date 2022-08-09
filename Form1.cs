using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;
using System.Media;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;
using ClosedXML.Excel;
using System.Reflection;


namespace Willkomme
{

    public partial class Form1 : Form
    {
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;
        private ZXing.BarcodeReader reader;


        public Form1()
        {
            InitializeComponent();

        }

        public void Setinf(string name, string date)
        {
            Action action = () => pictureBox.Visible = false;
            Action action2 = () => label1.Visible = false;
            Action action3 = () => label2.Visible = true;
            Action action4 = () => label2.Text = "Добро пожаловать, " + name;
            Action action5 = () => label3.Visible = true;
            Action action6 = () => label3.Text = "Дата посещения: " + date;
            Action action7 = () => label5.Visible = true;
            Action action8 = () => measure.Visible = true;

            if (InvokeRequired)
            {
                Invoke(action);
                Invoke(action2);
                Invoke(action3);
                Invoke(action4);
                Invoke(action5);
                Invoke(action6);
                Invoke(action7);
                Invoke(action8);
            }
        }
        
        public void incinf()
        {
            Action action = () => label1.Text = "QR-код не действителен";
            if (InvokeRequired)
                Invoke(action);
        }
        public void setcab(string cabine)
        {
            Action action = () => label4.Visible = true;
            Action action2 = () => label4.Text = "Ваш кабинет: " + cabine;
            Action action3 = () => pictureBox1.Visible = true;
            Action piccab = () => pictureBox1.Image = Properties.Resources._101;
            Action piccab2 = () => pictureBox1.Image = Properties.Resources._102;
            Action piccab3 = () => pictureBox1.Image = Properties.Resources._103;
            Action piccab4 = () => pictureBox1.Image = Properties.Resources._104;
            Action piccab5 = () => pictureBox1.Image = Properties.Resources._105;
            Action piccab6 = () => pictureBox1.Image = Properties.Resources._106;
            if (InvokeRequired)
            {
                Invoke(action);
                Invoke(action2);
                Invoke(action3);
                switch (cabine)
                {
                    case "101":
                        Invoke(piccab);
                        break;
                    case "102":
                        Invoke(piccab2);
                        break;
                    case "103":
                        Invoke(piccab3);
                        break;
                    case "104":
                        Invoke(piccab4);
                        break;
                    case "105":
                        Invoke(piccab5);
                        break;
                    case "106":
                        Invoke(piccab6);
                        break;
                }

            } 
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            comboBox1.Text = "";
            comboBox1.Items.Clear();
            if(ports.Length != 0)
            {
                comboBox1.Items.AddRange(ports);
                comboBox1.SelectedIndex = 0;
            }
            this.TopMost = false;
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

            reader = new ZXing.BarcodeReader();
            reader.Options.PossibleFormats = new List<ZXing.BarcodeFormat>();
            reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.QR_CODE);
            
            if (videoDevices.Count > 0)
            {
                foreach (FilterInfo device in videoDevices)
                {
                    lbCams.Items.Add(device.Name);
                }
                lbCams.SelectedIndex = 0;

            }
            
            
        }
        
        private void btnStart_Click(object sender, EventArgs e)
        {
            btnStart.Visible = false;
            lbCams.Visible = false;
            label1.Text = "Добро пожаловать!";
            comboBox1.Visible = false;
            
            videoSource = new VideoCaptureDevice(videoDevices[lbCams.SelectedIndex].MonikerString);
            videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
            videoSource.Start(); 
        }

        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap bitmap = (Bitmap)eventArgs.Frame.Clone();
            pictureBox.Image = bitmap;

            ZXing.Result result = reader.Decode((Bitmap)eventArgs.Frame.Clone());
            if (result != null)
            {
                string filePath = @"listr.xlsx";

                string textToFind = result.Text;

                using (var workbook = new XLWorkbook(filePath))
                {
                    string cabine;
                    var worksheet = workbook.Worksheets.First();
                    var column = worksheet.Column("I");
                    var columnCells = column.CellsUsed();
                    var cell = columnCells.First(i => i.Value.ToString().Contains(textToFind));
                    if (cell != null)
                    {
                        var row = cell.WorksheetRow();

                        string name = Convert.ToString(row.Cell("A").Value);
                        string date = Convert.ToString(row.Cell("D").Value);
                        string type = Convert.ToString(row.Cell("E").Value);
                        Setinf(name,date);
                        switch (type)
                        {
                            case "Хирург":
                                cabine = "101";
                                setcab(cabine);
                                break;
                            case "Педиатр":
                                cabine = "102";
                                setcab(cabine);
                                break;
                            case "Нефролог":
                                cabine = "103";
                                setcab(cabine);
                                break;
                            case "Онколог":
                                cabine = "104";
                                setcab(cabine);
                                break;
                            case "Терапевт":
                                cabine = "105";
                                setcab(cabine);
                                break;
                            case "Нарколог":
                                cabine = "106";
                                setcab(cabine);
                                break;
                        }
                    }
                    else
                    {
                        incinf();
                    }

                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (videoSource != null)
            {
                videoSource.SignalToStop();
                videoSource.WaitForStop();      
            }
        }

        private void measure_Click(object sender, EventArgs e)
        {
            if (videoSource != null)
            {
                videoSource.SignalToStop();
                videoSource.WaitForStop();
            }
            label5.ForeColor = Color.Black;
            measure.Visible = false;
            try
            {
                MySerialPort.PortName = comboBox1.Text;
                MySerialPort.Open();

            }
            catch
            {
                MessageBox.Show("Ошибка подключения");
            }

        }

        private void MySerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            label5.Text = MySerialPort.ReadLine();
            string value = MySerialPort.ReadLine();
            if (value == "0 ")
            { 

                MySerialPort.Close();
                label5.Text = "У вас высокая темпертатура, пожалуйста оденьте маску.";
                Thread.Sleep(10000);

                label2.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
                pictureBox1.Visible = false;
                label1.Text = "Добро пожаловать!";
                

                videoSource = new VideoCaptureDevice(videoDevices[lbCams.SelectedIndex].MonikerString);
                videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
                videoSource.Start();

            }
            if (value == "1") 
            {
                MySerialPort.Close();
                label5.Text = "Все в норме";
                Thread.Sleep(10000);

                label2.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
                pictureBox1.Visible = false;
                label1.Text = "Добро пожаловать!";
                

                videoSource = new VideoCaptureDevice(videoDevices[lbCams.SelectedIndex].MonikerString);
                videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
                videoSource.Start();
            }

        }
    }
}
