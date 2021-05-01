using EasyModbus;
using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO.Ports;
using ZedGraph;


namespace PM2220
{

    public partial class Form1 : Form       // Khai báo class Form1 thuộc Class Form
    {
        ModbusClient modbusClient; // Khai báo trường (Field) modbusClient thuộc Class ModbusClient
        //int[] Value = { 0};
        PointPairList point1;
        PointPairList point2;
        PointPairList point3;

        Microsoft.Office.Interop.Excel.Application xla;

        Microsoft.Office.Interop.Excel._Workbook wb;
        Microsoft.Office.Interop.Excel.Worksheet ws;
        Microsoft.Office.Interop.Excel.Range range;
        float Voltage1;
        float Voltage2;
        float Voltage3;
        int i = 2;
        public Form1()
        {
            InitializeComponent();


        }

        public void Form1_Load(object sender, EventArgs e)
        {
            //var listPort = new[] { "COM1", "COM2", "COM3", "COM4", "COM5" }; 
            comboBox1.DataSource = SerialPort.GetPortNames(); //tạo list cổng COM cho combobox

            GraphPane graphPane = zedGraphControl1.GraphPane;
            graphPane.Title.Text = "Voltage";
            graphPane.XAxis.Title.Text = "Thời gian";
            graphPane.YAxis.Title.Text = "Voltage";

            graphPane.XAxis.Type = AxisType.Date;
            graphPane.XAxis.Scale.Format = "hh:mm:ss";
            //graphPane.XAxis.Scale.Max = 30;
            graphPane.XAxis.Scale.MinorStep = 1;
            graphPane.XAxis.Scale.MinorUnit = DateUnit.Second;

            graphPane.XAxis.Scale.MajorUnit = DateUnit.Second;
            graphPane.XAxis.Scale.MajorStep = 5;
            graphPane.YAxis.Scale.Min = 0;
            graphPane.YAxis.Scale.Max = 260;
            point1 = new PointPairList();
            point2 = new PointPairList();
            point3 = new PointPairList();

            LineItem myCurve1 = graphPane.AddCurve("L1Voltage", point1, System.Drawing.Color.Red);

            LineItem myCurve2 = graphPane.AddCurve("L2Voltage", point2, System.Drawing.Color.Blue);

            LineItem myCurve3 = graphPane.AddCurve("L3Voltage", point3, System.Drawing.Color.Yellow);
            //points.Add(new XDate(DateTime.Now), 100);
            //points.Add(new XDate(DateTime.Now), 200);
            // points.Add(new XDate(DateTime.Now), 300);
            // graphPane.AxisChange();

        }




        private void button1_Click(object sender, EventArgs e) // Method kết nối với thiết bị  khi click vào
        {
            modbusClient = new ModbusClient(); // Khai bao ket noi voi modbus client tai cong COM4
            //Khai bao cac thuoc tinh cua cong COM
            modbusClient.SerialPort = comboBox1.SelectedItem.ToString();
            modbusClient.Baudrate = 19200;
            modbusClient.Parity = Parity.None;
            modbusClient.StopBits = StopBits.One;

            // Mở kết nối với ModbusClient
            try { modbusClient.Connect(); }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng kiểm tra lại cổng COM và cài đặt thiết bị: \n \n SlaveID = 1 \n Baudrate = 19200 \n Parity = None \n Stopbits = One",
                    "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (modbusClient.Connected)
            {
                MessageBox.Show("Kết nối thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //textBox4.Text = ("Connected");
            }
            textBox4.Text = modbusClient.Connected.ToString();
        }

        private void ReadValue() // Đọc giá trị dòng và điện áp theo địa chỉ thanh ghi
        {
            try
            {
                while (modbusClient.Connected)
                {
                    Voltage1 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3027, 2), ModbusClient.RegisterOrder.HighLow);
                    Voltage2 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3029, 2), ModbusClient.RegisterOrder.HighLow);
                    Voltage3 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3031, 2), ModbusClient.RegisterOrder.HighLow);
                    //Value = modbusClient.ReadHoldingRegisters(0, 15);
                    textBox1.BeginInvoke(new Action(() => { textBox1.Text = Voltage1.ToString(); }));
                    textBox2.BeginInvoke(new Action(() => { textBox2.Text = Voltage2.ToString(); }));
                    textBox3.BeginInvoke(new Action(() => { textBox3.Text = Voltage3.ToString(); }));
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Vui long kiểm tra kết nối với thiết bị", "lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync();
            }


            /*
           while (modbusClient.Connected)
            {
                try 
                {
                    ReadValue();
                    //System.Threading.Thread.Sleep(2000);
                }
                catch (ArgumentException ex)
                {
                    MessageBox.Show("Địa chỉ thanh ghi không chính xác. \n" + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Vui lòng kiểm tra lại kết nối với thiết bị", "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
           */
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ReadValue();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*
            while (true)
            {
                float Voltage1 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3019, 2), ModbusClient.RegisterOrder.HighLow);
                float Voltage2 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3021, 2), ModbusClient.RegisterOrder.HighLow);
                float Voltage3 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3023, 2), ModbusClient.RegisterOrder.HighLow);
            
                
            }
            */
            //Value = modbusClient.ReadHoldingRegisters(0, 15);
            //System.Threading.Thread.Sleep(2000);
            // textBox1.Text = Value[0].ToString();

            //textBox2.Text = Value[0].ToString();
            //textBox3.Text = Value[3].ToString();
            timer1.Start(); // set timer1.Enabled = true
                            // timer1.Enabled = true;


        }

        private void Exel()
        {
            xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            range = (Microsoft.Office.Interop.Excel.Range)ws.get_Range("A1", "D1");
            ws.Cells[1, 1] = "Thoi gian";
            ws.Cells[1, 2] = " L1 Voltage";
            ws.Cells[1, 3] = " L2 Voltage";
            ws.Cells[1, 4] = " L3 Voltage";
            range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range.EntireColumn.AutoFit();
            range.Font.Bold = true;
            ws.Columns.AutoFit();
            wb.SaveAs("C:\\Users\\os\\Desktop\\New folder (3)\\test.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Exel();
        }
        private void SavetoExel()
        {

            ws.Cells[i, 1] = DateTime.Now.ToString();
            ws.Cells[i, 2] = Voltage1.ToString();
            ws.Cells[i, 3] = Voltage2.ToString();
            ws.Cells[i, 4] = Voltage3.ToString();

            i++;
        }
        private void timer1_Tick(object sender, EventArgs e) // vẽ đồ thị theo khoảng thời gian Timer 1
        {

            point1.Add(new XDate(DateTime.Now), Voltage1);
            point2.Add(new XDate(DateTime.Now), Voltage2);
            point3.Add(new XDate(DateTime.Now), Voltage3);
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
            SavetoExel();
        }

        private void button5_Click(object sender, EventArgs e)
        {

            xla.Visible = !xla.Visible;

        }

        private void zedGraphControl1_Load(object sender, EventArgs e)
        {

        }
    }
}
