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
        float Current1;
        float Current2;
        float Current3;
        float Freq;
        float ActiveP; 
        float ReActiveP;
        float PFactor;
        float Energy;
        float CurrentUnbalance;
        int i = 2;
        public Form1()
        {
            InitializeComponent();


        }

        public void Form1_Load(object sender, EventArgs e)
        {
            
            comboBox1.DataSource = SerialPort.GetPortNames(); //tạo list cổng COM cho combobox

            /* Đoạn code này để khởi tạo đồ thị dùng thư viện ZedGraph*/
            GraphPane graphPane = zedGraphControl1.GraphPane;
            graphPane.Title.Text = "Voltage";
            graphPane.XAxis.Title.Text = "Thời gian";
            graphPane.YAxis.Title.Text = "Voltage";
            graphPane.XAxis.Type = AxisType.Date;
            graphPane.XAxis.Scale.Format = "hh:mm:ss";       
            graphPane.XAxis.Scale.MinorStep = 5;
            graphPane.XAxis.Scale.MinorUnit = DateUnit.Second;
            graphPane.XAxis.Scale.MajorUnit = DateUnit.Second;
            graphPane.XAxis.Scale.MajorStep = 30;
            graphPane.YAxis.Scale.Min = 0;
            graphPane.YAxis.Scale.Max = 260;
            point1 = new PointPairList();
            point2 = new PointPairList();
            point3 = new PointPairList();

            LineItem myCurve1 = graphPane.AddCurve("L1Voltage", point1, System.Drawing.Color.Red);

            LineItem myCurve2 = graphPane.AddCurve("L2Voltage", point2, System.Drawing.Color.Blue);

            LineItem myCurve3 = graphPane.AddCurve("L3Voltage", point3, System.Drawing.Color.Yellow);
            

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
                textBox4.Text = "Connected";
            }
           else { textBox4.Text = "Error"; }
        }

        private void ReadValue() // Đọc giá trị dòng và điện áp theo địa chỉ thanh ghi
        {
            try
            {
                while (modbusClient.Connected)
                {
                    // Đọc giá trị điện áp theo địa chỉ thanh ghi theo truyền thông Modbus và thư viện easyModbus
                    Voltage1 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3027, 2), ModbusClient.RegisterOrder.HighLow);
                    Voltage2 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3029, 2), ModbusClient.RegisterOrder.HighLow);
                    Voltage3 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3031, 2), ModbusClient.RegisterOrder.HighLow);
                    Current1 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(2999, 2), ModbusClient.RegisterOrder.HighLow);
                    Current2 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3001, 2), ModbusClient.RegisterOrder.HighLow);
                    Current3 = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3003, 2), ModbusClient.RegisterOrder.HighLow);
                    Freq = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3009, 2), ModbusClient.RegisterOrder.HighLow);
                    ActiveP = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3059, 2), ModbusClient.RegisterOrder.HighLow);
                    ReActiveP = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3067, 2), ModbusClient.RegisterOrder.HighLow);
                    PFactor = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3083, 2), ModbusClient.RegisterOrder.HighLow);
                    //Energy = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3067, 2), ModbusClient.RegisterOrder.HighLow);
                    CurrentUnbalance = ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(3017, 2), ModbusClient.RegisterOrder.HighLow);

                    // Hiển thị giá trị điện áp lên textbox theo thời gian thực bằng backgroundworker và beginInvoke
                    textBox1.BeginInvoke(new Action(() => { textBox1.Text = Voltage1.ToString(); }));
                    textBox2.BeginInvoke(new Action(() => { textBox2.Text = Voltage2.ToString(); }));
                    textBox3.BeginInvoke(new Action(() => { textBox3.Text = Voltage3.ToString(); }));
                    textBox5.BeginInvoke(new Action(() => { textBox5.Text = Current1.ToString(); }));
                    textBox6.BeginInvoke(new Action(() => { textBox6.Text = Current2.ToString(); }));
                    textBox7.BeginInvoke(new Action(() => { textBox7.Text = Current3.ToString(); }));
                    textBox8.BeginInvoke(new Action(() => { textBox8.Text = Freq.ToString(); }));
                    textBox9.BeginInvoke(new Action(() => { textBox9.Text = ActiveP.ToString(); }));
                    textBox10.BeginInvoke(new Action(() => { textBox10.Text = ReActiveP.ToString(); }));
                    textBox11.BeginInvoke(new Action(() => { textBox11.Text = PFactor.ToString(); }));
                    textBox12.BeginInvoke(new Action(() => { textBox12.Text = Energy.ToString(); }));
                    textBox13.BeginInvoke(new Action(() => { textBox13.Text = CurrentUnbalance.ToString(); }));
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Vui long kiểm tra kết nối với thiết bị", "lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }


        // cho phép backgroundWorker hoạt động
        private void button3_Click(object sender, EventArgs e) 
        {
            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync();
            }

        }


        //backgroundWorker thực hiện hàm Readvalue
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ReadValue();
        } 


        // Khởi động timer1
        private void button2_Click(object sender, EventArgs e)
        {
      
            timer1.Start(); 

        }


        // Hàm khởi tạo file exel
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


        // chạy hàm Exel
        private void button4_Click(object sender, EventArgs e)
        {
            
            Exel();
            //timer1.Start();
          //  SavetoExel();  

        }


        //hàm lưu giá trị đo vào file exel  
        private void SavetoExel()
        {

            ws.Cells[i, 1] = DateTime.Now.ToString();
            ws.Cells[i, 2] = Voltage1.ToString();
            ws.Cells[i, 3] = Voltage2.ToString();
            ws.Cells[i, 4] = Voltage3.ToString();

            i++;
        }

        // Update đồ thị và exel theo Timer 1
        private void timer1_Tick(object sender, EventArgs e) 
        {

            point1.Add(new XDate(DateTime.Now), Voltage1);
            point2.Add(new XDate(DateTime.Now), Voltage2);
            point3.Add(new XDate(DateTime.Now), Voltage3);
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
            SavetoExel();
        }

        // Hiển thị file Exel
        
        private void button5_Click(object sender, EventArgs e)
        {

            xla.Visible = !xla.Visible;

        }

       
    }
}
