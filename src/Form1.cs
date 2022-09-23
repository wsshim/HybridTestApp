using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;  //시리얼통신을 위해 추가해줘야 함
using System.Collections;
using System.Timers;
using System.Diagnostics;

namespace Serial_Communication
{
    public partial class Form1 : Form
    {
        const string app_ver = "v1.0.0";
        const double alarmCurrentValue = 4.0;
        const double DetectDryContactAlarm = 200 * 0.001221; // 0.001221 = 5 / 4095;
        const double DetectDryContactNormal = 3800 * 0.001221; // 0.001221 = 5 / 4095;
        const int comm_err_max = 10;
        const string tkb_port = "COM1";
        const string dut_port = "COM2";
        byte[] RCV1_data = new byte[100];
        byte[] RCV2_data = new byte[100];

        byte[] tkb_set_cmd = {0xA7,0,0,0,0,0x0F,0};

        int tkb_ready = 0;
        int dut_ready = 0;

        int rcv1_len = 0;
        int rcv2_len = 0;

        int comm1_err_cnt = 0;
        int comm2_err_cnt = 0;

        int data2_cnt = 1;

        float data_volt_current_peak = 0;
        float data_volt_current = 0;
        float data_volt_ex = 0;
        float data_volt_n2 = 0;
        float data_volt_tempP = 0;
        float data_volt_tempC = 0;
        float data_dry1 = 0;
        float data_dry2 = 0;

        int if_alarm_LOOP = 0;

        bool if_alarm_ex = false;
        bool if_alarm_n2 = false;
        bool if_alarm_tempP = false;
        bool if_alarm_tempC = false;
        bool if_status_Sol_N2 = false;
        bool if_status_Sol_CDA = false;
        bool if_status_LM = false;
        bool if_status_LOOP = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)  //폼이 로드되면
        {
            // comboBox_port_1.DataSource = SerialPort.GetPortNames(); //연결 가능한 시리얼포트 이름을 콤보박스에 가져오기 
            // comboBox_port_2.DataSource = SerialPort.GetPortNames(); //연결 가능한 시리얼포트 이름을 콤보박스에 가져오기 
            System_Ready_Dis1(0);
            System_Ready_Dis2(0);
            timer_init();
            Rcv1_buf_clear();
            Rcv2_buf_clear();
            label_app_ver.Text = app_ver;
        }

        private void timer_init()
        {
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 1000;
            timer.Elapsed += new ElapsedEventHandler(timer_handler);
            timer.Start();
        }
        private void timer_handler(object sender, ElapsedEventArgs e)
        {
            if (comm1_err_cnt++ > comm_err_max)
            {
                comm1_err_cnt = comm_err_max;
                System_Ready_Dis1(0);
                Rcv1_buf_clear();
            }
            if (comm2_err_cnt++ > comm_err_max)
            {
                comm2_err_cnt = comm_err_max;
                System_Ready_Dis2(0);
                Rcv2_buf_clear();
            }
        }


        private void Button_connect1_Click(object sender, EventArgs e)  //통신 연결하기 버튼
        {
            if (!serialPort1.IsOpen)  //시리얼포트가 열려 있지 않으면
            {
                
                serialPort1.PortName = tkb_port; // comboBox_port_1.Text;  //콤보박스의 선택된 COM포트명을 시리얼포트명으로 지정
                serialPort1.BaudRate = 19200;//Convert.ToInt32(comboBox_speed.Text);  //보레이트 변경이 필요하면 숫자 변경하기
                serialPort1.StopBits = StopBits.One;
                serialPort1.Parity = Parity.Odd;
                serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);

                serialPort1.Open();  //시리얼포트 열기

                button_connect1.Enabled = false;
            }
        }

        private void Button_connect2_Click(object sender, EventArgs e)
        {
            if (!serialPort2.IsOpen)  //시리얼포트가 열려 있지 않으면
            {
                serialPort2.PortName = dut_port; // comboBox_port_2.Text;  //콤보박스의 선택된 COM포트명을 시리얼포트명으로 지정
                serialPort2.BaudRate = 19200;//Convert.ToInt32(comboBox_speed.Text);  //보레이트 변경이 필요하면 숫자 변경하기
                serialPort2.StopBits = StopBits.One;
                serialPort2.Parity = Parity.Odd;
                serialPort2.DataReceived += new SerialDataReceivedEventHandler(serialPort2_DataReceived);

                serialPort2.Open();  //시리얼포트 열기

                button_connect2.Enabled = false;
            }
            
        }
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)  //수신 이벤트가 발생하면 이 부분이 실행된다.
        {
            this.Invoke(new EventHandler(MySerial1Received));  //메인 쓰레드와 수신 쓰레드의 충돌 방지를 위해 Invoke 사용. MySerial1Received로 이동하여 추가 작업 실행.
        }
        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)  //수신 이벤트가 발생하면 이 부분이 실행된다.
        {            
            this.Invoke(new EventHandler(MySerial2Received));  //메인 쓰레드와 수신 쓰레드의 충돌 방지를 위해 Invoke 사용. MySerial2Received로 이동하여 추가 작업 실행.
        }
        private void MySerial1Received(object s, EventArgs e)  //여기에서 수신 데이타를 사용자의 용도에 따라 처리한다.
        {
            int ReceiveData;
            int rx_cnt = serialPort1.BytesToRead;
            while (rx_cnt > 0)
            {
                rx_cnt--;
                ReceiveData = serialPort1.ReadByte();  //시리얼 버터에 수신된 데이타를 ReceiveData 읽어오기
                //richTextBox_received.Text = richTextBox_received.Text + string.Format("{0:X2}", ReceiveData);  //int 형식을 string형식으로 변환하여 출력
                RCV1_data[rcv1_len++] = (byte)ReceiveData;

                if (rcv1_len >= 20) Rcv1_buf_clear();
                if (RCV1_data[0] != 0xA6) Rcv1_buf_clear();
                else
                {
                    if (rcv1_len >= 12)
                    {
                        byte cs = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            cs ^= RCV1_data[i];
                        }

                        if ((cs == RCV1_data[11]) && (RCV1_data[10] == 0x0E))
                        {
                            comm1_err_cnt = 0;
                            System_Ready_Dis1(1);
                            
                            data_volt_current = (float)RCV1_data[1] / 10;
                            data_volt_ex = (float)RCV1_data[2] / 10;
                            data_volt_n2 = (float)RCV1_data[3] / 10;
                            data_volt_tempP = (float)RCV1_data[4] / 10;
                            data_volt_tempC = (float)RCV1_data[5] / 10;
                            data_dry1 = ValidateDryContactData(RCV1_data[6]);
                            data_dry2 = ValidateDryContactData(RCV1_data[7]);

                            if_alarm_LOOP = RCV1_data[9];

                            if_alarm_ex = Interface_alarm_parser(RCV1_data[8], 0);
                            if_alarm_n2 = Interface_alarm_parser(RCV1_data[8], 1);
                            if_alarm_tempP = Interface_alarm_parser(RCV1_data[8], 2);
                            if_alarm_tempC = Interface_alarm_parser(RCV1_data[8], 3);
                            if_status_Sol_N2 = Interface_alarm_parser(RCV1_data[8], 4);
                            if_status_Sol_CDA = Interface_alarm_parser(RCV1_data[8], 5);
                            if_status_LM = Interface_alarm_parser(RCV1_data[8], 6);
                            if_status_LOOP = Interface_alarm_parser(RCV1_data[8], 7);

                            if (data_volt_current > data_volt_current_peak) data_volt_current_peak = data_volt_current;


                            if (if_status_Sol_N2 == false)
                                label_status_Sol_N2.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
                            else
                                label_status_Sol_N2.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;

                            if (if_status_Sol_CDA == false)
                                label_status_Sol_CDA.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
                            else
                                label_status_Sol_CDA.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;

                            if (if_status_LM == false)
                                label_status_LM.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
                            else
                                label_status_LM.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;

                            if (if_status_LOOP == false)
                                label_status_LOOP.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
                            else
                                label_status_LOOP.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;

                            if (if_status_LM == true)
                            {
                                button_LM_Open.Enabled = false;
                                button_LOOP_Open.Enabled = false;
                                button_LOOP_Close.Enabled = false;
                            }
                            else if (if_status_LOOP == true)
                            {

                                button_LOOP_Open.Enabled = false;
                                button_LM_Open.Enabled = false;
                                button_LM_Close.Enabled = false;
                            }
                            else
                            {
                                button_LM_Open.Enabled = true;
                                button_LM_Close.Enabled = true;
                                button_LOOP_Open.Enabled = true;
                                button_LOOP_Close.Enabled = true;
                            }

                            if ((tkb_ready == 1) && (dut_ready == 1))
                            {
                                label_val_current.Text = string.Format("{0}", Math.Round(data_volt_current * 10) / 10); // current
                                label_val_peak.Text = string.Format("{0}", Math.Round(data_volt_current_peak * 10) / 10); // peak

                                label_ex_volt_interface.Text = string.Format("{0}", data_volt_ex); // exhaust
                                label_n2_volt_interface.Text = string.Format("{0}", data_volt_n2); // n2
                                label_tempP_volt_interface.Text = string.Format("{0}", data_volt_tempP); // PAD
                                label_tempC_volt_interface.Text = string.Format("{0}", data_volt_tempC); // CAT

                                label_ex_val_interface.Text = string.Format("{0}", Math.Round((data_volt_ex / 5) * 10) / 10);
                                label_n2_val_interface.Text = string.Format("{0}", Math.Round((data_volt_n2 * 60) * 10) / 10);
                                label_tempP_val_interface.Text = string.Format("{0}", Math.Round((data_volt_tempP * 10) * 10) / 10); // PAD
                                label_tempC_val_interface.Text = string.Format("{0}", Math.Round((data_volt_tempC * 10) * 10) / 10); // CAT


                                if (if_alarm_ex == false)
                                    label_exh_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_exh_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (if_alarm_n2 == false)
                                    label_n2_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_n2_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (if_alarm_tempP == false)
                                    label_pad_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_pad_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (if_alarm_tempC == false)
                                    label_cat_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_cat_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (if_alarm_LOOP == 0)
                                    label_loop_status_if.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_loop_status_if.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (data_dry1 == 0)
                                    label_dry1_status_if.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_dry1_status_if.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (data_dry2 == 0)
                                    label_dry2_status_if.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_dry2_status_if.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                            }
                        }

//                      StringBuilder hex = new StringBuilder(2);
//                      hex.AppendFormat("{0:x2}", RCV1_data[8]);
//                      label_debug1.Text = hex.ToString();

                        Rcv1_buf_clear();
                        return;
                    }
                }
            }
        }
        private void MySerial2Received(object s, EventArgs e)  //여기에서 수신 데이타를 사용자의 용도에 따라 처리한다.
        {
            int ReceiveData;
            int rx_cnt = serialPort2.BytesToRead;

            while (rx_cnt > 0)
            {
                rx_cnt--;
                ReceiveData = serialPort2.ReadByte();  //시리얼 버터에 수신된 데이타를 ReceiveData 읽어오기
                RCV2_data[rcv2_len++] = (byte)ReceiveData;
                if (rcv2_len >= 50) Rcv2_buf_clear();
                if (RCV2_data[0] != 0xA5) Rcv2_buf_clear();
                else
                {
                    if (rcv2_len >= 35)
                    {
                        byte cs = 0;
                        for (int i = 1; i < 34; i++)
                        {
                            cs ^= RCV2_data[i];
                        }

                        if ((cs == RCV2_data[34]) && (RCV2_data[33] == 0x0D))
                        {
                            comm2_err_cnt = 0;
                            System_Ready_Dis2(1);

                            if ((tkb_ready == 1) && (dut_ready == 1))
                            {
                                data2_cnt = 1;
                                label_ex_val_rs232.Text = string.Format("{0}.{1}{2}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // exhaust
                                label_n2_val_rs232.Text = string.Format("{0}{1}{2}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // n2
                                label_tempP_val_rs232.Text = string.Format("{0}{1}{2}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // pad
                                label_tempC_val_rs232.Text = string.Format("{0}{1}{2}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // cat
                                label_hum_val_rs232.Text = string.Format("{0}{1}{2}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // hum

                                if (RCV2_data[data2_cnt++] == 0) label_tiltX_val_rs232.Text = " ";
                                else label_tiltX_val_rs232.Text = "-";
                                label_tiltX_val_rs232.Text += string.Format("{0}{1}.{2}{3}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // tilt x

                                if (RCV2_data[data2_cnt++] == 0) label_tiltY_val_rs232.Text = " ";
                                else label_tiltY_val_rs232.Text = "-";
                                label_tiltY_val_rs232.Text += string.Format("{0}{1}.{2}{3}", RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++], RCV2_data[data2_cnt++]); // tilt y

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_exh_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_exh_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_n2_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_n2_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_pad_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_pad_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_cat_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_cat_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_cover_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_cover_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_hum_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_hum_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;

                                if (RCV2_data[data2_cnt++] == 0)
                                    label_tilt_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.Comm16x16;
                                else
                                    label_tilt_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.homeAlarm16x16;
                            }
                        }

                        Rcv2_buf_clear();
                        return;
                    }
                }
            }
        }
        private void Rcv1_buf_clear()
        {
            int i;
            for(i = 0; i < (rcv1_len); i++)
            {
                RCV1_data[i] = 0;
            }
            rcv1_len = 0;
        }
        private void Rcv2_buf_clear()
        {
            int i;
            for (i = 0; i < (rcv2_len); i++)
            {
                RCV2_data[i] = 0;
            }
            rcv2_len = 0;
        }

        private void Display_clear()
        {
            label_val_current.Text = string.Format("{0}", 0); // peak
            label_val_peak.Text = string.Format("{0}", 0); // peak

            label_ex_volt_interface.Text = string.Format("{0}", 0); // exhaust
            label_n2_volt_interface.Text = string.Format("{0}", 0); // n2
            label_tempP_volt_interface.Text = string.Format("{0}", 0); // PAD
            label_tempC_volt_interface.Text = string.Format("{0}", 0); // CAT

            label_ex_val_interface.Text = string.Format("{0}", 0);
            label_n2_val_interface.Text = string.Format("{0}", 0);
            label_tempP_val_interface.Text = string.Format("{0}", 0); // PAD
            label_tempC_val_interface.Text = string.Format("{0}", 0); // CAT

            label_ex_val_rs232.Text = string.Format("{0}.{1}{2}", 0, 0, 0); // exhaust
            label_n2_val_rs232.Text = string.Format("{0}{1}{2}", 0, 0, 0); // n2
            label_tempP_val_rs232.Text = string.Format("{0}{1}{2}", 0, 0, 0); // pad
            label_tempC_val_rs232.Text = string.Format("{0}{1}{2}", 0, 0, 0); // cat
            label_hum_val_rs232.Text = string.Format("{0}{1}{2}", 0, 0, 0); // hum

            label_tiltX_val_rs232.Text = " ";
            label_tiltX_val_rs232.Text += string.Format("{0}{1}.{2}{3}", 0, 0, 0, 0); // tilt x

            label_tiltY_val_rs232.Text = " ";
            label_tiltY_val_rs232.Text += string.Format("{0}{1}.{2}{3}", 0, 0, 0, 0); // tilt y

            label_exh_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_n2_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_pad_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_cat_status_interface.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;

            label_exh_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_n2_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_pad_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_cat_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            
            label_hum_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_tilt_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_cover_status_rs232.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;

            label_loop_status_if.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_dry1_status_if.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
            label_dry2_status_if.Image = HYBRID_TEST_APP.Properties.Resources.ImageCloseButton;
        }
        private void Button_send_Click(object sender, EventArgs e)  //보내기 버튼을 클릭하면
        {
            ArrayList ary = new ArrayList();
            ary.Add(checkBox1);            ary.Add(checkBox2);            ary.Add(checkBox3);            ary.Add(checkBox4);            ary.Add(checkBox5);
            ary.Add(checkBox6);            ary.Add(checkBox7);            ary.Add(checkBox8);            ary.Add(checkBox9);            ary.Add(checkBox10);
            ary.Add(checkBox11);           ary.Add(checkBox12);           ary.Add(checkBox13);           ary.Add(checkBox14);           ary.Add(checkBox15);
            ary.Add(checkBox16);           ary.Add(checkBox17);           ary.Add(checkBox18);           ary.Add(checkBox19);           ary.Add(checkBox20);
            ary.Add(checkBox21);           ary.Add(checkBox22);           ary.Add(checkBox23);           ary.Add(checkBox24);           ary.Add(checkBox25);
            ary.Add(checkBox26);           ary.Add(checkBox27);           ary.Add(checkBox28);           ary.Add(checkBox29);           ary.Add(checkBox30);
            ary.Add(checkBox31);           ary.Add(checkBox32);           ary.Add(checkBox33);           ary.Add(checkBox34);           ary.Add(checkBox35);
            ary.Add(checkBox36);           ary.Add(checkBox37);           ary.Add(checkBox38);           ary.Add(checkBox39);           ary.Add(checkBox40);
            ary.Add(checkBox41);           ary.Add(checkBox42);           ary.Add(checkBox43);           ary.Add(checkBox44);           ary.Add(checkBox45);
            ary.Add(checkBox46);           ary.Add(checkBox47);           ary.Add(checkBox48);           ary.Add(checkBox49);           ary.Add(checkBox50);

           
            
            ArrayList arrayList = new ArrayList();
            string bufer = "";
            for (int q = 0; q < ary.Count; q++)
            {
                if (((CheckBox)ary[q]).CheckState == CheckState.Checked)
                {

                    bufer = (    "1" + bufer);
                }
                else
                {
                    bufer = (  "0" + bufer);
                }
                if((q != 0 && q % 8 == 7) || (ary.Count-1 ==q))
                {
                    if ((ary.Count - 1 == q))
                    {
                        bufer = ("000000" + bufer);
                    }
                    arrayList.Add(bufer);
                    bufer = "";
                }   
            }
            string result = "a5";
            for(int q = 0; q < arrayList.Count; q++)
            {
                result += (ByteArrayToString(GetBytes((String)arrayList[q])));
            }
            result += "0d00";

            byte[] bytesToSend = new byte[10];
           /* for(int q = 0; q < bytesToSend.Length; q++)
            {
                Console.WriteLine("b:" + result.Substring(2 * q, 2));
                
                bytesToSend[q] = getHex(result.Substring(2 * q, 2));
            }*/

            textBox_send.Text = "";
            //textBox_send.Text = result;
            //Console.WriteLine("aa:" + ByteToString(bytesToSend));
            //textBox_send.Text = ByteArrayToString(getHex(result));
            bytesToSend = getHex_CS(result);
            textBox_send.Text = ByteArrayToString(bytesToSend);
            //serialPort2.Write(getHex(result), 0 , getHex(result).Length);
            if (serialPort2.IsOpen)
                serialPort2.Write(bytesToSend, 0, bytesToSend.Length);
        }
        private string ByteToString(byte[] strByte) { string str = Encoding.Default.GetString(strByte); return str; }
        private byte[] StringToByte(string str) { byte[] StrByte = Encoding.UTF8.GetBytes(str); return StrByte; }

        public byte[] getHex_CS(string hexString)
        {

            //string hexString = "A1B2C3";
            byte[] xbytes = new byte[hexString.Length / 2];
            for (int i = 0; i < xbytes.Length; i++)
            {
                xbytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            }
            for (int i = 1; i < xbytes.Length - 1; i++)
            {
                xbytes[xbytes.Length - 1] ^= xbytes[i];
            }
            return xbytes;//Convert.ToByte(srcValue, 16);

        }

        public byte[] getHex(string hexString)
        {

            //string hexString = "A1B2C3";
            byte[] xbytes = new byte[hexString.Length / 2];
            for (int i = 0; i < xbytes.Length; i++)
            {
                xbytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            }
            return xbytes;//Convert.ToByte(srcValue, 16);

        }
        // String을 바이트 배열로 변환
        public static byte[] GetBytes(string bitString)
        {
            return Enumerable.Range(0, bitString.Length / 8).
                Select(pos => Convert.ToByte(
                    bitString.Substring(pos * 8, 8),
                    2)
                ).ToArray();
        }
        public static string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }
        private int ValidateDryContactData(int data)
        {
            int alarmValue = 0;

            double calculateValue;
            calculateValue = (double)data / 10; // 펌웨어에서 10을 곱해서 받기때문에 10을 나눠준다

            if (calculateValue < DetectDryContactAlarm)
                alarmValue = 1;
            else if (calculateValue > DetectDryContactNormal)
                alarmValue = 0;
            else
                alarmValue = 0;

            return alarmValue;
        }
        private bool Interface_alarm_parser(byte alarm, byte index)
        {
            bool calculateValue = false;
            byte _buffer = 0;

            _buffer = (byte)(alarm >> index);
            _buffer = (byte)(_buffer & 0x01);

            if (_buffer == 0)
                calculateValue = false;
            else
                calculateValue = true;

            return calculateValue;
        }

        private void System_Ready_Dis1(int stat)
        {
            if (stat == 0)
            {
                label_con_tkb.Text = "NG";
                label_con_tkb.BackColor = Color.Red;
                tkb_ready = 0;
                Display_clear();
            }
            else
            {
                label_con_tkb.Text = "OK";
                label_con_tkb.BackColor = Color.Green;
                tkb_ready = 1;
            }
        }
        private void System_Ready_Dis2(int stat)
        {
            if (stat == 0)
            {
                label_con_hyb.Text = "NG";
                label_con_hyb.BackColor = Color.Red;
                dut_ready = 0;
                Display_clear();
            }
            else
            {
                label_con_hyb.Text = "OK";
                label_con_hyb.BackColor = Color.Green;
                dut_ready = 1;
            }
        }


        private void button_Sol_N2_Open_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = 1;
            tkb_set_cmd[2] = Convert.ToByte(if_status_Sol_CDA);
            tkb_set_cmd[3] = Convert.ToByte(if_status_LM);
            tkb_set_cmd[4] = Convert.ToByte(if_status_LOOP);

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_Sol_N2_Close_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = 0;
            tkb_set_cmd[2] = Convert.ToByte(if_status_Sol_CDA);
            tkb_set_cmd[3] = Convert.ToByte(if_status_LM);
            tkb_set_cmd[4] = Convert.ToByte(if_status_LOOP);

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_Sol_CDA_Open_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = Convert.ToByte(if_status_Sol_N2);
            tkb_set_cmd[2] = 1;
            tkb_set_cmd[3] = Convert.ToByte(if_status_LM);
            tkb_set_cmd[4] = Convert.ToByte(if_status_LOOP);

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_Sol_CDA_Close_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = Convert.ToByte(if_status_Sol_N2);
            tkb_set_cmd[2] = 0;
            tkb_set_cmd[3] = Convert.ToByte(if_status_LM);
            tkb_set_cmd[4] = Convert.ToByte(if_status_LOOP);

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_LM_Open_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = Convert.ToByte(if_status_Sol_N2);
            tkb_set_cmd[2] = Convert.ToByte(if_status_Sol_CDA);
            tkb_set_cmd[3] = 1;
            tkb_set_cmd[4] = Convert.ToByte(if_status_LOOP);

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_LM_Close_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = Convert.ToByte(if_status_Sol_N2);
            tkb_set_cmd[2] = Convert.ToByte(if_status_Sol_CDA);
            tkb_set_cmd[3] = 0;
            tkb_set_cmd[4] = Convert.ToByte(if_status_LOOP);

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_LOOP_Open_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = Convert.ToByte(if_status_Sol_N2);
            tkb_set_cmd[2] = Convert.ToByte(if_status_Sol_CDA);
            tkb_set_cmd[3] = Convert.ToByte(if_status_LM);
            tkb_set_cmd[4] = 1;

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }

        private void button_LOOP_Close_Click(object sender, EventArgs e)
        {
            tkb_set_cmd[1] = Convert.ToByte(if_status_Sol_N2);
            tkb_set_cmd[2] = Convert.ToByte(if_status_Sol_CDA);
            tkb_set_cmd[3] = Convert.ToByte(if_status_LM);
            tkb_set_cmd[4] = 0;

            tkb_set_cmd[6] = 0;

            for (int i = 1; i < 6; i++)
            {
                tkb_set_cmd[6] ^= tkb_set_cmd[i];
            }

            if (serialPort1.IsOpen) serialPort1.Write(tkb_set_cmd, 0, tkb_set_cmd.Length);
        }
        private void Button_disconnect1_Click(object sender, EventArgs e)  //통신 연결끊기 버튼
        {
            if (serialPort1.IsOpen)  //시리얼포트가 열려 있으면
            {
                serialPort1.Close();  //시리얼포트 닫기

                System_Ready_Dis1(0);
                button_connect1.Enabled = true;
            }
        }

        private void Button_disconnect2_Click(object sender, EventArgs e)
        {
            if (serialPort2.IsOpen)  //시리얼포트가 열려 있으면
            {
                serialPort2.Close();  //시리얼포트 닫기

                System_Ready_Dis2(0);
                button_connect2.Enabled = true;
            }
            
        }

        private void button_Peak_Clear_Click(object sender, EventArgs e)
        {
            data_volt_current_peak = 0;
        }
        private void checkBox61_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox61.CheckState == CheckState.Checked)
            {
                checkBox51.CheckState = CheckState.Checked;
                checkBox52.CheckState = CheckState.Checked;
                checkBox53.CheckState = CheckState.Checked;
                checkBox54.CheckState = CheckState.Checked;
                checkBox55.CheckState = CheckState.Checked;
                checkBox56.CheckState = CheckState.Checked;
                checkBox57.CheckState = CheckState.Checked;
                checkBox58.CheckState = CheckState.Checked;
                checkBox59.CheckState = CheckState.Checked;
                checkBox60.CheckState = CheckState.Checked;
            }
            else if (checkBox61.CheckState == CheckState.Unchecked)
            {
                checkBox51.CheckState = CheckState.Unchecked;
                checkBox52.CheckState = CheckState.Unchecked;
                checkBox53.CheckState = CheckState.Unchecked;
                checkBox54.CheckState = CheckState.Unchecked;
                checkBox55.CheckState = CheckState.Unchecked;
                checkBox56.CheckState = CheckState.Unchecked;
                checkBox57.CheckState = CheckState.Unchecked;
                checkBox58.CheckState = CheckState.Unchecked;
                checkBox59.CheckState = CheckState.Unchecked;
                checkBox60.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox60_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox60.CheckState == CheckState.Checked)
            {
                checkBox46.CheckState = CheckState.Checked;
                checkBox47.CheckState = CheckState.Checked;
                checkBox48.CheckState = CheckState.Checked;
                checkBox49.CheckState = CheckState.Checked;
                checkBox50.CheckState = CheckState.Checked;
            }
            else if (checkBox60.CheckState == CheckState.Unchecked)
            {
                checkBox46.CheckState = CheckState.Unchecked;
                checkBox47.CheckState = CheckState.Unchecked;
                checkBox48.CheckState = CheckState.Unchecked;
                checkBox49.CheckState = CheckState.Unchecked;
                checkBox50.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox59_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox59.CheckState == CheckState.Checked)
            {
                checkBox41.CheckState = CheckState.Checked;
                checkBox42.CheckState = CheckState.Checked;
                checkBox43.CheckState = CheckState.Checked;
                checkBox44.CheckState = CheckState.Checked;
                checkBox45.CheckState = CheckState.Checked;
            }
            else if (checkBox59.CheckState == CheckState.Unchecked)
            {
                checkBox41.CheckState = CheckState.Unchecked;
                checkBox42.CheckState = CheckState.Unchecked;
                checkBox43.CheckState = CheckState.Unchecked;
                checkBox44.CheckState = CheckState.Unchecked;
                checkBox45.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox58_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox58.CheckState == CheckState.Checked)
            {
                checkBox36.CheckState = CheckState.Checked;
                checkBox37.CheckState = CheckState.Checked;
                checkBox38.CheckState = CheckState.Checked;
                checkBox39.CheckState = CheckState.Checked;
                checkBox40.CheckState = CheckState.Checked;
            }
            else if (checkBox58.CheckState == CheckState.Unchecked)
            {
                checkBox36.CheckState = CheckState.Unchecked;
                checkBox37.CheckState = CheckState.Unchecked;
                checkBox38.CheckState = CheckState.Unchecked;
                checkBox39.CheckState = CheckState.Unchecked;
                checkBox40.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox57_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox57.CheckState == CheckState.Checked)
            {
                checkBox31.CheckState = CheckState.Checked;
                checkBox32.CheckState = CheckState.Checked;
                checkBox33.CheckState = CheckState.Checked;
                checkBox34.CheckState = CheckState.Checked;
                checkBox35.CheckState = CheckState.Checked;
            }
            else if (checkBox57.CheckState == CheckState.Unchecked)
            {
                checkBox31.CheckState = CheckState.Unchecked;
                checkBox32.CheckState = CheckState.Unchecked;
                checkBox33.CheckState = CheckState.Unchecked;
                checkBox34.CheckState = CheckState.Unchecked;
                checkBox35.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox56_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox56.CheckState == CheckState.Checked)
            {
                checkBox26.CheckState = CheckState.Checked;
                checkBox27.CheckState = CheckState.Checked;
                checkBox28.CheckState = CheckState.Checked;
                checkBox29.CheckState = CheckState.Checked;
                checkBox30.CheckState = CheckState.Checked;
            }
            else if (checkBox56.CheckState == CheckState.Unchecked)
            {
                checkBox26.CheckState = CheckState.Unchecked;
                checkBox27.CheckState = CheckState.Unchecked;
                checkBox28.CheckState = CheckState.Unchecked;
                checkBox29.CheckState = CheckState.Unchecked;
                checkBox30.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox55_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox55.CheckState == CheckState.Checked)
            {
                checkBox21.CheckState = CheckState.Checked;
                checkBox22.CheckState = CheckState.Checked;
                checkBox23.CheckState = CheckState.Checked;
                checkBox24.CheckState = CheckState.Checked;
                checkBox25.CheckState = CheckState.Checked;
            }
            else if (checkBox55.CheckState == CheckState.Unchecked)
            {
                checkBox21.CheckState = CheckState.Unchecked;
                checkBox22.CheckState = CheckState.Unchecked;
                checkBox23.CheckState = CheckState.Unchecked;
                checkBox24.CheckState = CheckState.Unchecked;
                checkBox25.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox54_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox54.CheckState == CheckState.Checked)
            {
                checkBox16.CheckState = CheckState.Checked;
                checkBox17.CheckState = CheckState.Checked;
                checkBox18.CheckState = CheckState.Checked;
                checkBox19.CheckState = CheckState.Checked;
                checkBox20.CheckState = CheckState.Checked;
            }
            else if (checkBox54.CheckState == CheckState.Unchecked)
            {
                checkBox16.CheckState = CheckState.Unchecked;
                checkBox17.CheckState = CheckState.Unchecked;
                checkBox18.CheckState = CheckState.Unchecked;
                checkBox19.CheckState = CheckState.Unchecked;
                checkBox20.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox53_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox53.CheckState == CheckState.Checked)
            {
                checkBox11.CheckState = CheckState.Checked;
                checkBox12.CheckState = CheckState.Checked;
                checkBox13.CheckState = CheckState.Checked;
                checkBox14.CheckState = CheckState.Checked;
                checkBox15.CheckState = CheckState.Checked;
            }
            else if (checkBox53.CheckState == CheckState.Unchecked)
            {
                checkBox11.CheckState = CheckState.Unchecked;
                checkBox12.CheckState = CheckState.Unchecked;
                checkBox13.CheckState = CheckState.Unchecked;
                checkBox14.CheckState = CheckState.Unchecked;
                checkBox15.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox52_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox52.CheckState == CheckState.Checked)
            {
                checkBox6.CheckState = CheckState.Checked;
                checkBox7.CheckState = CheckState.Checked;
                checkBox8.CheckState = CheckState.Checked;
                checkBox9.CheckState = CheckState.Checked;
                checkBox10.CheckState = CheckState.Checked;
            }
            else if (checkBox52.CheckState == CheckState.Unchecked)
            {
                checkBox6.CheckState = CheckState.Unchecked;
                checkBox7.CheckState = CheckState.Unchecked;
                checkBox8.CheckState = CheckState.Unchecked;
                checkBox9.CheckState = CheckState.Unchecked;
                checkBox10.CheckState = CheckState.Unchecked;
            }
        }

        private void checkBox51_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox51.CheckState == CheckState.Checked)
            {
                checkBox1.CheckState = CheckState.Checked;
                checkBox2.CheckState = CheckState.Checked;
                checkBox3.CheckState = CheckState.Checked;
                checkBox4.CheckState = CheckState.Checked;
                checkBox5.CheckState = CheckState.Checked;
            }
            else if (checkBox51.CheckState == CheckState.Unchecked)
            {
                checkBox1.CheckState = CheckState.Unchecked;
                checkBox2.CheckState = CheckState.Unchecked;
                checkBox3.CheckState = CheckState.Unchecked;
                checkBox4.CheckState = CheckState.Unchecked;
                checkBox5.CheckState = CheckState.Unchecked;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)  //시리얼포트가 열려 있으면
            {
                serialPort1.Close();  //시리얼포트 닫기
            }

            if (serialPort2.IsOpen)  //시리얼포트가 열려 있으면
            {
                serialPort2.Close();  //시리얼포트 닫기
            }
            if (MessageBox.Show("윈도우도 같이 종료 하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Close();
                Process.Start("shutdown.exe", "-s -t 10");
            }
            else
            {
                this.Close();
            }
            this.Close();
        }
    }
    
}
