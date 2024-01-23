using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;

using CheckBox = System.Windows.Forms.CheckBox;
using Point = System.Drawing.Point;

// list:
// - pv_p eff 統一由JQ計算
// - 缺少Insulation: 
// 新旺: 缺少-event、-dcp、-freq要另計、線電壓要另計、powerRST另計
// - [etoday]特例處理: solarEdge、ABB
// 奇景:SF相乘
// - 亞力: 無告警事件
// - JQ 電表無累積發電量欄位
// 

//3.1 電表設備模板

//4 匯出通訊協定書


namespace JQ自動生成器_v1
{
    public partial class Form1 : Form
    {
        int INV_QUANTITY = 0; // INV 數量
        int PV_QUANTITY = 0; // PV串列 數量
        int IRR_QUANTITY = 0; // IRR 數量
        string CAPACITY = "[0,0,0,0]"; // 裝置容量
        string FACTORYNAME = "AAAA";
        string FACTORYID = "12345678";
        string DREAMS_NAME = "dreams";
        // int INVTEMP_QUANTITY = 1; // INV_TEMP 數量
        int invModelkExport = 0; // 匯出INV模板編號
        int invJQType = 0; // 匯出JQ模板類型

        private CheckBox[] allCheckboxes;
        private CheckBox sender;

        public Form1() {
            InitializeComponent();
            // groupbox2初始化位置
            groupBox2.Location = new Point(140, groupBox2.Location.Y);
            groupBox2.Location = new Point(groupBox2.Location.X, 12);
            // checkbox數量表
            allCheckboxes = new CheckBox[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox8, checkBox9, checkBox10, checkBox11, checkBox12, checkBox13, checkBox14, checkBox15, checkBox16 };
        }

        private void Form1_Load(object sender, EventArgs e) {
            this.Size = new Size(586, 400);
            //this.Size = new Size(573, this.Size.Height);
            //this.Size = new Size(this.Size.Width, 360);
        }
        // 匯出JQ
        private async void button1_Click(object sender, EventArgs e) {
            // 讀取參數
            CAPACITY = textBox6.Text;
            FACTORYNAME = textBox1.Text;
            FACTORYID = textBox2.Text;
            DREAMS_NAME = textBox7.Text;

            if (checkBox23.Checked != true)
            {
                CAPACITY = "[0]";
                FACTORYNAME = textBox7.Text;
                textBox3.Text = "1";
                textBox4.Text = "1";
                textBox5.Text = "0";
                comboBox1.Text = "None";
                comboBox8.Text = "無";
            }
            if(checkBox24.Checked != true)
            {
                comboBox9.Text = "電錶來源";
            }

            INV_QUANTITY = int.Parse(textBox3.Text);
            PV_QUANTITY = int.Parse(textBox4.Text);
            IRR_QUANTITY = int.Parse(textBox5.Text);

            string text = "";
            string[] invLabel = new string[50];  // INV 告警事件表
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "JQ_Format.txt");
            string[] strings = new string[120];
            string[] indexes = new string[INV_QUANTITY];
            string[] indexes2 = new string[PV_QUANTITY];
            string[] indexes3 = new string[IRR_QUANTITY];

            // INV 告警表
            // Delta EventIndexes
            invLabel[1] = "AlarmIndexes:[\"alarm_E1\",\"alarm_E2\",\"alarm_E3\",\"alarm_W1\",\"alarm_W2\",\"alarm_F1\",\"alarm_F2\",\"alarm_F3\",\"alarm_F4\",\"alarm_F5\"]";
            // SolarEdge EventIndexes
            invLabel[2] = "AlarmIndexes:[\"event\"]";
            // HUAWEI-KTL EventIndexes
            invLabel[3] = "AlarmIndexes:[\"event_1\",\"event_2\",\"event_3\"]";
            // HUAWEI-KTL(OLD)Indexes
            invLabel[4] = "AlarmIndexes:[\"event_1\",\"event_2\",\"event_3\",\"event_4\",\"event_5\",\"event_6\",\"event_7\",\"event_8\",\"event_9\",\"event_10\",\"event_11\"]";
            // SUNGROW_SG110CX
            invLabel[5] = "AlarmIndexes:[\"event\"]";
            // SCHNEIDERTL20000E
            invLabel[6] = "AlarmIndexes:[\"event_1\",\"event_2\",\"event_3\",\"event_4\"]";
            // ABBPVS100_120
            invLabel[7] = "AlarmIndexes:[\"event_1\",\"eventVendor_1\",\"eventVendor_2\",\"eventVendor_3\"]";
            // None ALLIS
            invLabel[8] = "AlarmIndexes:[]";
            // PrimeVOLT PV60000T
            invLabel[9] = "AlarmIndexes:[\"error_1\",\"error_2\",\"error_3\"]";


            switch (comboBox1.Text) {
                case "None":
                    invLabel[0] = invLabel[8];
                    invJQType = 0;
                    break;
                case "DELTA":
                    invLabel[0] = invLabel[1];
                    invJQType = 1;
                    break;
                case "HUAWEI-KTL":
                    invLabel[0] = invLabel[4];
                    invJQType = 2;
                    break;
                case "HUAWEI-KTL(舊版本)":
                    invLabel[0] = invLabel[3];
                    invJQType = 3;
                    break;
                case "SolarEdge":
                    invLabel[0] = invLabel[2];
                    invJQType = 4;
                    break;
                case "SUNGROW_SG110CX":
                    invLabel[0] = invLabel[5];
                    invJQType = 5;
                    break;
                case "ABBPVS100_120":
                    invLabel[0] = invLabel[7];
                    invJQType = 6;
                    break;
                case "SCHNEIDERTL20000E":
                    invLabel[0] = invLabel[6];
                    invJQType = 7;
                    break;
                case "ALLIS(無提供告警事件)":
                    invLabel[0] = invLabel[8];
                    invJQType = 8;
                    break;
                case "PrimeVOLT PV60000T":
                    invLabel[0] = invLabel[9];
                    invJQType = 9;
                    break;
                default:
                    invLabel[0] = "\n";
                    break;
            }

            for (int i = 0; i < INV_QUANTITY; i++) indexes[i] = (i + 1).ToString("D2");  //  1 -> "01"

            for (int i = 0; i < PV_QUANTITY; i++) {
                indexes2[i] = (i + 1).ToString();
            }
            for (int i = 0; i < IRR_QUANTITY; i++) {
                indexes3[i] = (i + 1).ToString("D2");  // 1 -> "01"
            }

            // def 
            strings[0] = "def FormatFloat: .*100|floor|./100;\n";
            strings[1] = "def IRRFormatFloat: .*1000|floor|./1000;\n";
            strings[2] = "def GenerateCapacity($source_indexes):map($source_indexes);\n";
            strings[3] = "def GenerateINVStatus($source_prefix; $source_indexes; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status !=0 then \"設備斷訊\" else 0 end);\n";
            strings[4] = "def GenerateStatus($source_prefix; $data):(if $data.\"\\($source_prefix)\".status !=0 then \"設備斷訊\" else 0 end);\n";
            strings[5] = "def EventOneArray($source_prefix; $event_indexes; $data):$event_indexes | map(if $data.\"\\($source_prefix)\".status == 0 then $data.\"\\($source_prefix)\".\"\\((.|tostring))\" else -1 end);\n";
            strings[6] = "def GenerateEvent($source_prefix; $source_indexes; $a; $data):$source_indexes | map(if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then $data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\" | FormatFloat else 0 end);\n";
            strings[7] = "def Generate($source; $tag_prefix; $data):(if $data.\"\\($source)\".status == 0 then(if ($data.\"\\($source)\".\"\\($tag_prefix)\" ) != null then($data.\"\\($source)\".\"\\($tag_prefix)\") | FormatFloat else ($data.\"\\($source)\".\"\\($tag_prefix)\")end)else 0 end);\n";
            strings[8] = "def GenerateConstSourceByIndex($source; $tag_prefix; $tag_indexes; $data):(($tag_indexes | map( if $data.\"\\($source)\".status == 0 then ( if ($data.\"\\($source)\".\"\\($tag_prefix + (.| tostring))\") != null then $data.\"\\($source)\".\"\\($tag_prefix + (.| tostring))\"| FormatFloat else ($data.\"\\($source)\".\"\\($tag_prefix + (.| tostring))\")end )else 0 end )));\n";
            strings[9] = "def GenerateByIndex($source_prefix; $source_indexes; $tag_prefix; $tag_indexes; $data):$source_indexes | map(GenerateConstSourceByIndex($source_prefix+(.|tostring); $tag_prefix; $tag_indexes; $data));\n";
            strings[10] = "def GenerateA($source_prefix; $source_indexes; $a; $data):$source_indexes | map( if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then( if ($data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) > 0 then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" | FormatFloat else 0 end )else 0 end);\n";
            strings[11] = "def GenerateIRRArray($source_prefix; $source_indexes; $a; $data):$source_indexes | map( if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then( if ($data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) > 0 then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" | IRRFormatFloat else 0 end )else 0 end);\n";
            strings[12] = "def GenerateABDivisor($source_prefix; $source_indexes; $a; $b; $divisor; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then(if (( $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\") != 0) then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" * $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\" / $divisor | FormatFloat else 0 end)else 0 end);\n";
            strings[13] = "def GenerateACrossor($source_prefix; $source_indexes; $a; $crossor; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" * $crossor | FormatFloat else 0 end);\n";
            strings[14] = "def GeneratetempSourceByIndex($source; $tag_prefix; $tag_indexes; $data):$tag_indexes | map(if $data.\"\\($source)\".status == 0 then $data.\"\\($source)\".\"\\($tag_prefix)\" | FormatFloat else 0 end);\n";
            strings[15] = "def GenerateABCrossor($source; $tag_prefix_a; $tag_prefix_b; $tag_indexes; $data):(($tag_indexes | map( if $data.\"\\($source)\".status == 0 then (if ($data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\") != null then $data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\" * $data.\"\\($source)\".\"\\($tag_prefix_b + (.tostring))\"| FormatFloat else \" - 1\" end )else 0 end)) | map(select(. != \" - 1\")));\n";
            strings[16] = "def GenerateCalculateABCrossorByIndex($source_prefix; $source_indexes; $tag_a; $tag_b;$tag_indexes; $data):$source_indexes | map(GenerateABCrossor($source_prefix+(.|tostring); $tag_a; $tag_b; $tag_indexes; $data));\n";            
            strings[17] = "def GenerateABCrossoranddivide1000($source; $tag_prefix_a; $tag_prefix_b; $tag_indexes; $data):(($tag_indexes | map(if $data.\"\\($source)\".status == 0 then (if ($data.\"\\($source)\".\"\\($tag_prefix_a+(.|tostring))\") and ($data.\"\\($source)\".\"\\($tag_prefix_b+(.|tostring))\") != null then $data.\"\\($source)\".\"\\($tag_prefix_a+(.|tostring))\" * $data.\"\\($source)\".\"\\($tag_prefix_b+(.|tostring))\" / 1000 | FormatFloat else($data.\"\\($source)\".\"\\($tag_prefix_a+(.|tostring))\")end )else 0 end)));\n";
            strings[18] = "def GenerateCalculateABCrossorByIndexanddivide1000($source_prefix; $source_indexes; $tag_a;$tag_b; $tag_indexes; $data):$source_indexes | map(GenerateABCrossoranddivide1000($source_prefix+(.|tostring); $tag_a; $tag_b; $tag_indexes;$data));\n";
            strings[19] = "def GenerateAMinusB($source_prefix; $source_indexes; $a; $b; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then(if (( $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix + (.tostring))\".\"\\($b)\") != 0) then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" - $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\"| FormatFloat else 0 end)else 0 end);\n";
            strings[20] = "def GenerateADivideB($source_prefix; $source_indexes; $a; $b; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then (if (( $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\") != 0) then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" / $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\" | FormatFloat else 0 end)else 0 end);\n";
            strings[21] = "def GenerateInsulationA($source_prefix; $source_indexes; $a; $data): $source_indexes | map( if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then( if ($data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) > 0 then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" | IRRFormatFloat else 0 end )else 0 end);\n";
            strings[22] = "def GenerateEvent($source_prefix; $source_indexes; $tag_indexes; $data):$source_indexes | map(GenerateEventB($source_prefix+(.|tostring); $tag_indexes; $data));\n";
            strings[23] = "def GenerateCrossor($source; $tag_prefix; $crossor; $data):(if $data.\"\\($source)\".status == 0 then(if ($data.\"\\($source)\".\"\\($tag_prefix)\" ) != null then($data.\"\\($source)\".\"\\($tag_prefix)\")*$crossor | FormatFloat else($data.\"\\($source)\".\"\\($tag_prefix)\")end)else 0 end);\n";
            strings[24] = "def GenerateEventT($source_prefix; $source_indexes; $tag_indexes; $data):$source_indexes | map(EventOneArray($source_prefix+(.|tostring); $tag_indexes; $data));\n";
            strings[25] = "def GenerateACrosserSF($source_prefix; $source_indexes; $a; $b; $data): $source_indexes | map( if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then ( if (( $data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\" ) != 0) then (($data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\" ))) | FormatFloat else 0 end ) else 0 end );\n";
            strings[26] = "def GeneratePrimeVoltdcp($source_prefix; $source_indexes; $a; $b; $c; $d; $data): $source_indexes | map( if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then (($data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\") + ($data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\") + ($data.\"\\($source_prefix+(.|tostring))\".\"\\($c)\") + ($data.\"\\($source_prefix+(.|tostring))\".\"\\($d)\")) | FormatFloat else 0 end );\n";
            strings[27] = "def GeneratePrimeVoltfreq($source_prefix; $source_indexes; $a; $b; $c; $data): $source_indexes | map( if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then ((($data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\") + ($data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\") + ($data.\"\\($source_prefix+(.|tostring))\".\"\\($c)\"))/3 ) | FormatFloat else 0 end );\n";
            strings[28] = "def GenerateConstSourceByIndexCrosserSF($source; $tag_prefix; $tag_indexes; $a; $data):(($tag_indexes | map( if $data.\"\\($source)\".status == 0 then (if ($data.\"\\($source)\".\"\\($tag_prefix+(.|tostring))\") != null then ($data.\"\\($source)\".\"\\($tag_prefix+(.|tostring))\")*pow(10;($data.\"\\($source)\".\"\\($a)\"))| FormatFloat else ($data.\"\\($source)\".\"\\($tag_prefix+(.|tostring))\") end )else 0 end)));\n";
            strings[29] = "def GenerateByIndexCrosserSF($source_prefix; $source_indexes; $tag_prefix; $tag_indexes; $a; $data):$source_indexes | map( GenerateConstSourceByIndexCrosserSF($source_prefix+(.|tostring); $tag_prefix; $tag_indexes; $a; $data));\n";
            strings[30] = "def GenerateACrosserBCrosserSF($source_prefix; $source_indexes; $a; $b; $c; $d; $data):$source_indexes | map(if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then(if (( $data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\") != 0)then(($data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($c)\"))) * (($data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($d)\")))/1000 | FormatFloat else 0 end)else 0 end);\n";
            strings[31] = "def GenerateADivideBCrosserSF($source_prefix; $source_indexes; $a; $b; $c; $d; $data):$source_indexes | map(if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then (if (( $data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\") != 0) then (($data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($c)\"))) / (($data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($d)\"))) | FormatFloat else 0 end)else 0 end);\n";
            strings[32] = "def GenerateArrayAMinusBCrosserSF($source_prefix; $source_indexes; $a; $b; $c; $data):$source_indexes | map(if $data.\"\\($source_prefix+(.|tostring))\".status == 0 then(if (( $data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\") != 0)then(($data.\"\\($source_prefix+(.|tostring))\".\"\\($a)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($c)\"))) - (($data.\"\\($source_prefix+(.|tostring))\".\"\\($b)\")*pow(10;($data.\"\\($source_prefix+(.|tostring))\".\"\\($c)\"))) | FormatFloat else 0 end)else 0 end);\n";
            // []index定義 
            strings[33] = "\n\n{\ninverterIndexes: [\"" + string.Join("\",\"", indexes) + "\"],\n";
            strings[34] = "PVIndexes: [\"" + string.Join("\",\"", indexes2) + "\"],\n";
            strings[35] = "iRRIndexes: [\"" + string.Join("\",\"", indexes3) + "\"],\n";
            strings[36] = "invTempIndexes: [1],\n"; // inv temp 數量固定為1
            strings[37] = "capacityIndexes: " + CAPACITY + ",\n";
            strings[38] = invLabel[0];
            strings[39] = ",\n";
            strings[40] = "data: .\n} |";

            // main payload
            strings[42] = "\n\n{\nfactoryName: \"" + FACTORYNAME + "\",";
            strings[43] = "\ntimestamp: (now+28800|strftime(\"%Y-%m-%dT%H:%M:%SZ\")),";
            strings[44] = "\ndetail:\n {"; // detail{}
            strings[45] = "\n factoryId: \"" + FACTORYID + "\"";
            strings[46] = ",\n factoryName: \"" + FACTORYNAME + "\"";
            strings[47] = ",\n pv_v: GenerateByIndex(\"inv_\"; .inverterIndexes; \"pv_v\"; .PVIndexes; .data)";
            strings[48] = ",\n pv_a: GenerateByIndex(\"inv_\"; .inverterIndexes; \"pv_a\"; .PVIndexes; .data)";
            // pv_p統一由 pv_v * pv_a / 1000
            strings[49] = ",\n pv_p: GenerateCalculateABCrossorByIndexanddivide1000(\"inv_\"; .inverterIndexes; \"pv_a\"; \"pv_v\"; .PVIndexes; .data)";
            //strings[49] = ",\n pv_p: GenerateByIndex(\"inv_\"; .inverterIndexes; \"pv_p\"; .PVIndexes; .data)";
            strings[50] = ",\n PF: GenerateA(\"inv_\"; .inverterIndexes; \"PF\"; .data)";
            strings[51] = ",\n Vrn: GenerateA(\"inv_\"; .inverterIndexes; \"Vrs\"; .data)";
            strings[52] = ",\n Vsn: GenerateA(\"inv_\"; .inverterIndexes; \"Vst\"; .data)";
            strings[53] = ",\n Vtn: GenerateA(\"inv_\"; .inverterIndexes; \"Vrt\"; .data)";
            strings[54] = ",\n Rc: GenerateA(\"inv_\"; .inverterIndexes; \"Rc\"; .data)";
            strings[55] = ",\n Sc: GenerateA(\"inv_\"; .inverterIndexes; \"Sc\"; .data)";
            strings[56] = ",\n Tc: GenerateA(\"inv_\"; .inverterIndexes; \"Tc\"; .data)";
            strings[57] = ",\n temp: GenerateByIndex(\"inv_\"; .inverterIndexes; \"temp_\"; .invTempIndexes; .data)";
            strings[58] = ",\n State: GenerateA(\"inv_\"; .inverterIndexes; \"State\"; .data)";
            strings[59] = ",\n acp: GenerateA(\"inv_\"; .inverterIndexes; \"acp\"; .data)";
            strings[60] = ",\n dcp: GenerateA(\"inv_\"; .inverterIndexes; \"dcp\"; .data)";
            //EFF統一由 DCP/ACP
            strings[61] = ",\n Eff: GenerateADivideB(\"inv_\"; .inverterIndexes; \"acp\"; \"dcp\"; .data)";
            strings[62] = ",\n AC_kWh: GenerateA(\"inv_\"; .inverterIndexes; \"AC_kWh\"; .data)";
            strings[63] = ",\n freq: GenerateA(\"inv_\"; .inverterIndexes; \"freq\"; .data)";
            //strings[64] = ",\n event: GenerateEvent(\"inv_\"; .inverterIndexes; .AlarmIndexes; .data)";
            strings[65] = ",\n event: GenerateEventT(\"inv_\"; .inverterIndexes; .AlarmIndexes; .data)";
            strings[66] = ",\n Insulation: GenerateA(\"inv_\"; .inverterIndexes; \"Insulation\"; .data)";
            strings[68] = ",\n E_today: GenerateA(\"inv_\"; .inverterIndexes; \"E_today\"; .data)";
            strings[69] = ",\n capacity: (.capacityIndexes)";
            strings[70] = ",\n power_r: GenerateABCrossor(\"inv_\"; \"Vrs\"; \"Rc\"; .inverterIndexes; .data)";
            strings[71] = ",\n power_s: GenerateABCrossor(\"inv_\"; \"Vst\"; \"Sc\"; .inverterIndexes; .data)";
            strings[72] = ",\n power_t: GenerateABCrossor(\"inv_\"; \"Vrt\"; \"Tc\"; .inverterIndexes; .data)";
            if (comboBox8.Text == "無") strings[80] = ",\n IRR: [null]";
            else strings[80] = ",\n IRR: GenerateIRRArray(\"IRR_\"; .iRRIndexes; \"IRR\"; .data)";
            // 日照計
            if (checkBox18.Checked == true) strings[81] = ",\n PVTemp: Generate(\"PV_TEMP\"; \"temp\"; .data)";
            if (checkBox17.Checked == true) strings[82] = ",\n ENVTemp: Generate(\"ENV_TEMP\"; \"temp\"; .data)";
            // 風速計
            if (checkBox19.Checked == true) strings[83] = ",\n Anemometer: Generate(\"Anemometer\"; \"Wind_speed\"; .data)";
            // 水位計
            if (checkBox20.Checked == true) strings[84] = ",\n Waterlevel: Generate(\"Waterlevel\"; \"Measurement_output_value\"; .data)";
            // Errormessage{}
            strings[85] = ",\n ErrorMessage: {";
            strings[86] = "\n  inv: GenerateINVStatus(\"inv_\"; .inverterIndexes; .data)";
            if (comboBox8.Text == "無") strings[87] = ",\n  IRR: [null]";
            else strings[87] = ",\n  IRR: GenerateINVStatus(\"IRR_\"; .iRRIndexes; .data)";
            if (checkBox18.Checked == true) strings[88] = ",\n  PVTemp: GenerateStatus(\"PV_TEMP\";.data)";
            if (checkBox17.Checked == true) strings[89] = ",\n  ENVtemp: GenerateStatus(\"ENV_TEMP\";.data)";
            if (checkBox19.Checked == true)
            { // 風速計 (ErrorMessage)
                strings[90] = ",\n  Anemometer: GenerateStatus(\"Anemometer\";.data)";
            }
            if (checkBox20.Checked == true)
            { // 水位計 (ErrorMessage)
                strings[91] = ",\n  Waterlevel: GenerateStatus(\"Waterlevel\";.data)";
            }
            if (checkBox21.Checked == true)
            { // 低壓電錶 (ErrorMessage)
                strings[92] = ",\n  LV_Meter: GenerateStatus(\"LV_meter\";.data)";
            }
            if (checkBox22.Checked == true)
            { // 高壓電錶 (ErrorMessage)
                strings[93] = ",\n  HV_Meter: GenerateStatus(\"HV_meter\";.data)";
            }
            strings[94] = "\n  }"; // Errormessage end{}

            strings[95] = ",\n etc:\n {\n"; // etc{}
            // 低壓電錶(value)
            if (checkBox21.Checked == true) strings[96] = "\n  \"LV-meter\": {\n   Vln_a: Generate(\"LV_meter\"; \"Vln_a\"; .data),\n   Vln_b: Generate(\"LV_meter\"; \"Vln_b\"; .data),\n   Vln_c: Generate(\"LV_meter\"; \"Vln_c\"; .data),\n   Vll_ab: Generate(\"LV_meter\"; \"Vll_ab\"; .data),\n   Vll_bc: Generate(\"LV_meter\"; \"Vll_bc\"; .data),\n   Vll_ca: Generate(\"LV_meter\"; \"Vll_ca\"; .data),\n   I_a: Generate(\"LV_meter\"; \"I_a\"; .data),\n   I_b: Generate(\"LV_meter\"; \"I_b\"; .data),\n   I_c: Generate(\"LV_meter\"; \"I_c\"; .data),\n   Freq: Generate(\"LV_meter\"; \"freq\"; .data),\n   P: Generate(\"LV_meter\"; \"P\"; .data),\n   kWh: Generate(\"LV_meter\"; \"AC_kWh\"; .data),\n   KVAR_tot: Generate(\"LV_meter\"; \"Q\"; .data),\n   KVA_tot: Generate(\"LV_meter\"; \"S\"; .data)\n   }";
            // 高壓電錶(value)
            if (checkBox22.Checked == true) {
                if (checkBox21.Checked == true) strings[97] = ",";
                strings[98] = "\n  \"HV-meter\": {\n   Vln_a: Generate(\"HV_meter\"; \"Vln_a\"; .data),\n   Vln_b: Generate(\"HV_meter\"; \"Vln_b\"; .data),\n   Vln_c: Generate(\"HV_meter\"; \"Vln_c\"; .data),\n   Vll_ab: Generate(\"HV_meter\"; \"Vll_ab\"; .data),\n   Vll_bc: Generate(\"HV_meter\"; \"Vll_bc\"; .data),\n   Vll_ca: Generate(\"HV_meter\"; \"Vll_ca\"; .data),\n   I_a: Generate(\"HV_meter\"; \"I_a\"; .data),\n   I_b: Generate(\"HV_meter\"; \"I_b\"; .data),\n   I_c: Generate(\"HV_meter\"; \"I_c\"; .data),\n   Freq: Generate(\"HV_meter\"; \"freq\"; .data),\n   P: Generate(\"HV_meter\"; \"P\"; .data),\n   kWh: Generate(\"HV_meter\"; \"AC_kWh\"; .data),\n   KVAR_tot: Generate(\"HV_meter\"; \"Q\"; .data),\n   KVA_tot: Generate(\"HV_meter\"; \"S\"; .data)\n   }";
            }
            // DREAMS電錶(value)
            if (checkBox24.Checked == true)
            {
                if (comboBox9.Text == "低壓電錶")
                {
                    if (checkBox21.Checked == true || checkBox22.Checked == true) strings[99] = ",";
                    strings[100] = "\n \"DREAMS-meter\": \n{\r\n   DREAMS_name: \"" + DREAMS_NAME + "\",\r\n   currentPhaseA: GenerateCrossor(\"LV_meter\"; \"I_a\"; 10; .data),\r\n   currentPhaseB: GenerateCrossor(\"LV_meter\"; \"I_b\"; 10; .data),\r\n   currentPhaseC: GenerateCrossor(\"LV_meter\"; \"I_c\"; 10; .data),\r\n   currentPhaseN: 0,\r\n   voltagePhaseA: GenerateCrossor(\"LV_meter\"; \"Vll_ab\"; 100; .data),\r\n   voltagePhaseB: GenerateCrossor(\"LV_meter\"; \"Vll_bc\"; 100; .data),\r\n   voltagePhaseC: GenerateCrossor(\"LV_meter\"; \"Vll_ca\"; 100; .data),\r\n   control_result_01_25: [{\"pf\":0,\"p\":0},{\"pf\":0,\"p\":0}],\r\n   control_result_26_50: [{\"pf\":0,\"p\":0},{\"pf\":0,\"p\":0}],\r\n   p_sum: GenerateCrossor(\"LV_meter\"; \"P\"; 1000; .data),\r\n   q_sum: GenerateCrossor(\"LV_meter\"; \"Q\"; 1000; .data),\r\n   pf_avg: GenerateCrossor(\"LV_meter\"; \"PF\"; 100; .data),\r\n   frequency: GenerateCrossor(\"LV_meter\"; \"freq\"; 10; .data),\r\n   total_kWh: GenerateCrossor(\"LV_meter\"; \"AC_kWh\"; 1000; .data),\r\n   irradiance: Generate(\"IRR_\"; \"IRR_01\"; .data),\r\n   p_setting:0,\r\n   q_setting:0,\r\n   pf_setting:0,\r\n   vpset_setting:0,\r\n   itemTimestamp: (now-(now%60)| floor),\r\n   timestamp: (now+28800|strftime(\"%Y-%m-%dT%H:%M:%SZ\"))\r\n  }";
                }
                else if (comboBox9.Text == "高壓電錶")
                {
                    if (checkBox21.Checked == true || checkBox22.Checked == true) strings[99] = ",";
                    strings[100] = "\n \"DREAMS-meter\": \n{\r\n   DREAMS_name: \"" + DREAMS_NAME + "\",\r\n   currentPhaseA: GenerateCrossor(\"HV_meter\"; \"I_a\"; 10; .data),\r\n   currentPhaseB: GenerateCrossor(\"HV_meter\"; \"I_b\"; 10; .data),\r\n   currentPhaseC: GenerateCrossor(\"HV_meter\"; \"I_c\"; 10; .data),\r\n   currentPhaseN: 0,\r\n   voltagePhaseA: GenerateCrossor(\"HV_meter\"; \"Vll_ab\"; 100; .data),\r\n   voltagePhaseB: GenerateCrossor(\"HV_meter\"; \"Vll_bc\"; 100; .data),\r\n   voltagePhaseC: GenerateCrossor(\"HV_meter\"; \"Vll_ca\"; 100; .data),\r\n   control_result_01_25: [{\"pf\":0,\"p\":0},{\"pf\":0,\"p\":0}],\r\n   control_result_26_50: [{\"pf\":0,\"p\":0},{\"pf\":0,\"p\":0}],\r\n   p_sum: GenerateCrossor(\"HV_meter\"; \"P\"; 1000; .data),\r\n   q_sum: GenerateCrossor(\"HV_meter\"; \"Q\"; 1000; .data),\r\n   pf_avg: GenerateCrossor(\"HV_meter\"; \"PF\"; 100; .data),\r\n   frequency: GenerateCrossor(\"HV_meter\"; \"freq\"; 10; .data),\r\n   total_kWh: GenerateCrossor(\"HV_meter\"; \"AC_kWh\"; 1000; .data),\r\n   irradiance: Generate(\"IRR_\"; \"IRR_01\"; .data),\r\n   p_setting:0,\r\n   q_setting:0,\r\n   pf_setting:0,\r\n   vpset_setting:0,\r\n   itemTimestamp: (now-(now%60)| floor),\r\n   timestamp: (now+28800|strftime(\"%Y-%m-%dT%H:%M:%SZ\"))\r\n  }";
                }
                else if (comboBox9.Text == "獨立電錶")
                {
                    if (checkBox21.Checked == true || checkBox22.Checked == true) strings[99] = ",";
                    strings[100] = "\n \"DREAMS-meter\": \n{\r\n   DREAMS_name: \"" + DREAMS_NAME + "\",\r\n   currentPhaseA: GenerateCrossor(\"DREAMS_meter\"; \"I_a\"; 10; .data),\r\n   currentPhaseB: GenerateCrossor(\"DREAMS_meter\"; \"I_b\"; 10; .data),\r\n   currentPhaseC: GenerateCrossor(\"DREAMS_meter\"; \"I_c\"; 10; .data),\r\n   currentPhaseN: 0,\r\n   voltagePhaseA: GenerateCrossor(\"DREAMS_meter\"; \"Vll_ab\"; 100; .data),\r\n   voltagePhaseB: GenerateCrossor(\"DREAMS_meter\"; \"Vll_bc\"; 100; .data),\r\n   voltagePhaseC: GenerateCrossor(\"DREAMS_meter\"; \"Vll_ca\"; 100; .data),\r\n   control_result_01_25: [{\"pf\":0,\"p\":0},{\"pf\":0,\"p\":0}],\r\n   control_result_26_50: [{\"pf\":0,\"p\":0},{\"pf\":0,\"p\":0}],\r\n   p_sum: GenerateCrossor(\"DREAMS_meter\"; \"P\"; 1000; .data),\r\n   q_sum: GenerateCrossor(\"DREAMS_meter\"; \"Q\"; 1000; .data),\r\n   pf_avg: GenerateCrossor(\"DREAMS_meter\"; \"PF\"; 100; .data),\r\n   frequency: GenerateCrossor(\"DREAMS_meter\"; \"freq\"; 10; .data),\r\n   total_kWh: GenerateCrossor(\"DREAMS_meter\"; \"AC_kWh\"; 1000; .data),\r\n   irradiance: Generate(\"IRR_\"; \"IRR_01\"; .data),\r\n   p_setting:0,\r\n   q_setting:0,\r\n   pf_setting:0,\r\n   vpset_setting:0,\r\n   itemTimestamp: (now-(now%60)| floor),\r\n   timestamp: (now+28800|strftime(\"%Y-%m-%dT%H:%M:%SZ\"))\r\n  }";
                }
            }
            if (checkBox25.Checked == true)
            { // smartlogger
                if (checkBox21.Checked == true || checkBox22.Checked == true || checkBox24.Checked == true) strings[101] = ",\n";
                strings[102] = "  smartLogger:(.data.smartlogger | if .time >= 0 then (.time*0) else 1 end)";
            }
            strings[105] = "\n  },\n"; // etc end{}
            strings[106] = " SYSTIME: (now | floor | tostring)\n";
            strings[107] = " }\n}"; // detail end{} main emd{}
            strings[108] = "";

            // SolarEdge 特殊JQ   1.SF計算 2.E_today使用script做計算 (4.solarEdge、6.ABB) *temp Insulation event未修
            if (invJQType == 4) {
                strings[47] = ",\n pv_v: GenerateByIndexCrosserSF(\"inv_\"; .inverterIndexes; \"pv_v\"; .PVIndexes; \"pv_v_SF\"; .data)";
                strings[48] = ",\n pv_a: GenerateByIndexCrosserSF(\"inv_\"; .inverterIndexes; \"pv_a\"; .PVIndexes; \"pv_a_SF\"; .data)";
                strings[49] = ",\n pv_p: GenerateByIndexCrosserSF(\"inv_\"; .inverterIndexes; \"pv_p\"; .PVIndexes; \"pv_p_SF\"; .data)";
                strings[50] = ",\n PF: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"PF\"; \"PF_SF\"; .data)";
                strings[51] = ",\n Vrn: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"Vrs\"; \"voltage_SF\"; .data)";
                strings[52] = ",\n Vsn: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"Vst\"; \"voltage_SF\"; .data)";
                strings[53] = ",\n Vtn: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"Vrt\"; \"voltage_SF\"; .data)";
                strings[54] = ",\n Rc: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"Rc\"; \"current_SF\"; .data)";
                strings[55] = ",\n Sc: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"Sc\"; \"current_SF\"; .data)";
                strings[56] = ",\n Tc: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"Tc\"; \"current_SF\"; .data)";
                strings[57] = ",\n temp: GenerateByIndexCrosserSF(\"inv_\"; .inverterIndexes; \"temp_\"; .invTempIndexes; \"temp_SF\"; .data)";
                // strings[58] = ",\n State: GenerateA(\"inv_\"; .inverterIndexes; \"State\"; .data)";
                strings[59] = ",\n acp: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"acp\"; \"acp_SF\"; .data)";
                strings[60] = ",\n dcp: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"dcp\"; \"dcp_SF\"; .data)";
                strings[61] = ",\n Eff: GenerateADivideBCrosserSF(\"inv_\"; .inverterIndexes; \"acp\"; \"dcp\"; \"acp_SF\"; \"dcp_SF\"; .data)";
                strings[62] = ",\n AC_kWh: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"AC_kWh\"; \"AC_kWh_SF\"; .data)";
                strings[63] = ",\n freq: GenerateACrosserSF(\"inv_\"; .inverterIndexes; \"freq\"; \"freq_SF\"; .data)";
                strings[68] = ",\n E_today: GenerateArrayAMinusBCrosserSF(\"inv_\"; .inverterIndexes; \"AC_kWh\"; \"AC_kWh_base\"; \"AC_kWh_SF\"; .data)";
                strings[70] = ",\n power_r:GenerateACrosserBCrosserSF(\"inv_\"; .inverterIndexes; \"Vrs\"; \"Rc\"; \"voltage_SF\"; \"current_SF\"; .data)";
                strings[71] = ",\n power_s:GenerateACrosserBCrosserSF(\"inv_\"; .inverterIndexes; \"Vst\"; \"Sc\"; \"voltage_SF\"; \"current_SF\"; .data)";
                strings[72] = ",\n power_t:GenerateACrosserBCrosserSF(\"inv_\"; .inverterIndexes; \"Vrt\"; \"Tc\"; \"voltage_SF\"; \"current_SF\"; .data)";
            }
            if (invJQType == 4 || invJQType == 6) strings[68] = ",\n E_today: GenerateArrayAMinusB(\"inv_\"; .inverterIndexes; \"AC_kwh\"; \"AC_kwh_base\"; .data)";

            // PrimeVOLT 特殊JQ
            if (invJQType == 9) {
                strings[60] = ",\n dcp:GeneratePrimeVoltdcp(\"inv_\"; .inverterIndexes; \"dcp_1\"; \"dcp_2\"; \"dcp_3\"; \"dcp_4\"; .data)";
                strings[63] = ",\n freq:GeneratePrimeVoltfreq(\"inv_\"; .inverterIndexes; \"freq_1\"; \"freq_2\"; \"freq_3\"; .data)";
            }

            // 字串串接
            for (int i = 0; i < 120; i++) text += strings[i]; 

            // 寫檔到桌面
            using (StreamWriter writetext = new StreamWriter(filePath)) writetext.WriteLine(text); 
            label27.Visible = true;
            await Task.Delay(3000);  // 等待 3 秒
            label27.Visible = false;
        }

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox23.Checked)
            {
                textBox1.Enabled = true;
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                textBox5.Enabled = true;
                textBox6.Enabled = true;
                comboBox1.Enabled = true;
                comboBox8.Enabled = true;
                checkBox17.Enabled = true;
                checkBox18.Enabled = true;
                checkBox19.Enabled = true;
                checkBox20.Enabled = true;
                checkBox21.Enabled = true;
                checkBox22.Enabled = true;
                checkBox25.Enabled = true;

            }
            else
            {
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                comboBox1.Enabled = false;
                comboBox8.Enabled = false;
                checkBox17.Enabled = false;
                checkBox18.Enabled = false;
                checkBox19.Enabled = false;
                checkBox20.Enabled = false;
                checkBox21.Enabled = false;
                checkBox22.Enabled = false;
                checkBox25.Enabled = false;
            }
        }
        private void checkBox24_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox24.Checked)
            {
                comboBox9.Enabled = true;
                textBox7.Enabled = true;
            }
            else
            {
                comboBox9.Enabled = false;
                textBox7.Enabled = false;
            }
        }
        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e) {
            if (comboBox8.SelectedItem.ToString() == "無") {
                textBox5.Text = "0";
                textBox5.ReadOnly = true;
                textBox5.BackColor = Color.Gray;
            }
            else {
                textBox5.Text = "1";
                textBox5.ReadOnly = false;
                textBox5.BackColor = textBox4.BackColor;
            }
        }

        private void btnINVMode_Click(object sender, EventArgs e) {
            panelBrand.Top = 105;
            btnINVMode.BackColor = Color.PowderBlue;
            btnJQMode.BackColor = Color.Transparent;
            btnINVList.BackColor = Color.Transparent;
            btnINVPRO.BackColor = Color.Transparent;
            groupBox1.Visible = false;
            groupBox2.Visible = true;
            //groupBox3.Visible = false;
            //groupBox4.Visible = false;
        }

        private void btnJQMode_Click(object sender, EventArgs e) {
            panelBrand.Top = 145;
            btnINVMode.BackColor = Color.Transparent;
            btnJQMode.BackColor = Color.PowderBlue;
            btnINVList.BackColor = Color.Transparent;
            btnINVPRO.BackColor = Color.Transparent;
            groupBox2.Visible = false;
            groupBox1.Visible = true;
            //groupBox3.Visible = false;
            //groupBox4.Visible = false;
        }
        private void btnINVList_Click(object sender, EventArgs e) {
            panelBrand.Top = 185;
            btnINVMode.BackColor = Color.Transparent;
            btnJQMode.BackColor = Color.Transparent;
            btnINVList.BackColor = Color.PowderBlue;
            btnINVPRO.BackColor = Color.Transparent;
            groupBox2.Visible = false;
            groupBox1.Visible = false;
            //groupBox3.Visible = false;
            //groupBox4.Visible = false;
        }
        private void btnINVPRO_Click(object sender, EventArgs e) {
            panelBrand.Top = 225;
            btnINVMode.BackColor = Color.Transparent;
            btnJQMode.BackColor = Color.Transparent;
            btnINVList.BackColor = Color.Transparent;
            btnINVPRO.BackColor = Color.PowderBlue;
            groupBox2.Visible = false;
            groupBox1.Visible = false;
        }
        
        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(1);
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(2);
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(3);
        }
        private void checkBox4_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(4);
        }
        private void checkBox5_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(5);
        }
        private void checkBox6_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(6);
        }
        private void checkBox7_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(7);
        }
        private void checkBox8_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(8);
        }
        private void checkBox9_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(9);
        }
        private void checkBox10_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(10);
        }
        private void checkBox11_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(11);
        }
        private void checkBox12_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(12);
        }
        private void checkBox13_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(13);
        }
        private void checkBox14_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(14);
        }
        private void checkBox15_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(15);
        }
        private void checkBox16_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes(16);
        }
        // 選取一個checkbox後，則無法選擇其他checkbox
        private void DisableOtherCheckboxes(int index) { 
            CheckBox selectedCheckbox = (CheckBox)sender;
            foreach (CheckBox checkbox in allCheckboxes) {
                if (checkbox != selectedCheckbox) checkbox.Enabled = false;
                else checkbox.Enabled = true;
            }
            invModelkExport = index;
        }
        // 恢復所有checkbox狀態
        private void button3_Click(object sender, EventArgs e) { 
            UncheckAllCheckBoxes();
        }
        private void UncheckAllCheckBoxes() {
            foreach (CheckBox checkBox in allCheckboxes)
            {
                checkBox.Checked = false;
                checkBox.Enabled = true;
            }
            foreach (CheckBox checkBox in allCheckboxes)
            {
                checkBox.Checked = false;
                checkBox.Enabled = true;
            }
        }
        // 產生INV 模板
        private void button2_Click(object sender, EventArgs e) {
            ExportExcel(invModelkExport);
        }

        private void ExportExcel(int invModelkExport) {
            string str = "";
            if (invModelkExport == 1) str = "台達電DELTA-RPI.csv"; 
            else if (invModelkExport == 2) str = "華為HUAWEI-KTL.csv";
            else if (invModelkExport == 3) str = "施奈德-TL20000E.csv";
            else if (invModelkExport == 4) str = "陽光電源SUNGROW-SG110CX.csv";
            else if (invModelkExport == 5) str = "亞力ALLIS-PLUS-20K.csv";
            else if (invModelkExport == 6) str = "亞力ALLIS-TOUGH-20K.csv";
            else if (invModelkExport == 7) str = "SolarEdge-SE.csv";
            else if (invModelkExport == 8) str = "PrimeVolt-PV-60000T.csv";
            else if (invModelkExport == 9) str = "KACO";
            else if (invModelkExport == 10) str = "ABB";
            else if (invModelkExport == 11) str = "華為HUAWEI-KTL(舊版本).csv";
            else str = "TEST.csv";

            // 讀取路徑: 桌面-file - TEXT.csv
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePathInput = Path.Combine(desktopPath, "ACME_Builder\\file2\\", str);

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Open(filePathInput);
            var worksheet = workbook.Sheets[1];
            var range = worksheet.UsedRange;
            // 輸出路徑 
            var filePath = Path.Combine(desktopPath, str);
            using (var writer = new StreamWriter(filePath)) {
                for (int row = 1; row <= range.Rows.Count; row++) {
                    for (int col = 1; col <= range.Columns.Count; col++) {
                        var cell = range.Cells[row, col];
                        var value = cell.Value2;
                        writer.Write(value + ",");
                    }
                    writer.WriteLine();
                }
            }
            // 關閉 Excel 應用程式
            excelApp.Quit();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    } // class
}
