using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//list:
// 一定要上傳pv_p? Insulation? 
//1. [eff]要acp/dcp?
//2. [acp/dcp]要加總?
//3. [etoday]有提供?
//4. event

//2.1. 增加匯出設備模板 版本

//3.1 電表設備模板


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
        // int INVTEMP_QUANTITY = 1; // INV_TEMP 數量

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
            this.Size = new Size(573, this.Size.Width);
            this.Size = new Size(this.Size.Width, 360);
        }
        private async void button1_Click(object sender, EventArgs e) {
            // 讀取參數
            INV_QUANTITY = int.Parse(textBox3.Text);
            PV_QUANTITY = int.Parse(textBox4.Text);
            IRR_QUANTITY = int.Parse(textBox5.Text);
            CAPACITY = textBox6.Text;
            FACTORYNAME = textBox1.Text;
            FACTORYID = textBox2.Text;

            string text = "";
            string[] invLabel = new string[50];  // INV 告警事件表
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "JQ_Format.txt");
            string[] strings = new string[100];
            string[] indexes = new string[INV_QUANTITY];
            string[] indexes2 = new string[PV_QUANTITY];
            string[] indexes3 = new string[IRR_QUANTITY];


            // INV 告警表
            // DeltaEventIndexes
            invLabel[1] = "AlarmIndexes:[\"alarm_E1\",\"alarm_E2\",\"alarm_E3\",\"alarm_W1\",\"alarm_W2\",\"alarm_F1\",\"alarm_F2\",\"alarm_F3\",\"alarm_F4\",\"alarm_F5\"]";
            // SolarEventIndexes
            invLabel[2] = "AlarmIndexes:[\"event_Gb\",\"event_M1\",\"event_M2\",\"event_M3\"]";
            // HUAWEISUN2000EventIndexes
            invLabel[3] = "AlarmIndexes:[\"event_1\",\"event_2\",\"event_3\"]";
            // HUAWEI36KTLIndexes
            invLabel[4] = "AlarmIndexes:[\"event_1\",\"event_2\",\"event_3\",\"event_4\",\"event_5\",\"event_6\",\"event_7\",\"event_8\",\"event_9\",\"event_10\",\"event_11\"]";
            // SUNGROW_SG110CX
            invLabel[5] = "AlarmIndexes:[\"event\"]";
            // SCHNEIDERTL20000E
            invLabel[6] = "AlarmIndexes:[\"event_1\",\"event_2\",\"event_3\",\"event_4\"]";
            // ABBPVS100_120
            invLabel[7] = "AlarmIndexes:[\"event_1\",\"eventVendor_1\",\"eventVendor_2\",\"eventVendor_3\"]";
            // -
            invLabel[8] = "AlarmIndexes:[]";

            switch (comboBox1.Text) {
                case "DeltaEvent":
                    invLabel[0] = invLabel[1];
                    break;
                case "HUAWEI36KTL":
                    invLabel[0] = invLabel[4];
                    break;
                case "HUAWEISUN2000":
                    invLabel[0] = invLabel[3];
                    break;
                case "Solar":
                    invLabel[0] = invLabel[2];
                    break;
                case "SUNGROW_SG110CX":
                    invLabel[0] = invLabel[5];
                    break;
                case "ABBPVS100_120":
                    invLabel[0] = invLabel[7];
                    break;
                case "SCHNEIDERTL20000E":
                    invLabel[0] = invLabel[6];
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
            strings[15] = "def GenerateOnlyOneTempIndex($source_prefix; $source_indexes; $tag_prefix; $tag_indexes; $data):$source_indexes | map(GeneratetempSourceByIndex($source_prefix+(.|tostring); $tag_prefix; $tag_indexes; $data));\n";
            strings[16] = "def GenerateABCrossor($source; $tag_prefix_a; $tag_prefix_b; $tag_indexes; $data):(($tag_indexes | map( if $data.\"\\($source)\".status == 0 then (if ($data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\") != null then $data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\" * $data.\"\\($source)\".\"\\($tag_prefix_b + (.tostring))\"| FormatFloat else \" - 1\" end )else 0 end)) | map(select(. != \" - 1\")));\n";
            strings[17] = "def GenerateCalculateABCrossorByIndex($source_prefix; $source_indexes; $tag_a; $tag_b;$tag_indexes; $data):$source_indexes | map(GenerateABCrossor($source_prefix+(.|tostring); $tag_a; $tag_b; $tag_indexes; $data));\n";
            strings[18] = "def GenerateABCrossoranddivide1000($source; $tag_prefix_a; $tag_prefix_b; $tag_indexes; $data):(($tag_indexes | map( if $data.\"\\($source)\".status == 0 then ( if ($data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\") and ($data.\"\\($source)\".\"\\($tag_prefix_b + (.| tostring))\") != null then $data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\" * $data.\"\\($source)\".\"\\($tag_prefix_b + (.tostring))\" / 1000 | FormatFloat else($data.\"\\($source)\".\"\\($tag_prefix_a + (.| tostring))\")end )else 0 end)));\n";
            strings[19] = "def GenerateCalculateABCrossorByIndexanddivide1000($source_prefix; $source_indexes; $tag_a;$tag_b; $tag_indexes; $data):$source_indexes | map(GenerateABCrossoranddivide1000($source_prefix+(.|tostring); $tag_a; $tag_b; $tag_indexes;$data));\n";
            strings[20] = "def GenerateAMinusB($source_prefix; $source_indexes; $a; $b; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then(if (( $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix + (.tostring))\".\"\\($b)\") != 0) then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" - $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\"| FormatFloat else 0 end)else 0 end);\n";
            strings[21] = "def GenerateADivideB($source_prefix; $source_indexes; $a; $b; $data):$source_indexes | map(if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then (if (( $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) != 0 and ( $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\") != 0) then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" / $data.\"\\($source_prefix + (.| tostring))\".\"\\($b)\" | FormatFloat else 0 end)else 0 end);\n";
            strings[22] = "def GenerateInsulationA($source_prefix; $source_indexes; $a; $data): $source_indexes | map( if $data.\"\\($source_prefix + (.| tostring))\".status == 0 then( if ($data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" ) > 0 then $data.\"\\($source_prefix + (.| tostring))\".\"\\($a)\" | IRRFormatFloat else 0 end )else 0 end);\n";
            strings[23] = "def GenerateEvent($source_prefix; $source_indexes; $tag_indexes; $data):$source_indexes | map(GenerateEventB($source_prefix+(.|tostring); $tag_indexes; $data));\n";
            strings[24] = "def GenerateEventB($source; $tag_indexes; $data):(($tag_indexes | map ( if $data.\"\\($source)\".status == 0 then ( if ($data.\"\\($source)\".\"\\((.+))\") != null then $data.\"\\($source)\".\"\\((.+))\" | FormatFloat else ($data.\"\\($source)\".\"\\((.+))\") end )else 0 end)));\n";
            strings[25] = "def GenerateEventT($source_prefix; $source_indexes; $tag_indexes; $data):$source_indexes | map(EventOneArray($source_prefix+(.|tostring); $tag_indexes; $data));";

            // []index定義 
            strings[33] = "\n\n{\ninverterIndexes: [\"" + string.Join("\",\"", indexes) + "\"],\n";
            strings[34] = "PVIndexes: [\"" + string.Join("\",\"", indexes2) + "\"],\n";
            strings[35] = "iRRIndexes: [\"" + string.Join("\",\"", indexes3) + "\"],\n";
            strings[36] = "invTempIndexes: [1],\n"; // inv temp 數量固定為1
            strings[37] = "capacityIndexes: " + CAPACITY + ",\n";
            strings[38] = invLabel[0];
            strings[39] = ",\n";
            strings[40] = "data: .\n} |\n\n";

            // main payload
            strings[42] = "{\nfactoryName: \"" + FACTORYNAME + "\",\n"; // main{}
            strings[43] = "timestamp: (now+28800|strftime(\"%Y-%m-%dT%H:%M:%SZ\")),\n";
            strings[44] = "detail:\n{\n"; // detail{}
            strings[45] = "factoryId: \"" + FACTORYID + "\"";
            strings[46] = ",\nfactoryName: \"" + FACTORYNAME + "\"";
            strings[47] = ",\npv_v: GenerateByIndex(\"inv_\"; .inverterIndexes; \"pv_v\"; .PVIndexes; .data)";
            strings[48] = ",\npv_a: GenerateByIndex(\"inv_\"; .inverterIndexes; \"pv_a\"; .PVIndexes; .data)";
            strings[49] = ",\npv_p: GenerateByIndex(\"inv_\"; .inverterIndexes; \"pv_p\"; .PVIndexes; .data)";
            strings[50] = ",\nPF: GenerateA(\"inv_\"; .inverterIndexes; \"PF\"; .data)";
            strings[51] = ",\nVrn: GenerateA(\"inv_\"; .inverterIndexes; \"Vrs\"; .data)";
            strings[52] = ",\nVsn: GenerateA(\"inv_\"; .inverterIndexes; \"Vst\"; .data)";
            strings[53] = ",\nVtn: GenerateA(\"inv_\"; .inverterIndexes; \"Vrt\"; .data)";
            strings[54] = ",\nRc: GenerateA(\"inv_\"; .inverterIndexes; \"Rc\"; .data)";
            strings[55] = ",\nSc: GenerateA(\"inv_\"; .inverterIndexes; \"Sc\"; .data)";
            strings[56] = ",\nTc: GenerateA(\"inv_\"; .inverterIndexes; \"Tc\"; .data)";
            strings[57] = ",\ntemp: GenerateOnlyOneTempIndex(\"inv_\"; .inverterIndexes; \"temp\"; .invTempIndexes; .data)";
            strings[58] = ",\nState: GenerateA(\"inv_\"; .inverterIndexes; \"State\"; .data)";
            strings[59] = ",\nacp: GenerateA(\"inv_\"; .inverterIndexes; \"acp\"; .data)";
            strings[60] = ",\ndcp: GenerateA(\"inv_\"; .inverterIndexes; \"dcp\"; .data)";
            strings[61] = ",\nEff: GenerateADivideB(\"inv_\"; .inverterIndexes; \"acp\"; \"dcp\"; .data)";
            strings[62] = ",\nAC_kWh: GenerateA(\"inv_\"; .inverterIndexes; \"AC_kwh\"; .data)";
            strings[63] = ",\nfreq: GenerateA(\"inv_\"; .inverterIndexes; \"freq\"; .data)";
            strings[64] = ",\nevent: GenerateEvent(\"inv_\"; .inverterIndexes; .AlarmIndexes; .data)";
            strings[65] = ",\nevent: GenerateEventT(\"inv_\"; .inverterIndexes; .AlarmIndexes; .data)";

            strings[68] = ",\nE_today: ";
            strings[69] = ",\ncapacity: (.capacityIndexes)";
            if (comboBox8.Text == "無") strings[70] = ",\nIRR: [null]";
            else strings[70] = ",\nIRR: GenerateIRRArray(\"IRR_\"; .iRRIndexes; \"IRR\"; .data)";
            strings[71] = ",\nPVTemp: Generate(\"PV_TEMP\"; \"temp\"; .data)";
            strings[72] = ",\nENVTemp: Generate(\"ENV_TEMP\"; \"temp\"; .data)";
            // 風速計
            if (comboBox4.Text == "有") strings[73] = ",\nAnemometer: Generate(\"Anemometer\"; \"Wind_speed\"; .data)";
            // 水位計
            if (comboBox5.Text == "有") strings[74] = ",\nWaterlevel: Generate(\"Waterlevel\"; \"Measurement_output_value\"; .data)";

            strings[75] = ",\nErrorMessage:\n{"; // Errormessage{}
            strings[76] = "\ninv: GenerateINVStatus(\"inv_\"; .inverterIndexes; .data)";
            if (comboBox8.Text == "無") strings[77] = ",\nIRR: [null]";
            else strings[77] = ",\nIRR: GenerateINVStatus(\"IRR_\"; .iRRIndexes; .data)";

            strings[78] = ",\nPVTemp: GenerateStatus(\"PV_TEMP\";.data)";
            strings[79] = ",\nENVtemp: GenerateStatus(\"ENV_TEMP\";.data)";
            if (comboBox4.Text == "有")
            { // 風速計 (ErrorMessage)
                strings[80] = ",\nAnemometer: GenerateStatus(\"Anemometer\";.data)";
            }
            if (comboBox5.Text == "有")
            { // 水位計 (ErrorMessage)
                strings[81] = ",\nWaterlevel: GenerateStatus(\"Waterlevel\";.data)";
            }
            if (comboBox6.Text == "有")
            { // 低壓電錶 (ErrorMessage)
                strings[82] = ",\nLV_Meter: GenerateStatus(\"LV_meter\";.data)";
            }
            if (comboBox7.Text == "有")
            { // 高壓電表 (ErrorMessage)
                strings[83] = ",\nHV_Meter: GenerateStatus(\"HV_meter\";.data)";
            }
            strings[84] = "\n}"; // Errormessage end{}

            strings[85] = ",\netc:\n{"; // etc{}
            // 低壓電錶(value)
            if (comboBox6.Text == "有") strings[86] = "\n\"LV-meter\": {\nVln_a: Generate(\"LV-meter\"; \"Vln_a\"; .data),\nVln_b: Generate(\"LV-meter\"; \"Vln_b\"; .data),\nVln_c: Generate(\"LV-meter\"; \"Vln_c\"; .data),\nVll_ab: Generate(\"LV-meter\"; \"Vll_ab\"; .data),\nVll_bc: Generate(\"LV-meter\"; \"Vll_bc\"; .data),\nVll_ca: Generate(\"LV-meter\"; \"Vll_ca\"; .data),\nI_a: Generate(\"LV-meter\"; \"I_a\"; .data),\nI_b: Generate(\"LV-meter\"; \"I_b\"; .data),\nI_c: Generate(\"LV-meter\"; \"I_c\"; .data),\nFreq: Generate(\"LV-meter\"; \"freq\"; .data),\nP: Generate(\"LV-meter\"; \"P\"; .data),\nKVAR_tot: Generate(\"LV-meter\"; \"Q\"; .data),\nKVA_tot: Generate(\"LV-meter\"; \"S\"; .data)\n}";
            // 高壓電錶(value)
            if (comboBox7.Text == "有")
            {
                if (comboBox6.Text == "有") strings[87] = ",";
                strings[88] = "\n\"HV-meter\": {\nVln_a: Generate(\"HV_meter\"; \"Vln_a\"; .data),\nVln_b: Generate(\"HV_meter\"; \"Vln_b\"; .data),\nVln_c: Generate(\"HV_meter\"; \"Vln_c\"; .data),\nVll_ab: Generate(\"HV_meter\"; \"Vll_ab\"; .data),\nVll_bc: Generate(\"HV_meter\"; \"Vll_bc\"; .data),\nVll_ca: Generate(\"HV_meter\"; \"Vll_ca\"; .data),\nI_a: Generate(\"HV_meter\"; \"I_a\"; .data),\nI_b: Generate(\"HV_meter\"; \"I_b\"; .data),\nI_c: Generate(\"HV_meter\"; \"I_c\"; .data),\nFreq: Generate(\"HV_meter\"; \"freq\"; .data),\nP: Generate(\"HV_meter\"; \"P\"; .data),\nKVAR_tot: Generate(\"HV_meter\"; \"Q\"; .data),\nKVA_tot: Generate(\"HV_meter\"; \"S\"; .data)\n}";
            }
            strings[95] = "},\n"; // etc end{}
            strings[96] = "SYSTIME: (now | floor | tostring)\n";
            strings[97] = "}\n}"; // detail end{} main emd{}
            strings[98] = "";

            for (int i = 0; i < 100; i++) text += strings[i]; // 字串串接

            using (StreamWriter writetext = new StreamWriter(filePath)) writetext.WriteLine(text); // 寫檔到桌面
            label27.Visible = true;
            await Task.Delay(3000);  // 等待 3 秒
            label27.Visible = false;
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
            panelBrand.Top = 93;
            btnINVMode.BackColor = Color.PowderBlue;
            btnJQMode.BackColor = Color.Transparent;
            btnINVList.BackColor = Color.Transparent;
            btnINVPRO.BackColor = Color.Transparent;
            /*
            groupBox2.Location = new Point(118, groupBox2.Location.Y);
            groupBox2.Location = new Point(groupBox2.Location.X, 12);*/
            groupBox1.Visible = false;
            groupBox2.Visible = true;
            //groupBox3.Visible = false;
            //groupBox4.Visible = false;
        }

        private void btnJQMode_Click(object sender, EventArgs e) {
            panelBrand.Top = 133;
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
            panelBrand.Top = 173;
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
            panelBrand.Top = 213;
            btnINVMode.BackColor = Color.Transparent;
            btnJQMode.BackColor = Color.Transparent;
            btnINVList.BackColor = Color.Transparent;
            btnINVPRO.BackColor = Color.PowderBlue;
            groupBox2.Visible = false;
            groupBox1.Visible = false;
        }
        
        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox4_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox5_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox6_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox7_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox8_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox9_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox10_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox11_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox12_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox13_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox14_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox15_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void checkBox16_CheckedChanged(object sender, EventArgs e) {
            DisableOtherCheckboxes();
        }
        private void DisableOtherCheckboxes() { // 選取一個checkbox後，則無法選擇其他checkbox
            CheckBox selectedCheckbox = (CheckBox)sender;
            foreach (CheckBox checkbox in allCheckboxes) {
                if (checkbox != selectedCheckbox) checkbox.Enabled = false;
                else checkbox.Enabled = true;
            }
        }
        private void button3_Click(object sender, EventArgs e) { // 恢復所有checkbox狀態
            UncheckAllCheckBoxes();
        }
        private void UncheckAllCheckBoxes() {
            foreach (CheckBox checkBox in allCheckboxes) {
                checkBox.Checked = false;
                checkBox.Enabled = true;
            }
            foreach (CheckBox checkBox in allCheckboxes) {
                checkBox.Checked = false;
                checkBox.Enabled = true;
            }
        }

    } // class
}
