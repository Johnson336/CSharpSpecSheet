using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace CSharpSpecSheet
{
    public partial class SpecSheet : Form
    {
        [DllImport("User32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);

        SQLiteConnection m_dbConnection;

        public SpecSheet()
        {
            InitializeComponent();
            setDate();

            initializeDatabase();
            initializeDropdowns();

        }

        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }

        private void initializeDatabase()
        {   
            if (!File.Exists("db.sqlite"))
            {
                SQLiteConnection.CreateFile("db.sqlite");
            }
            
            m_dbConnection = new SQLiteConnection("Data Source=db.sqlite;Version=3;");
            m_dbConnection.Open();

            // create table to hold dropdown data
            //executeSQL("DROP TABLE IF EXISTS dropdown_data");
            executeSQL("CREATE TABLE IF NOT EXISTS dropdown_data (category TEXT NOT NULL, name TEXT UNIQUE NOT NULL, frequency INT NOT NULL, id INT PRIMARY KEY)");

            //executeSQL("DROP TABLE IF EXISTS archive");
            executeSQL("CREATE TABLE IF NOT EXISTS archive ("+
                        "id INT PRIMARY KEY,"+
                        "timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,"+
                        "ispf TEXT NOT NULL,"+
                        "date TEXT NOT NULL,"+
                        "condition TEXT NOT NULL,"+
                        "brand TEXT NOT NULL,"+
                        "serial TEXT UNIQUE NOT NULL,"+
                        "model TEXT NOT NULL,"+
                        "formfactor TEXT NOT NULL,"+
                        "cpuqty TEXT NOT NULL,"+
                        "cpucores TEXT NOT NULL,"+
                        "checkht TEXT NOT NULL,"+
                        "cpuspeed TEXT NOT NULL,"+
                        "cputype TEXT NOT NULL,"+
                        "busspeed TEXT NOT NULL,"+
                        "cpuname TEXT NOT NULL,"+
                        "memorysize TEXT NOT NULL,"+
                        "memoryrating TEXT NOT NULL,"+
                        "memorytype TEXT NOT NULL,"+
                        "memoryspeed TEXT NOT NULL,"+
                        "weight TEXT NOT NULL,"+
                        "hddqty TEXT NOT NULL,"+
                        "hddsize TEXT,"+
                        "hddtype TEXT,"+
                        "hddrpm TEXT,"+
                        "hddserial TEXT,"+
                        "video TEXT NOT NULL,"+
                        "videomodel TEXT NOT NULL,"+
                        "vram TEXT NOT NULL,"+
                        "optical TEXT NOT NULL,"+
                        "drivesnone TEXT NOT NULL,"+
                        "drivesfdd TEXT NOT NULL,"+
                        "drivestape TEXT NOT NULL,"+
                        "lcdsize TEXT NOT NULL,"+
                        "networknone TEXT NOT NULL,"+
                        "ethernet TEXT NOT NULL,"+
                        "modem TEXT NOT NULL,"+
                        "wifi TEXT NOT NULL,"+
                        "bt TEXT NOT NULL,"+
                        "coa TEXT NOT NULL,"+
                        "osno TEXT NOT NULL,"+
                        "osyes TEXT NOT NULL,"+
                        "notes TEXT NOT NULL,"+
                        "accnone TEXT NOT NULL,"+
                        "accac TEXT NOT NULL,"+
                        "accpower TEXT NOT NULL,"+
                        "accbatt TEXT NOT NULL,"+
                        "accextbatt TEXT NOT NULL,"+
                        "accfinger TEXT NOT NULL,"+
                        "accwebcam TEXT NOT NULL,"+
                        "acckeyboard TEXT NOT NULL,"+
                        "accmouse TEXT NOT NULL,"+
                        "damage TEXT NOT NULL,"+
                        "usb TEXT NOT NULL,"+
                        "numethernet TEXT NOT NULL,"+
                        "nummodem TEXT NOT NULL,"+
                        "vga TEXT NOT NULL,"+
                        "dvi TEXT NOT NULL,"+
                        "svideo TEXT NOT NULL,"+
                        "ps2 TEXT NOT NULL,"+
                        "audio TEXT NOT NULL,"+
                        "esatap TEXT NOT NULL,"+
                        "numserial TEXT NOT NULL,"+
                        "parallel TEXT NOT NULL,"+
                        "pcmcia TEXT NOT NULL,"+
                        "sdcard TEXT NOT NULL,"+
                        "firewire TEXT NOT NULL,"+
                        "esata TEXT NOT NULL,"+
                        "hdmi TEXT NOT NULL,"+
                        "scsi TEXT NOT NULL,"+
                        "displayport TEXT NOT NULL,"+
                        "version TEXT NOT NULL,"+
                        "tester TEXT NOT NULL,"+
                        "caddyqty TEXT NOT NULL,"+
                        "caddyna TEXT NOT NULL)"
                        );

            //executeSQL("DROP TABLE IF EXISTS cpu_data");
            executeSQL("CREATE TABLE IF NOT EXISTS cpu_data (cpuseries TEXT NOT NULL, cputype TEXT NOT NULL, busspeed TEXT NOT NULL, cpuspeed TEXT NOT NULL, cpucores TEXT NOT NULL, cpuht TEXT NOT NULL, model TEXT NOT NULL, formfactor TEXT NOT NULL, id INT PRIMARY KEY, UNIQUE(cpuseries, model, formfactor))");


        }

        private void initializeDropdowns()
        {
            var c = GetAll(this, typeof(ComboBox));
            //initDropdown("formfactor", dropFormfactor);
            //initDropdown("optical", dropOptical);
            foreach (var item in c)
            {
                initDropdown(((ComboBox)item));
            }
        }

        private void initDropdown(ComboBox drop)
        {
            SQLiteDataReader reader = executeSQLReader("SELECT * FROM dropdown_data WHERE category='" + drop.Name + "' ORDER BY frequency DESC");
            drop.Items.Clear();
            while (reader.Read())
            {
                drop.Items.Add(reader["name"]);
            }

        }

        private void executeSQL(string stmt)
        {
            SQLiteCommand cmd = new SQLiteCommand(stmt, m_dbConnection);

            //lblStatusBar2.Text = lblStatusBar.Text;
            //lblStatusBar.Text = stmt + ": " + cmd.ExecuteNonQuery() + " rows affected.";
            cmd.ExecuteNonQuery();
        }

        private SQLiteDataReader executeSQLReader(string stmt)
        {
            SQLiteCommand command = new SQLiteCommand(stmt, m_dbConnection);
            return command.ExecuteReader();
        }

        private void saveDropdown(ComboBox drop)
        {
            if ((drop.Name == "") || (drop.Text == "")) {
                return;
            }

            //executeSQL("UPDATE OR IGNORE dropdown_data SET category='" + drop.Name + "', name='" + drop.Text + "', frequency = frequency + 1 WHERE category='" + drop.Name + "' AND name='" + drop.Text + "'");
            executeSQL("UPDATE OR IGNORE dropdown_data SET frequency = frequency + 1 WHERE category='" + drop.Name + "' AND name='" + drop.Text + "'");

            executeSQL("INSERT OR IGNORE INTO dropdown_data (category, name, frequency) VALUES ('" + drop.Name + "', '" + drop.Text + "', 1)");


            initDropdown(drop);
        }

        private void saveCPUInfo()
        {
            //executeSQL("UPDATE OR IGNORE cpu_data SET cpuseries='" + dropCPUName.Text + "', cputype='" + dropCPUType.Text + "', busspeed='" + txtBusSpeed.Text + "', cpuspeed='" + double.Parse(txtCPUSpeed.Text) + ", cpucores='" + spinCPUCores.Value + "' WHERE cpuseries='" + dropCPUName.Text + "'");
            executeSQL("UPDATE OR IGNORE cpu_data SET cputype='" + dropCPUType.Text + "', busspeed='" + txtBusSpeed.Text + "', cpuspeed='" + Math.Round(double.Parse(txtCPUSpeed.Text), 2) + "', cpucores='" + spinCPUCores.Value + "', cpuht='" + checkHT.Checked + "' WHERE cpuseries='" + dropCPUName.Text + "' AND model='" + txtModel.Text + "' AND formfactor='" + dropFormfactor.Text + "'");
            executeSQL("INSERT OR IGNORE INTO cpu_data (cpuseries, cputype, busspeed, cpuspeed, cpucores, cpuht, model, formfactor) VALUES('" + dropCPUName.Text + "', '" + dropCPUType.Text + "', '" + txtBusSpeed.Text + "', '" + Math.Round(double.Parse(txtCPUSpeed.Text), 2) + "', '" + spinCPUCores.Value + "', '" + checkHT.Checked + "', '" + txtModel.Text + "', '" + dropFormfactor.Text + "')");
        }

        private void setDate()
        {
            /* Set labelDate to current date upon form init */
            labelDate.Text = DateTime.Now.ToString("MM/dd/yy");
        }

        private void checkDrivesNone_CheckedChanged(object sender, EventArgs e)
        {
            if (checkDrivesNone.Checked == true) {
                checkDrivesFDD.Checked = false;
                checkDrivesTape.Checked = false;
            } else if ((checkDrivesFDD.Checked == false) && (checkDrivesTape.Checked == false))
            {
                checkDrivesNone.Checked = true;
            }
        }

        private void checkDrivesFDD_CheckedChanged(object sender, EventArgs e)
        {
            checkDrives(checkDrivesFDD, e);
        }

        private void checkDrivesTape_CheckedChanged(object sender, EventArgs e)
        {
            checkDrives(checkDrivesTape, e);
        }

        private void checkDrives(CheckBox s, EventArgs e)
        {
            if (s.Checked == true)
            {
                checkDrivesNone.Checked = false;
            }
            else if ((checkDrivesTape.Checked == false) && (checkDrivesFDD.Checked == false))
            {
                checkDrivesNone.Checked = true;
            }
        }


        private void spinHDDQty_ValueChanged(object sender, EventArgs e)
        {
            if (spinHDDQty.Value > 0)
            {
                txtHDDSize.Text = "";
                txtHDDSize.Enabled = true;
                dropHDDType.Enabled = true;
                dropHDDRPM.Enabled = true;
                txtHDDSerial.Text = "";
                txtHDDSerial.Enabled = true;
            } else
            {
                txtHDDSize.Text = "N/A";
                txtHDDSize.Enabled = false;
                dropHDDType.Text = "";
                dropHDDType.Enabled = false;
                dropHDDRPM.Text = "";
                dropHDDRPM.Enabled = false;
                txtHDDSerial.Text = "N/A";
                txtHDDSerial.Enabled = false;
            }

        }

        private void checkNetworkNone_CheckedChanged(object sender, EventArgs e)
        {
            if (checkNetworkNone.Checked == true)
            {
                checkEthernet.Checked = false;
                checkModem.Checked = false;
                checkWiFi.Checked = false;
                checkBT.Checked = false;
            } else if ((checkEthernet.Checked == false) && (checkModem.Checked == false) && (checkWiFi.Checked == false) && (checkBT.Checked == false))
            {
                checkNetworkNone.Checked = true;
            }
        }

        private void checkEthernet_CheckedChanged(object sender, EventArgs e)
        {
            checkNetwork(checkEthernet, e);
        }

        private void checkModem_CheckedChanged(object sender, EventArgs e)
        {
            checkNetwork(checkModem, e);
        }

        private void checkWiFi_CheckedChanged(object sender, EventArgs e)
        {
            checkNetwork(checkWiFi, e);
        }

        private void checkBT_CheckedChanged(object sender, EventArgs e)
        {
            checkNetwork(checkBT, e);
        }

        private void checkNetwork(CheckBox s, EventArgs e)
        {
            if (s.Checked == true)
            {
                checkNetworkNone.Checked = false;
            }
            else if ((checkEthernet.Checked == false) && (checkModem.Checked == false) && (checkWiFi.Checked == false) && (checkBT.Checked == false))
            {
                checkNetworkNone.Checked = true;
            }
        }

        private void checkAccNone_CheckedChanged(object sender, EventArgs e)
        {
            if (checkAccNone.Checked == true)
            {
                checkAccAC.Checked = false;
                checkAccPower.Checked = false;
                checkAccBatt.Checked = false;
                checkAccExtBatt.Checked = false;
                checkAccFinger.Checked = false;
                checkAccWebcam.Checked = false;
                checkAccKeyboard.Checked = false;
                checkAccMouse.Checked = false;
            } else if ((checkAccAC.Checked == false) && (checkAccPower.Checked == false) && (checkAccBatt.Checked == false) && (checkAccExtBatt.Checked == false) && (checkAccFinger.Checked == false) && (checkAccWebcam.Checked == false) && (checkAccKeyboard.Checked == false) && ( checkAccMouse.Checked == false))
            {
                checkAccNone.Checked = true;
            }
        }

        private void checkAccAC_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccAC, e);
        }

        private void checkAccPower_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccPower, e);
        }

        private void checkAccBatt_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccBatt, e);
        }

        private void checkAccExtBatt_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccExtBatt, e);
        }

        private void checkAccFinger_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccFinger, e);
        }

        private void checkAccWebcam_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccWebcam, e);
        }

        private void checkAccKeyboard_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccKeyboard, e);
        }

        private void checkAccMouse_CheckedChanged(object sender, EventArgs e)
        {
            checkAccessories(checkAccMouse, e);
        }

        private void checkAccessories(CheckBox s, EventArgs e)
        {
            if (s.Checked == true)
            {
                checkAccNone.Checked = false;
            }
            else if ((checkAccAC.Checked == false) && (checkAccPower.Checked == false) && (checkAccBatt.Checked == false) && (checkAccExtBatt.Checked == false) && (checkAccFinger.Checked == false) && (checkAccWebcam.Checked == false) && (checkAccKeyboard.Checked == false) && (checkAccMouse.Checked == false))
            {
                checkAccNone.Checked = true;
            }
        }

        private string SerializeObject(ComboBox toSerialize)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(toSerialize.GetType());

            using (StringWriter textWriter = new StringWriter())
            {
                xmlSerializer.Serialize(textWriter, toSerialize);
                return textWriter.ToString();
            }
        }

        private void printButton_Click(object sender, EventArgs e)
        {
            //performPrint();


            archiveSerial();
            saveCPUInfo();

        }

        private void generateLabel()
        {

            string directory = AppDomain.CurrentDomain.BaseDirectory + "data";
            System.IO.Directory.CreateDirectory(directory); // create archive directory if it doesn't already exist 
            string filename = directory + "\\SpecSheetData.csv";

            string output = txtISPF.Text + ", " + labelDate.Text + ", " + dropCondition.Text + ", " + dropBrand.Text + ", " + txtSerial.Text + ", " +
                txtModel.Text + ", " + dropFormfactor.Text + ", " + spinCPUQty.Value + ", " + spinCPUCores.Value + ", " + checkHT.Checked + ", " + txtCPUSpeed.Text + ", " +
                dropCPUType.Text + ", " + txtBusSpeed.Text + ", " + dropCPUName.Text + ", " + dropMemorySize.Text + ", " + dropMemoryRating.Text + ", " + dropMemoryType.Text + ", " +
                dropMemorySpeed.Text + ", " + txtWeight.Text + ", " + spinHDDQty.Value + ", " + txtHDDSize.Text + ", " + dropHDDType.Text + ", " + dropHDDRPM.Text + ", " +
                txtHDDSerial.Text + ", " + dropVideo.Text + ", " + txtVideoModel.Text + ", " + txtVRAM.Text + ", " + dropOptical.Text + ", " + checkDrivesNone.Checked + ", " +
                checkDrivesFDD.Checked + ", " + checkDrivesTape.Checked + ", " + txtLCDSize.Text + ", " + checkNetworkNone.Checked + ", " + checkEthernet.Checked + ", " +
                checkModem.Checked + ", " + checkWiFi.Checked + ", " + checkBT.Checked + ", " + dropCOA.Text + ", " + radioOSNo.Checked + ", " + radioOSYes.Checked + ", " +
                txtNotes.Text + ", " + checkAccNone.Checked + ", " + checkAccAC.Checked + ", " + checkAccPower.Checked + ", " + checkAccBatt.Checked + ", " +
                checkAccExtBatt.Checked + ", " + checkAccFinger.Checked + ", " + checkAccWebcam.Checked + ", " + checkAccKeyboard.Checked + ", " + checkAccMouse.Checked + ", " +
                dropDamage.Text + ", " + txtUSB.Text + ", " + txtEthernet.Text + ", " + txtModem.Text + ", " + txtVGA.Text + ", " + txtDVI.Text + ", " + txtSVideo.Text + ", " +
                txtPS2.Text + ", " + txtAudio.Text + ", " + txteSATAp.Text + ", " + txtNumSerial.Text + ", " + txtParallel.Text + ", " + txtPCMCIA.Text + ", " +
                txtSDCard.Text + ", " + txtFirewire.Text + ", " + txteSATA.Text + ", " + txtHDMI.Text + ", " + txtSCSI.Text + ", " + txtDisplayPort.Text + ", " +
                labelVersion.Text + ", " + txtTester.Text;

            File.WriteAllText(filename, output);

        }

        private void performPrint()
        {
            //string filename = AppDomain.CurrentDomain.BaseDirectory + "SpecSheet.lbx";
            string filename = AppDomain.CurrentDomain.BaseDirectory + "SpecSheet.txt";
            System.Diagnostics.Process.Start(filename);

            //IntPtr ptrFF = FindWindow(null, "P-Touch");
            int processExists = 0;
            while (processExists == 0)
            { // this should delay sending keystrokes until the process actually starts
                processExists = Process.GetProcessesByName("Notepad").Length;
            }

            Process p = Process.GetProcessesByName("Notepad").FirstOrDefault();
            if (p != null)
            {
                IntPtr h = p.MainWindowHandle;
                SetForegroundWindow(h);
                SendKeys.SendWait("^p");
                SendKeys.SendWait("~");
                
            }

        }

        private void archiveSerial()
        {
            if (txtSerial.Text == "")
            {
                return;
            }
            /*
            *******old load method from csv files*******

            string directory = AppDomain.CurrentDomain.BaseDirectory + "archive";
            System.IO.Directory.CreateDirectory(directory); // create archive directory if it doesn't already exist 
            string filename = directory + "\\" + txtSerial.Text + ".csv";

            string output = txtISPF.Text + ", " + labelDate.Text + ", " + dropCondition.Text + ", " + dropBrand.Text + ", " + txtSerial.Text + ", " + 
                txtModel.Text + ", " + dropFormfactor.Text + ", " + spinCPUQty.Value + ", " + spinCPUCores.Value + ", " + checkHT.Checked + ", " + txtCPUSpeed.Text + ", " + 
                dropCPUType.Text + ", " + txtBusSpeed.Text + ", " + dropCPUName.Text + ", " + dropMemorySize.Text + ", " + dropMemoryRating.Text + ", " + dropMemoryType.Text + ", " + 
                dropMemorySpeed.Text + ", " + txtWeight.Text + ", " + spinHDDQty.Value + ", " + txtHDDSize.Text + ", " + dropHDDType.Text + ", " + dropHDDRPM.Text + ", " + 
                txtHDDSerial.Text + ", " + dropVideo.Text + ", " + txtVideoModel.Text + ", " + txtVRAM.Text + ", " + dropOptical.Text + ", " + checkDrivesNone.Checked + ", " + 
                checkDrivesFDD.Checked + ", " + checkDrivesTape.Checked + ", " + txtLCDSize.Text + ", " + checkNetworkNone.Checked + ", " + checkEthernet.Checked + ", " + 
                checkModem.Checked + ", " + checkWiFi.Checked + ", " + checkBT.Checked + ", " + dropCOA.Text + ", " + radioOSNo.Checked + ", " + radioOSYes.Checked + ", " + 
                txtNotes.Text + ", " + checkAccNone.Checked + ", " + checkAccAC.Checked + ", " + checkAccPower.Checked + ", " + checkAccBatt.Checked + ", " + 
                checkAccExtBatt.Checked + ", " + checkAccFinger.Checked + ", " + checkAccWebcam.Checked + ", " + checkAccKeyboard.Checked + ", " + checkAccMouse.Checked + ", " + 
                dropDamage.Text + ", " + txtUSB.Text + ", " + txtEthernet.Text + ", " + txtModem.Text + ", " + txtVGA.Text + ", " + txtDVI.Text + ", " + txtSVideo.Text + ", " + 
                txtPS2.Text + ", " + txtAudio.Text + ", " + txteSATAp.Text + ", " + txtNumSerial.Text + ", " + txtParallel.Text + ", " + txtPCMCIA.Text + ", " + 
                txtSDCard.Text + ", " + txtFirewire.Text + ", " + txteSATA.Text + ", " + txtHDMI.Text + ", " + txtSCSI.Text + ", " + txtDisplayPort.Text + ", " + 
                labelVersion.Text + ", " + txtTester.Text;

            File.WriteAllText(filename, output);
            */

            // new load method using sqlite database
            executeSQL("UPDATE OR IGNORE archive SET "+
                "ispf='" + txtISPF.Text + "', "+
                "date='" + labelDate.Text + "', " +
                "condition='" + dropCondition.Text + "', " +
                "brand='" + dropBrand.Text + "', " +
                //"serial='" + txtSerial.Text + "', " +
                "model='" + txtModel.Text + "', " +
                "formfactor='" + dropFormfactor.Text + "', " +
                "cpuqty='" + spinCPUQty.Value + "', " +
                "cpucores='" + spinCPUCores.Value + "', " +
                "checkht='" + checkHT.Checked + "', " +
                "cpuspeed='" + txtCPUSpeed.Text + "', " +
                "cputype='" + dropCPUType.Text + "', " +
                "busspeed='" + txtBusSpeed.Text + "', " +
                "cpuname='" + dropCPUName.Text + "', " +
                "memorysize='" + dropMemorySize.Text + "', " +
                "memoryrating='" + dropMemoryRating.Text + "', " +
                "memorytype='" + dropMemoryType.Text + "', " +
                "memoryspeed='" + dropMemorySpeed.Text + "', " +
                "weight='" + txtWeight.Text + "', " +
                "hddqty='" + spinHDDQty.Value + "', " +
                "hddsize='" + txtHDDSize.Text + "', " +
                "hddtype='" + dropHDDType.Text + "', " +
                "hddrpm='" + dropHDDRPM.Text + "', " +
                "hddserial='" + txtHDDSerial.Text + "', " +
                "video='" + dropVideo.Text + "', " +
                "videomodel='" + txtVideoModel.Text + "', " +
                "vram='" + txtVRAM.Text + "', " +
                "optical='" + dropOptical.Text + "', " +
                "drivesnone='" + checkDrivesNone.Checked + "', " +
                "drivesfdd='" + checkDrivesFDD.Checked + "', " +
                "drivestape='" + checkDrivesTape.Checked + "', " +
                "lcdsize='" + txtLCDSize.Text + "', " +
                "networknone='" + checkNetworkNone.Checked + "', " +
                "ethernet='" + checkEthernet.Checked + "', " +
                "modem='" + checkModem.Checked + "', " +
                "wifi='" + checkWiFi.Checked + "', " +
                "bt='" + checkBT.Checked + "', " +
                "coa='" + dropCOA.Text + "', " +
                "osno='" + radioOSNo.Checked + "', " +
                "osyes='" + radioOSYes.Checked + "', " +
                "notes='" + txtNotes.Text.Replace("'", "''") + "', " +
                "accnone='" + checkAccNone.Checked + "', " +
                "accac='" + checkAccAC.Checked + "', " +
                "accpower='" + checkAccPower.Checked + "', " +
                "accbatt='" + checkAccBatt.Checked + "', " +
                "accextbatt='" + checkAccExtBatt.Checked + "', " +
                "accfinger='" + checkAccFinger.Checked + "', " +
                "accwebcam='" + checkAccWebcam.Checked + "', " +
                "acckeyboard='" + checkAccKeyboard.Checked + "', " +
                "accmouse='" + checkAccMouse.Checked + "', " +
                "damage='" + dropDamage.Text + "', " +
                "usb='" + txtUSB.Text + "', " +
                "numethernet='" + txtEthernet.Text + "', " +
                "nummodem='" + txtModem.Text + "', " +
                "vga='" + txtVGA.Text + "', " +
                "dvi='" + txtDVI.Text + "', " +
                "svideo='" + txtSVideo.Text + "', " +
                "ps2='" + txtPS2.Text + "', " +
                "audio='" + txtAudio.Text + "', " +
                "esatap='" + txteSATAp.Text + "', " +
                "numserial='" + txtNumSerial.Text + "', " +
                "parallel='" + txtParallel.Text + "', " +
                "pcmcia='" + txtPCMCIA.Text + "', " +
                "sdcard='" + txtSDCard.Text + "', " +
                "firewire='" + txtFirewire.Text + "', " +
                "esata='" + txteSATA.Text + "', " +
                "hdmi='" + txtHDMI.Text + "', " +
                "scsi='" + txtSCSI.Text + "', " +
                "displayport='" + txtDisplayPort.Text + "', " +
                "version='" + labelVersion.Text + "', " +
                "tester='" + txtTester.Text + "', "+
                "caddyqty='" + spinCaddyQTY.Value + "', "+
                "caddyna='" + checkCaddyNA.Checked + "', "+
                "checkmcf='" + checkMCF.Checked + "', "+
                "checkfg='" + checkFG.Checked + "' "+
                "WHERE serial='" + txtSerial.Text + "'");

            executeSQL("INSERT OR IGNORE INTO archive (ispf, date, condition, brand, serial, model, formfactor, cpuqty, cpucores, checkht, cpuspeed, cputype, busspeed, cpuname, memorysize, memoryrating, memorytype, memoryspeed, weight, hddqty, hddsize, hddtype, hddrpm, hddserial, video, videomodel, vram, optical, drivesnone, drivesfdd, drivestape, lcdsize, networknone, ethernet, modem, wifi, bt, coa, osno, osyes, notes, accnone, accac, accpower, accbatt, accextbatt, accfinger, accwebcam, acckeyboard, accmouse, damage, usb, numethernet, nummodem, vga, dvi, svideo, ps2, audio, esatap, numserial, parallel, pcmcia, sdcard, firewire, esata, hdmi, scsi, displayport, version, tester, caddyqty, caddyna, checkmcf, checkfg) "+
                "VALUES ('" + txtISPF.Text + "', " +
                "'" + labelDate.Text + "', " +
                "'" + dropCondition.Text + "', " +
                "'" + dropBrand.Text + "', " +
                "'" + txtSerial.Text + "', " +
                "'" + txtModel.Text + "', " +
                "'" + dropFormfactor.Text + "', " +
                "'" + spinCPUQty.Value + "', " +
                "'" + spinCPUCores.Value + "', " +
                "'" + checkHT.Checked + "', " +
                "'" + txtCPUSpeed.Text + "', " +
                "'" + dropCPUType.Text + "', " +
                "'" + txtBusSpeed.Text + "', " +
                "'" + dropCPUName.Text + "', " +
                "'" + dropMemorySize.Text + "', " +
                "'" + dropMemoryRating.Text + "', " +
                "'" + dropMemoryType.Text + "', " +
                "'" + dropMemorySpeed.Text + "', " +
                "'" + txtWeight.Text + "', " +
                "'" + spinHDDQty.Value + "', " +
                "'" + txtHDDSize.Text + "', " +
                "'" + dropHDDType.Text + "', " +
                "'" + dropHDDRPM.Text + "', " +
                "'" + txtHDDSerial.Text + "', " +
                "'" + dropVideo.Text + "', " +
                "'" + txtVideoModel.Text + "', " +
                "'" + txtVRAM.Text + "', " +
                "'" + dropOptical.Text + "', " +
                "'" + checkDrivesNone.Checked + "', " +
                "'" + checkDrivesFDD.Checked + "', " +
                "'" + checkDrivesTape.Checked + "', " +
                "'" + txtLCDSize.Text + "', " +
                "'" + checkNetworkNone.Checked + "', " +
                "'" + checkEthernet.Checked + "', " +
                "'" + checkModem.Checked + "', " +
                "'" + checkWiFi.Checked + "', " +
                "'" + checkBT.Checked + "', " +
                "'" + dropCOA.Text + "', " +
                "'" + radioOSNo.Checked + "', " +
                "'" + radioOSYes.Checked + "', " +
                "'" + txtNotes.Text.Replace("'", "''") + "', " +
                "'" + checkAccNone.Checked + "', " +
                "'" + checkAccAC.Checked + "', " +
                "'" + checkAccPower.Checked + "', " +
                "'" + checkAccBatt.Checked + "', " +
                "'" + checkAccExtBatt.Checked + "', " +
                "'" + checkAccFinger.Checked + "', " +
                "'" + checkAccWebcam.Checked + "', " +
                "'" + checkAccKeyboard.Checked + "', " +
                "'" + checkAccMouse.Checked + "', " +
                "'" + dropDamage.Text + "', " +
                "'" + txtUSB.Text + "', " +
                "'" + txtEthernet.Text + "', " +
                "'" + txtModem.Text + "', " +
                "'" + txtVGA.Text + "', " +
                "'" + txtDVI.Text + "', " +
                "'" + txtSVideo.Text + "', " +
                "'" + txtPS2.Text + "', " +
                "'" + txtAudio.Text + "', " +
                "'" + txteSATAp.Text + "', " +
                "'" + txtNumSerial.Text + "', " +
                "'" + txtParallel.Text + "', " +
                "'" + txtPCMCIA.Text + "', " +
                "'" + txtSDCard.Text + "', " +
                "'" + txtFirewire.Text + "', " +
                "'" + txteSATA.Text + "', " +
                "'" + txtHDMI.Text + "', " +
                "'" + txtSCSI.Text + "', " +
                "'" + txtDisplayPort.Text + "', " +
                "'" + labelVersion.Text + "', " +
                "'" + txtTester.Text + "', " +
                "'" + spinCaddyQTY.Value + "', " +
                "'" + checkCaddyNA.Checked + "', " + 
                "'" + checkMCF.Checked + "', " +
                "'" + checkFG.Checked + "')"


                );

        }

        private void loadSerialFromArchive(string s)
        {
            /*
            *******old load method from csv files*******

            string directory = AppDomain.CurrentDomain.BaseDirectory + "archive";
            System.IO.Directory.CreateDirectory(directory); // create archive directory if it doesn't already exist
            string filename = directory + "\\" + s + ".csv";

            if (File.Exists(filename))
            {
                string[] input = new string[71];
                input = File.ReadAllText(filename).Split(new[] { ", " }, StringSplitOptions.None);

                if (input.Length != 72)
                {
                    return;
                }

                
                txtISPF.Text = input[0];
                //labelDate.Text = input[1];
                dropCondition.Text = input[2];
                dropBrand.Text = input[3];
                txtSerial.Text = input[5];
                txtModel.Text = input[6];
                dropFormfactor.Text = input[7];
                spinCPUQty.Value = int.Parse(input[8]);
                spinCPUCores.Value = int.Parse(input[9]);
                checkHT.Checked = bool.Parse(input[10]);
                txtCPUSpeed.Text = input[11];
                dropCPUType.Text = input[12];
                txtBusSpeed.Text = input[13];
                dropCPUName.Text = input[14];
                dropMemorySize.Text = input[15];
                dropMemoryRating.Text = input[16];
                dropMemoryType.Text = input[17];
                dropMemorySpeed.Text = input[18];
                txtWeight.Text = input[19];
                spinHDDQty.Value = int.Parse(input[20]);
                txtHDDSize.Text = input[21];
                dropHDDType.Text = input[22];
                dropHDDRPM.Text = input[23];
                txtHDDSerial.Text = input[24];
                dropVideo.Text = input[25];
                txtVideoModel.Text = input[26];
                txtVRAM.Text = input[27];
                dropOptical.Text = input[28];
                checkDrivesNone.Checked = bool.Parse(input[29]);
                checkDrivesFDD.Checked = bool.Parse(input[30]);
                checkDrivesTape.Checked = bool.Parse(input[31]);
                txtLCDSize.Text = input[32];
                checkNetworkNone.Checked = bool.Parse(input[33]);
                checkEthernet.Checked = bool.Parse(input[34]);
                checkModem.Checked = bool.Parse(input[35]);
                checkWiFi.Checked = bool.Parse(input[36]);
                checkBT.Checked = bool.Parse(input[37]);
                dropCOA.Text = input[38];
                radioOSNo.Checked = bool.Parse(input[39]);
                radioOSYes.Checked = bool.Parse(input[40]);
                txtNotes.Text = input[41];
                checkAccNone.Checked = bool.Parse(input[42]);
                checkAccAC.Checked = bool.Parse(input[43]);
                checkAccPower.Checked = bool.Parse(input[44]);
                checkAccBatt.Checked = bool.Parse(input[45]);
                checkAccExtBatt.Checked = bool.Parse(input[46]);
                checkAccFinger.Checked = bool.Parse(input[47]);
                checkAccWebcam.Checked = bool.Parse(input[48]);
                checkAccKeyboard.Checked = bool.Parse(input[49]);
                checkAccMouse.Checked = bool.Parse(input[50]);
                dropDamage.Text = input[51];
                txtUSB.Text = input[52];
                txtEthernet.Text = input[53];
                txtModem.Text = input[54];
                txtVGA.Text = input[55];
                txtDVI.Text = input[56];
                txtSVideo.Text = input[57];
                txtPS2.Text = input[58];
                txtAudio.Text = input[59];
                txteSATAp.Text = input[60];
                txtNumSerial.Text = input[61];
                txtParallel.Text = input[62];
                txtPCMCIA.Text = input[63];
                txtSDCard.Text = input[64];
                txtFirewire.Text = input[65];
                txteSATA.Text = input[66];
                txtHDMI.Text = input[67];
                txtSCSI.Text = input[68];
                txtDisplayPort.Text = input[69];
                //labelVersion.Text = input[70];
                txtTester.Text = input[71];

            }

            */

            // new load method using sqlite database
            SQLiteDataReader input = executeSQLReader("SELECT * FROM archive WHERE serial='" + s + "' ORDER BY timestamp DESC");
            while (input.Read())
            {

                txtISPF.Text = (string)input["ispf"];
                //labelDate.Text = input[1];
                dropCondition.Text = (string)input["condition"];
                dropBrand.Text = (string)input["brand"];
                //txtSerial.Text = (string)input["serial"];
                txtModel.Text = (string)input["model"];
                dropFormfactor.Text = (string)input["formfactor"];
                spinCPUQty.Value = int.Parse((string)input["cpuqty"]);
                spinCPUCores.Value = int.Parse((string)input["cpucores"]);
                checkHT.Checked = bool.Parse((string)input["checkht"]);
                txtCPUSpeed.Text = (string)input["cpuspeed"];
                dropCPUType.Text = (string)input["cputype"];
                txtBusSpeed.Text = (string)input["busspeed"];
                dropCPUName.Text = (string)input["cpuname"];
                dropMemorySize.Text = (string)input["memorysize"];
                dropMemoryRating.Text = (string)input["memoryrating"];
                dropMemoryType.Text = (string)input["memorytype"];
                dropMemorySpeed.Text = (string)input["memoryspeed"];
                txtWeight.Text = (string)input["weight"];
                spinHDDQty.Value = int.Parse((string)input["hddqty"]);
                txtHDDSize.Text = (string)input["hddsize"];
                dropHDDType.Text = (string)input["hddtype"];
                dropHDDRPM.Text = (string)input["hddrpm"];
                txtHDDSerial.Text = (string)input["hddserial"];
                dropVideo.Text = (string)input["video"];
                txtVideoModel.Text = (string)input["videomodel"];
                txtVRAM.Text = (string)input["vram"];
                dropOptical.Text = (string)input["optical"];
                checkDrivesNone.Checked = bool.Parse((string)input["drivesnone"]);
                checkDrivesFDD.Checked = bool.Parse((string)input["drivesfdd"]);
                checkDrivesTape.Checked = bool.Parse((string)input["drivestape"]);
                txtLCDSize.Text = (string)input["lcdsize"];
                checkNetworkNone.Checked = bool.Parse((string)input["networknone"]);
                checkEthernet.Checked = bool.Parse((string)input["ethernet"]);
                checkModem.Checked = bool.Parse((string)input["modem"]);
                checkWiFi.Checked = bool.Parse((string)input["wifi"]);
                checkBT.Checked = bool.Parse((string)input["bt"]);
                dropCOA.Text = (string)input["coa"];
                radioOSNo.Checked = bool.Parse((string)input["osno"]);
                radioOSYes.Checked = bool.Parse((string)input["osyes"]);
                txtNotes.Text = (string)input["notes"];
                checkAccNone.Checked = bool.Parse((string)input["accnone"]);
                checkAccAC.Checked = bool.Parse((string)input["accac"]);
                checkAccPower.Checked = bool.Parse((string)input["accpower"]);
                checkAccBatt.Checked = bool.Parse((string)input["accbatt"]);
                checkAccExtBatt.Checked = bool.Parse((string)input["accextbatt"]);
                checkAccFinger.Checked = bool.Parse((string)input["accfinger"]);
                checkAccWebcam.Checked = bool.Parse((string)input["accwebcam"]);
                checkAccKeyboard.Checked = bool.Parse((string)input["acckeyboard"]);
                checkAccMouse.Checked = bool.Parse((string)input["accmouse"]);
                dropDamage.Text = (string)input["damage"];
                txtUSB.Text = (string)input["usb"];
                txtEthernet.Text = (string)input["numethernet"];
                txtModem.Text = (string)input["nummodem"];
                txtVGA.Text = (string)input["vga"];
                txtDVI.Text = (string)input["dvi"];
                txtSVideo.Text = (string)input["svideo"];
                txtPS2.Text = (string)input["ps2"];
                txtAudio.Text = (string)input["audio"];
                txteSATAp.Text = (string)input["esatap"];
                txtNumSerial.Text = (string)input["numserial"];
                txtParallel.Text = (string)input["parallel"];
                txtPCMCIA.Text = (string)input["pcmcia"];
                txtSDCard.Text = (string)input["sdcard"];
                txtFirewire.Text = (string)input["firewire"];
                txteSATA.Text = (string)input["esata"];
                txtHDMI.Text = (string)input["hdmi"];
                txtSCSI.Text = (string)input["scsi"];
                txtDisplayPort.Text = (string)input["displayport"];
                //labelVersion.Text = input[70];
                txtTester.Text = (string)input["tester"];
                spinCaddyQTY.Value = int.Parse((string)input["caddyqty"]);
                checkCaddyNA.Checked = bool.Parse((string)input["caddyna"]);
            }


        }

        private void txtSerial_TextChanged(object sender, EventArgs e)
        {
            loadSerialFromArchive(txtSerial.Text);
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            Controls.Clear();
            InitializeComponent();
            setDate();

            m_dbConnection.Close();
            
            initializeDatabase();
            initializeDropdowns();
        }

        private void SpecSheet_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_dbConnection.Close();
            
        }

        private void dropFormfactor_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropBrand_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropCondition_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropHDDType_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropHDDRPM_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropVideo_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropOptical_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropCPUType_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropCOA_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropMemorySize_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropMemoryRating_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropMemoryType_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropMemorySpeed_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropDamage_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropCPUType_SelectedValueChanged(object sender, EventArgs e)
        {
            dropCPUName.Text = "";
        }

        private void dropCPUType_TextUpdate(object sender, EventArgs e)
        {
            dropCPUName.Text = "";
        }

        private void adminButton_Click(object sender, EventArgs e)
        {
            AdminMenu adminWindow = new AdminMenu();
            adminWindow.Show();
        }

        private void dropCPUName_Leave(object sender, EventArgs e)
        {
            saveDropdown((ComboBox)sender);
        }

        private void dropCPUName_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // load up cpu informations from database
            SQLiteDataReader result = executeSQLReader("SELECT * FROM cpu_data WHERE cpuseries='" + dropCPUName.Text + "'");
            while (result.Read())
            {
                dropCPUType.Text = (string)result["cputype"];
                txtBusSpeed.Text = (string)result["busspeed"];
                txtCPUSpeed.Text = (string)result["cpuspeed"];
                spinCPUCores.Value = int.Parse((string)result["cpucores"]);
                checkHT.Checked = bool.Parse((string)result["cpuht"]);
            }
        }


        private void initdropCPUName(object sender, EventArgs e)
        {
            //load up cpu list from database
            SQLiteDataReader result = executeSQLReader("SELECT * FROM cpu_data WHERE model='" + txtModel.Text + "' AND formfactor='" + dropFormfactor.Text + "'");
            if (result.HasRows)
            {
                dropCPUName.Items.Clear();
                while (result.Read())
                {
                    dropCPUName.Items.Add(result["cpuseries"]);
                }
            } else
            {
                initDropdown(dropCPUName);
            }

        }

        private string[,] memory_matrix = new string[15,3] { 
                          { "DDR", "100MHz", "PC-1600" },
                          { "DDR", "133MHz", "PC-2100" },
                          { "DDR", "166MHz", "PC-2700" },
                          { "DDR", "200MHz", "PC-3200" },
                          { "DDR2", "400MHz", "PC2-3200" },
                          { "DDR2", "533MHz", "PC2-4200" },
                          { "DDR2", "667MHz", "PC2-5300" },
                          { "DDR2", "800MHz", "PC2-6400" },
                          { "DDR2", "1066MHz", "PC2-8500" },
                          { "DDR3", "800MHz", "PC3-6400" },
                          { "DDR3", "1066MHz", "PC3-8500" },
                          { "DDR3", "1333MHz", "PC3-10600" },
                          { "DDR3", "1600MHz", "PC3-12800" },
                          { "DDR3", "1866MHz", "PC3-14900" },
                          { "DDR3", "2133MHz", "PC3-17000" }
                        };

        private void dropMemoryRating_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i=0;i < (memory_matrix.Length/3);i++) 
            {
                if (dropMemoryRating.Text.Equals(memory_matrix[i,2]))
                {
                    dropMemorySpeed.Text = memory_matrix[i,1];
                    dropMemoryType.Text = memory_matrix[i,0];
                }
            }

        }

        private void buttonAddModule_Click(object sender, EventArgs e)
        {
            string module = dropMemorySize.Text + " " + dropMemoryType.Text + " " + dropMemoryRating.Text;
            listMemoryModules.Items.Add(module);

        }

        private void buttonRemoveModule_Click(object sender, EventArgs e)
        {
            if (listMemoryModules.SelectedIndex != -1)
            {
                for (int i = listMemoryModules.SelectedItems.Count - 1; i >= 0; i--)
                    listMemoryModules.Items.Remove(listMemoryModules.SelectedItems[i]);
            }
            listMemoryModules.ClearSelected();
        }

        private void checkFG_CheckedChanged(object sender, EventArgs e)
        {
            buttonAddModule.Visible = checkFG.Checked;
            buttonRemoveModule.Visible = checkFG.Checked;
            frameMemoryModules.Visible = checkFG.Checked;
            listMemoryModules.Visible = checkFG.Checked;
            frameMemoryType.Visible = !checkFG.Checked;
            dropMemoryType.Visible = !checkFG.Checked;
            frameMemorySpeed.Visible = !checkFG.Checked;
            dropMemorySpeed.Visible = !checkFG.Checked;
        }
    }
}
