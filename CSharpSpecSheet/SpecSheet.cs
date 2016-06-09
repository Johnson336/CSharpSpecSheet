using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpSpecSheet
{
    public partial class SpecSheet : Form
    {
        public SpecSheet()
        {
            InitializeComponent();
            setDate();
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

        private void printButton_Click(object sender, EventArgs e)
        {
            archiveSerial();

        }

        private void archiveSerial()
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory + "archive";
            System.IO.Directory.CreateDirectory(directory); /* create archive directory if it doesn't already exist */
            string filename = directory + "\\" + txtSerial.Text + ".csv";

            string output = txtISPF.Text + ", " + labelDate.Text + ", " + dropCondition.Text + ", " + dropBrand.Text + ", " + txtBrandOther.Text + ", " + txtSerial.Text + ", " + 
                txtModel.Text + ", " + dropFormfactor.Text + ", " + spinCPUQty.Value + ", " + spinCPUCores.Value + ", " + checkHT.Checked + ", " + txtCPUSpeed.Text + ", " + 
                dropCPUType.Text + ", " + txtBusSpeed.Text + ", " + txtCPUName.Text + ", " + dropMemorySize.Text + ", " + dropMemoryRating.Text + ", " + dropMemoryType.Text + ", " + 
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

        private void loadSerialFromArchive(string s)
        {

            string directory = AppDomain.CurrentDomain.BaseDirectory + "archive";
            System.IO.Directory.CreateDirectory(directory); /* create archive directory if it doesn't already exist */
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
                labelDate.Text = input[1];
                dropCondition.Text = input[2];
                dropBrand.Text = input[3];
                txtBrandOther.Text = input[4];
                txtSerial.Text = input[5];
                txtModel.Text = input[6];
                dropFormfactor.Text = input[7];
                spinCPUQty.Value = int.Parse(input[8]);
                spinCPUCores.Value = int.Parse(input[9]);
                checkHT.Checked = bool.Parse(input[10]);
                txtCPUSpeed.Text = input[11];
                dropCPUType.Text = input[12];
                txtBusSpeed.Text = input[13];
                txtCPUName.Text = input[14];
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
                labelVersion.Text = input[70];
                txtTester.Text = input[71];





            }


        }

        private void txtSerial_TextChanged(object sender, EventArgs e)
        {
            loadSerialFromArchive(txtSerial.Text);
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            this.Controls.Clear();
            this.InitializeComponent();
            this.setDate();
        }
    }
}
