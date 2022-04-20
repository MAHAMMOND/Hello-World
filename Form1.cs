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


namespace PowerSupplyLogger_WINDOW
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Icon = Properties.Resources.microchip;
        }


        private void LoadButton_Click(object sender, EventArgs e)
        /*
         Loads txt file program and sets up application to test boards
         */
        {
            DisableAllButtons();
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {

                    try//to open file dialog
                    {
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;

                        if (filePath.Substring(filePath.Length - 4) != ".txt")
                        {
                            SanityException SE = new SanityException("You tried to open something that wasn't a text file. Please make sure you selected the right file and that the file extention is .txt");
                            throw SE;
                        }

                        //Read the contents of the file into a stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            Program.excelFilePath = reader.ReadLine();
                            Program.voltageMin = Convert.ToDouble(reader.ReadLine());
                            Program.voltageMax = Convert.ToDouble(reader.ReadLine());
                            Program.currentMin = Convert.ToDouble(reader.ReadLine());
                            Program.currentMax = Convert.ToDouble(reader.ReadLine());
                            Program.powerSupplyAddress = reader.ReadLine();
                            Program.testCurrent = Convert.ToDouble(reader.ReadLine());
                            Program.testVoltage = Convert.ToDouble(reader.ReadLine());
                            reader.Close();

                        }
                        Program.sanityCheck();
                       
                        try//to connect to power supply and update GUI
                        {
                            Program.initializePowerSupply(Program.powerSupplyAddress);
                            Program.setTestCurrent(Program.testCurrent);
                            Program.setTestVoltage(Program.testVoltage);
                            UpdateVoltageCurrentExpectedLabels();
                 
                            this.ResetButton.Enabled = true;
                            //ResetButton.Click += ResetButton_Click;
                            this.ResetButton.ForeColor = SystemColors.ControlText;
                            this.NewPartNumberButton.Enabled = true;
                            //NewPartNumberButton.Click += NewPartNumberButton_Click;
                            this.NewPartNumberButton.ForeColor = SystemColors.ControlText;
                            ResetAll();
                            this.NewPartNumberButton.Focus();
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not connect to power supply! Please connect and retry.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }


                    }
                    catch (SanityException ex)
                    {
                        MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not read program text file correctly. Please make sure it is in the correct format (it should be a txt file)", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }



                }
            }
            Application.DoEvents();
            this.LoadButton.Enabled = true;
            //LoadButton.Click += LoadButton_Click;

        }

        private void UpdateVoltageCurrentExpectedLabels()
        {
            string VoltageString = Convert.ToString(Program.voltageMin) + " - " + Convert.ToString(Program.voltageMax);
            //multiplying current by 1000 to convert to mA from A
            string CurrentString = Convert.ToString(Program.currentMin * 1000) + " - " + Convert.ToString(Program.currentMax * 1000);
            ExpectedVoltageRangeLabel_Content.Text = VoltageString;
            ExpectedCurrentRangeLabel_Content.Text = CurrentString;
            ExpectedVoltageRangeLabel_Content.Refresh();
            ExpectedCurrentRangeLabel_Content.Refresh();

        }

        private void BoardTestFailedLabelUpdate()
        {
            this.PassLabel.ForeColor = SystemColors.ControlDark;
            this.PassLabel.BackColor = SystemColors.Control;
            this.FailLabel.ForeColor = Color.DarkRed;
            this.FailLabel.BackColor = Color.Tomato;
            this.PassLabel.Refresh();
            this.FailLabel.Refresh();
        }

        private void BoardTestPassedLabelUpdate()
        {
            this.PassLabel.ForeColor = Color.Green;
            this.PassLabel.BackColor = Color.LightGreen;
            this.FailLabel.ForeColor = SystemColors.ControlDark;
            this.FailLabel.BackColor = SystemColors.Control;
            this.PassLabel.Refresh();
            this.FailLabel.Refresh();
        }

        private void BoardTestNeutralLabelUpdate()
        {
            this.PassLabel.ForeColor = SystemColors.ControlDark;
            this.PassLabel.BackColor = SystemColors.Control;
            this.FailLabel.ForeColor = SystemColors.ControlDark;
            this.FailLabel.BackColor = SystemColors.Control;
            this.PassLabel.Refresh();
            this.FailLabel.Refresh();
        }

        private void VoltagePassedLabelUpdate_Passed()
        {
            VoltagePassedLabel.Text = "PASSED";
            VoltagePassedLabel.ForeColor = Color.Green;
            VoltagePassedLabel.BackColor = Color.Aquamarine;
            VoltagePassedLabel.Refresh();
        }

        private void VoltagePassedLabelUpdate_Failed()
        {
            VoltagePassedLabel.Text = "FAILED";
            VoltagePassedLabel.ForeColor = Color.Red;
            VoltagePassedLabel.BackColor = Color.Pink;
            VoltagePassedLabel.Refresh();
        }

        private void VoltagePassedLabelUpdate_Neutral()
        {
            VoltagePassedLabel.Text = "-------";
            VoltagePassedLabel.ForeColor = SystemColors.ControlDark;
            VoltagePassedLabel.BackColor = SystemColors.Control;
            VoltagePassedLabel.Refresh();
        }

        private void CurrentPassedLabelUpdate_Passed()
        {
            CurrentPassedLabel.Text = "PASSED";
            CurrentPassedLabel.ForeColor = Color.Green;
            CurrentPassedLabel.BackColor = Color.Aquamarine;
            CurrentPassedLabel.Refresh();
        }

        private void CurrentPassedLabelUpdate_Failed()
        {
            CurrentPassedLabel.Text = "FAILED";
            CurrentPassedLabel.ForeColor = Color.Red;
            CurrentPassedLabel.BackColor = Color.Pink;
            CurrentPassedLabel.Refresh();
        }

        private void CurrentPassedLabelUpdate_Neutral()
        {
            CurrentPassedLabel.Text = "-------";
            CurrentPassedLabel.ForeColor = SystemColors.ControlDark;
            CurrentPassedLabel.BackColor = SystemColors.Control;
            CurrentPassedLabel.Refresh();
        }







        private void ResetButton_Click(object sender, EventArgs e)
        {
            ResetAll();
        }

        private void ResetAll()
        {
            Program.measuredCurrent = 0.0;
            Program.measuredVoltage = 0.0;
            Program.currentPassed = false;
            Program.voltagePassed = false;
            Program.boardPassed = false;
            Program.boardPartNumber = "";
            this.MeasuredVoltageLabel_Content.Text = "-------";
            this.MeasuredCurrentLabel_Content.Text = "-------";
            this.BoardPNLabel_Content.Text = "-------";
            BoardTestNeutralLabelUpdate();
            VoltagePassedLabelUpdate_Neutral();
            CurrentPassedLabelUpdate_Neutral();
            this.MeasureButton.Enabled = false;
            //MeasureButton.Click -= MeasureButton_Click;
        }

        private void ResetAllButPartNumber()
        {
            Program.measuredCurrent = 0.0;
            Program.measuredVoltage = 0.0;
            Program.currentPassed = false;
            Program.voltagePassed = false;
            Program.boardPassed = false;
            this.MeasuredVoltageLabel_Content.Text = "-------";
            this.MeasuredCurrentLabel_Content.Text = "-------";
            BoardTestNeutralLabelUpdate();
            VoltagePassedLabelUpdate_Neutral();
            CurrentPassedLabelUpdate_Neutral();
        }

        private void getTextScanner()
        //create new pop-up window with text box to fill in with scanner and update board part number from Program
        {

            PartNumberForm pnForm = new PartNumberForm();
            if (pnForm.ShowDialog(this) == DialogResult.OK)
            {
                Program.boardPartNumber = pnForm.CPN;
                ResetAllButPartNumber();
                //Console.WriteLine(Program.boardPartNumber);
                this.MeasureButton.Enabled = true;
                //MeasureButton.Click += MeasureButton_Click;

            }
            pnForm.Dispose();


        }



        private void FailLabel_Click(object sender, EventArgs e)
        {

        }

        public void Set_BoardPN(string inString)
        {
            Program.boardPartNumber = inString;
            //Console.WriteLine(Program.boardPartNumber);
        }

        private void NewPartNumberButton_Click(object sender, EventArgs e)
        {
            DisableAllButtons();
            //ResetAllButPartNumber();
            getTextScanner();
            BoardPNLabel_Content.Text = Program.boardPartNumber;
            if (Program.boardPartNumber == "")
            {
                this.BoardPNLabel_Content.Text = "-------";
            }
            Application.DoEvents();
            this.LoadButton.Enabled = true;
            //LoadButton.Click += LoadButton_Click;
            this.NewPartNumberButton.Enabled = true;
            //NewPartNumberButton.Click += NewPartNumberButton_Click;
            this.ResetButton.Enabled = true;
            //ResetButton.Click += ResetButton_Click;
            BoardPNLabel_Content.Refresh();
            this.MeasureButton.Focus();

        }



        private void SaveToCSV()
        {
            try
            {
                Program.testTime = DateTime.Now;
                using (FileStream filestream = new FileStream(Program.excelFilePath, FileMode.Append, FileAccess.Write))
                {
                    using (StreamWriter streamwriter = new StreamWriter(filestream))
                    {
                        var line = string.Format("{0},{1},{2},{3},{4}", Program.boardPartNumber, Convert.ToString(Program.measuredVoltage), Convert.ToString(Program.measuredCurrent), Convert.ToString(Program.boardPassed), Program.testTime.ToString());
                        streamwriter.WriteLine(line);
                        streamwriter.Flush();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not save to CSV file! Please make sure it is not open and that the file path is correct. ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void MeasureVoltage_Click()
        {
            Program.measuredVoltage = Program.measureVoltage();
            MeasuredVoltageLabel_Content.Text = Convert.ToString(Program.measuredVoltage);
            MeasuredVoltageLabel_Content.Refresh();
        }

        private void MeasureCurrent_Click()
        {
            Program.measuredCurrent = Program.measureCurrent();
            MeasuredCurrentLabel_Content.Text = Convert.ToString(1000 * Program.measuredCurrent); //multiply by 1000 to convert to mA
            MeasuredCurrentLabel_Content.Refresh();
        }

        private void DisableAllButtons()
        {
            ResetButton.Enabled = false;
            //ResetButton.Click -= ResetButton_Click;
            LoadButton.Enabled = false;
            //LoadButton.Click -= LoadButton_Click;
            MeasureButton.Enabled = false;
            //MeasureButton.Click -= MeasureButton_Click;
            NewPartNumberButton.Enabled = false;
            //NewPartNumberButton.Click -= NewPartNumberButton_Click;
        }

        private void EnableAllButtons()
        {
            ResetButton.Enabled = true;
            //ResetButton.Click += ResetButton_Click;
            LoadButton.Enabled = true;
            //LoadButton.Click += LoadButton_Click;
            MeasureButton.Enabled = true;
            //MeasureButton.Click += MeasureButton_Click;
            NewPartNumberButton.Enabled = true;
            //NewPartNumberButton.Click += NewPartNumberButton_Click;
        }

        private void MeasureButton_Click(object sender, EventArgs e)
        {
            DisableAllButtons();
            Program.turnOutput_On();
            System.Threading.Thread.Sleep(3000);//Wait 1 second for power supply to turn on
            MeasureVoltage_Click();
            MeasureCurrent_Click();
            Program.turnOutput_Off();
            if (Program.measuredCurrent > Program.currentMin && Program.measuredCurrent < Program.currentMax)
            //current passed within bounds
            {
                Program.currentPassed = true;
                CurrentPassedLabelUpdate_Passed();
            }
            else
            //current did not pass
            {
                Program.currentPassed = false;
                CurrentPassedLabelUpdate_Failed();
            }
            if (Program.measuredVoltage > Program.voltageMin && Program.measuredVoltage < Program.voltageMax)
            //voltage passed within bounds
            {
                Program.voltagePassed = true;
                VoltagePassedLabelUpdate_Passed();
            }
            else
            //voltage did not pass
            {
                Program.voltagePassed = false;
                VoltagePassedLabelUpdate_Failed();
            }
            if (Program.voltagePassed && Program.currentPassed)
            //board passed
            {
                Program.boardPassed = true;
                BoardTestPassedLabelUpdate();
            }
            else
            //board failed
            {
                Program.boardPassed = false;
                BoardTestFailedLabelUpdate();
            }
            SaveToCSV();
            Application.DoEvents();
            EnableAllButtons();
            this.NewPartNumberButton.Focus();
        }

    }
    
    public class SanityException : Exception
    {
        public SanityException() : base() { }
        public SanityException(string message) : base(message) { }
        public SanityException(string message, Exception e) : base(message, e) { }

        private string info;
        public string Info
        {
            get { return info; }
            set { info = value; }
        }
    }







}
