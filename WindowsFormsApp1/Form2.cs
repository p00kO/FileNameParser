using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        private bool calibrationValueshanged = false;
        DataSet ds;
        public Form2()
        {
            InitializeComponent();
            LoadXMLCalibrationFile();
            calibrationValueshanged = false;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (calibrationValueshanged)
            {
                DialogResult dialogResult = MessageBox.Show("You are about to change the calibration file for this microscope. " +
                                                            "This change will be applied to all future files. Reverts can only be done manually." +
                                                            " \n Are you sure you want to change the calibration file?",
                                                             "Update Calibration file ?",
                                                             MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    updateXMLCalibrationFile();
                    calibrationValueshanged = false;
                    label3.Text = "Calibration data updated to a new .xml calibration file...";
                }
            }
            else
            {
                label3.Text = "Nothing to update...";
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (this.calibrationValueshanged)
            {
                DialogResult dialogResult = MessageBox.Show("You made some changes to the Microscope calibrations but did not Update. " +
                                                            "\n Do you want to continue editing ?", 
                                                             "Continue Editing ?", 
                                                             MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    this.calibrationValueshanged = false;
                }
                else
                {
                    return;
                }
            }
            Close();            
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            calibrationValueshanged = true;
            label3.Text = "Objective/Relay values changed...";
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            calibrationValueshanged = true;
            label3.Text = "MicroscopeInfo values changed...";
        }
        private void LoadXMLCalibrationFile()
        {
            // Assumes a set xml file layout
            ds = new DataSet();
            ds.ReadXml(FileIO.getCurrentCalibrationXML());
            dataGridView1.DataSource = ds.Tables[1];
            textBox1.Text = ds.Tables[0].Rows[0]["MicroscopeInfo"].ToString();
            label3.Text = "Calibration data loaded from turretParameter.xml...";
        }
        private void updateXMLCalibrationFile()
        {
            ds.Tables[0].Rows[0]["MicroscopeInfo"] = textBox1.Text;
            FileIO.createNewTurretObjectiveRelayXML(ds); // Create new file...                         
            label3.Text = "Calibration data updated and saved to testytest.xml...";
        }
    }
}