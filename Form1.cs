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

namespace Equity_Checking
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private bool nonNumberEntered = false;
        private bool characterCorrected = false;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "XLS files (*.xls)|*.xls;*.xlsx;*.xlsm";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                dissableAll();
                TheClass.createapliXL();
                if (!listBox1.Items.Contains("File: " + Path.GetFileName(dlg.FileName)))
                {
                    string nombreOWB = TheClass.openingWB(dlg.FileName);
                    listBox1.Items.Add("File: " + nombreOWB);
                    TheClass.createnewWB(nombreOWB);
                    TheClass.addnewWBtoArrayLst();
                }
                changeFieldsbydefault();
                enableAll();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            dissableAll();
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            changeFieldsbydefault();
            listBox1.Items.Clear();
            TheClass.cerrartodo();
            TheClass.cleanWorkbooksArray();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text != "Source") && (textBox2.Text != "IDs") && (textBox3.Text != "Translations"))
            {
                if (TheClass.checklimitcolums(textBox1.Text, textBox2.Text, textBox3.Text.Split(new Char[] { ' ', ',', '-' }), button4.BackColor) == true)
                {
                    dissableAll();
                    if (button4.BackColor == Color.DarkRed)
                    {
                        TheClass.superswitch(int.Parse(textBox1.Text), int.Parse(textBox2.Text), textBox3.Text.Split(new Char[] { ' ', ',', '-' }), textBox4.Text, textBox5.Text, checkBox1.Checked, checkBox2.Checked);
                    }
                    else
                    {
                        string[] columnsLoc = new string[textBox3.Text.Split(new Char[] { ' ', ',', '-' }).Length];
                        for (int i = 0; i < columnsLoc.Length; i++)
                        {
                            if (textBox3.Text.Split(new Char[] { ' ', ',', '-' })[i] != "")
                            {
                                columnsLoc[i] = TheClass.GetIndexInAlphabet(textBox3.Text.Split(new Char[] { ' ', ',', '-' })[i][0]).ToString();
                            }
                        }
                        TheClass.superswitch(TheClass.GetIndexInAlphabet(textBox1.Text[0]), TheClass.GetIndexInAlphabet(textBox2.Text[0]), columnsLoc, textBox4.Text, textBox5.Text, checkBox1.Checked, checkBox2.Checked);
                    }
                    enableAll();
                    if (TheClass.thereareissues() == true)
                    {
                        SaveFileDialog svf = new SaveFileDialog();
                        DialogResult svfR;
                        svf.OverwritePrompt = false;

                        svf.Filter = "XLS files (*.xls)|*.xls";
                        do
                        {
                            svfR = svf.ShowDialog();
                        } while (listBox1.Items.Contains("File: " + Path.GetFileName(svf.FileName)));
                        if (svfR == DialogResult.OK)
                        {
                            dissableAll();
                            checkBox1.Enabled = false;
                            checkBox2.Enabled = false;
                            TheClass.cerraroWB();
                            try
                            {
                                TheClass.returnWBReport().SaveAs(svf.FileName);
                                MessageBox.Show("Export success", "Congratulations!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                textBox1.Text = "Nice";
                                textBox2.Text = "Work";
                                textBox3.Text = "=)";
                            }
                            catch
                            {
                                MessageBox.Show("Cannot save the Report in \n" + svf.FileName + "\n Try to close the Report file before saving on it", "Fail", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                                textBox1.Text = "Good";
                                textBox2.Text = "Try";
                                textBox3.Text = "=)";
                            }
                            listBox1.Items.Clear();
                            TheClass.cerrartodo();
                            TheClass.cleanWorkbooksArray();
                        }
                    }
                    else
                    {
                        dissableAll();
                        MessageBox.Show("There are not numeric issues!", "Fail", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        listBox1.Items.Clear();
                        TheClass.cerrartodo();
                        TheClass.cleanWorkbooksArray();
                        textBox1.Text = "Good";
                        textBox2.Text = "Try";
                        textBox3.Text = "=)";
                    }
                }
                else
                {
                    MessageBox.Show("Please make sure that you insert \nintegers as column values less than 37 and different than 0.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Please make sure that you fill \nin all the mandatory fields properly.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            enableAll();
            changeFieldsbydefault();
            if (button4.BackColor == Color.DarkRed)
            {
                button4.BackColor = Color.GreenYellow;
            }
            else
            {
                button4.BackColor = Color.DarkRed;
            }
        }

        private void enableAll()
        {
            if (textBox1.Text == "Source")
            {
                textBox1.ForeColor = Color.Gray;
            }
            else
            {
                textBox1.ForeColor = Color.Black;
            }
            if (textBox2.Text == "IDs")
            {
                textBox2.ForeColor = Color.Gray;
            }
            else
            {
                textBox2.ForeColor = Color.Black;
            }
            if (textBox3.Text == "Translations")
            {
                textBox3.ForeColor = Color.Gray;
            }
            else
            {
                textBox3.ForeColor = Color.Black;
            }
            if (textBox4.Text == "</")
            {
                textBox4.ForeColor = Color.Gray;
            }
            else
            {
                textBox4.ForeColor = Color.Black;
            }
            if (textBox5.Text == "/>")
            {
                textBox5.ForeColor = Color.Gray;
            }
            else
            {
                textBox5.ForeColor = Color.Black;
            }
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
        }
        private void changeFieldsbydefault()
        {
            textBox1.Text = "Source";
            textBox2.Text = "IDs";
            textBox3.Text = "Translations";
            textBox4.Text = "</";
            textBox5.Text = "/>";
        }
        private void dissableAll()
        {
            textBox1.ForeColor = SystemColors.GrayText;
            textBox2.ForeColor = SystemColors.GrayText;
            textBox3.ForeColor = SystemColors.GrayText;
            textBox4.ForeColor = SystemColors.GrayText;
            textBox5.ForeColor = SystemColors.GrayText;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }
        protected override void OnClosed(EventArgs e)
        {
            TheClass.cerrartodo();
        }

        private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            // Initialize the flag to false.
            nonNumberEntered = false;
            characterCorrected = false;

            if (button4.BackColor == Color.DarkRed)
            {
                // Determine whether the keystroke is a number from the top of the keyboard. 
                if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
                {
                    // Determine whether the keystroke is a number from the keypad. 
                    if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                    {
                        if (e.KeyCode != Keys.Back)
                        {
                            // A non-numerical keystroke was pressed. 
                            // Set the flag to true and evaluate in KeyPress event.
                            nonNumberEntered = true;
                        }
                    }
                }
            }
            else if (button4.BackColor == Color.GreenYellow)
            {
                if (e.KeyCode != Keys.Back)
                {
                    if ((e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9) || (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9) || textBox1.Text.Length != 0)
                    {
                        // Determine whether the keystroke is a number from the keypad. 
                        // A non-numerical keystroke was pressed. 
                        // Set the flag to true and evaluate in KeyPress event.
                        characterCorrected = true;
                    }
                    else if ((e.KeyCode < Keys.A) || (e.KeyCode > Keys.Z))
                    {
                        characterCorrected = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number. 
            if ((Control.ModifierKeys == Keys.Shift) || ModifierKeys == (Keys.Control | Keys.Alt))
            {
                nonNumberEntered = true;
            }
        }
        private void textBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            // Check for the flag being set in the KeyDown event. 
            if ((nonNumberEntered == true) || (characterCorrected == true))
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
        }
        private void textBox1_GotFocus(object sender, EventArgs e)
        {
            if (textBox1.Text == "Source")
            {
                textBox1.Text = "";
                textBox1.ForeColor = SystemColors.WindowText;
            }
        }
        private void textBox1_LostFocus(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.ForeColor = SystemColors.GrayText;
                textBox1.Text = "Source";

            }
        }
        private void textBox2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            nonNumberEntered = false;
            characterCorrected = false;

            if (button4.BackColor == Color.DarkRed)
            {
                // Determine whether the keystroke is a number from the top of the keyboard. 
                if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
                {
                    // Determine whether the keystroke is a number from the keypad. 
                    if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                    {
                        if (e.KeyCode != Keys.Back)
                        {
                            // A non-numerical keystroke was pressed. 
                            // Set the flag to true and evaluate in KeyPress event.
                            nonNumberEntered = true;
                        }
                    }
                }
            }

            else if (button4.BackColor == Color.GreenYellow)
            {
                if (e.KeyCode != Keys.Back)
                {
                    if ((e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9) || (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9) || textBox2.Text.Length != 0)
                    {
                        // Determine whether the keystroke is a number from the keypad. 
                        // A non-numerical keystroke was pressed. 
                        // Set the flag to true and evaluate in KeyPress event.
                        characterCorrected = true;
                    }
                    else if ((e.KeyCode < Keys.A) || (e.KeyCode > Keys.Z))
                    {
                        characterCorrected = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number. 
            if (Control.ModifierKeys == Keys.Shift || ModifierKeys == (Keys.Control | Keys.Alt))
            {
                nonNumberEntered = true;
            }
        }
        private void textBox2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            // Check for the flag being set in the KeyDown event. 
            if ((nonNumberEntered == true) || (characterCorrected == true))
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
        }
        private void textBox2_GotFocus(object sender, EventArgs e)
        {
            if (textBox2.Text == "IDs")
            {
                textBox2.Text = "";
                textBox2.ForeColor = SystemColors.WindowText;
            }
        }
        private void textBox2_LostFocus(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.ForeColor = SystemColors.GrayText;
                textBox2.Text = "IDs";

            }
        }
        private void textBox3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            nonNumberEntered = false;
            characterCorrected = false;

            if (button4.BackColor == Color.DarkRed)
            {
                // Determine whether the keystroke is a number from the top of the keyboard. 
                if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
                {
                    // Determine whether the keystroke is a number from the keypad. 
                    if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                    {
                        if (e.KeyCode != Keys.Back)
                        {
                            // A non-numerical keystroke was pressed. 
                            // Set the flag to true and evaluate in KeyPress event.
                            nonNumberEntered = true;
                        }
                    }
                }
            }
            else if (button4.BackColor == Color.GreenYellow)
            {
                if (e.KeyCode != Keys.Back)
                {
                    if ((e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9) || (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9))
                    {
                        // Determine whether the keystroke is a number from the keypad. 
                        // A non-numerical keystroke was pressed. 
                        // Set the flag to true and evaluate in KeyPress event.
                        characterCorrected = true;
                    }
                    if ((e.KeyCode >= Keys.A) && (e.KeyCode <= Keys.Z))
                    {
                        if (textBox3.Text == "")
                        {
                            characterCorrected = false;
                        }
                        else if ((textBox3.Text[textBox3.Text.Length - 1].ToString() == " ") || (textBox3.Text[textBox3.Text.Length - 1].ToString() == "-") || (textBox3.Text[textBox3.Text.Length - 1].ToString() == ","))
                        {
                            characterCorrected = false;
                        }
                        else
                        {
                            characterCorrected = true;
                        }
                    }
                    else
                    {
                        characterCorrected = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number. 
            if (Control.ModifierKeys == Keys.Shift || ModifierKeys == (Keys.Control | Keys.Alt))
            {
                nonNumberEntered = true;
            }
        }
        private void textBox3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            // Check for the flag being set in the KeyDown event. 
            if ((nonNumberEntered == true) || (characterCorrected == true))
            {
                if (e.KeyChar != ' ' && e.KeyChar != '-' && e.KeyChar != ',')
                {
                    // Stop the character from being entered into the control since it is non-numerical.
                    e.Handled = true;
                }
            }
        }
        private void textBox3_GotFocus(object sender, EventArgs e)
        {
            if (textBox3.Text == "Translations")
            {
                textBox3.Text = "";
                textBox3.ForeColor = SystemColors.WindowText;
            }
        }
        private void textBox3_LostFocus(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                textBox3.ForeColor = SystemColors.GrayText;
                textBox3.Text = "Translations";

            }
        }
        private void textBox4_GotFocus(object sender, EventArgs e)
        {
            if (textBox4.Text == "</")
            {
                textBox4.Text = "";
                textBox4.ForeColor = SystemColors.WindowText;
            }
        }
        private void textBox4_LostFocus(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                textBox4.ForeColor = SystemColors.GrayText;
                textBox4.Text = "</";

            }
        }
        private void textBox5_GotFocus(object sender, EventArgs e)
        {
            if (textBox5.Text == "/>")
            {
                textBox5.Text = "";
                textBox5.ForeColor = SystemColors.WindowText;
            }
        }
        private void textBox5_LostFocus(object sender, EventArgs e)
        {
            if (textBox5.Text == "")
            {
                textBox5.ForeColor = SystemColors.GrayText;
                textBox5.Text = "/>";

            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if ((!checkBox2.Checked) && (!checkBox1.Checked))
            {
                dissableAll();
            }
            else
            {
                enableAll();
            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if ((!checkBox2.Checked) && (!checkBox1.Checked))
            {
                dissableAll();
            }
            else
            {
                enableAll();
            }
        }
    }
}
