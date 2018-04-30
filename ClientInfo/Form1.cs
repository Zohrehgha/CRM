using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClientInfo
{
    public partial class Form1 : Form
    {
        ADODB.Connection Con;
        ADODB.Recordset Rs;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'clientDBDataSet.ClientTable' table. You can move, or remove it, as needed.
            this.clientTableTableAdapter.Fill(this.clientDBDataSet.ClientTable);
            Con = new ADODB.Connection();
            Rs = new ADODB.Recordset();
            Con.Provider = "Microsoft.jet.oledb.4.0";
            Con.ConnectionString = "C:\\ITD\\Term 3\\C#\\assignment\\assignment6\\ClientDB.mdb";
            Con.Open();
            Rs.Open("Select * from ClientTable", Con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic);
            Rs.MoveFirst();
            dataGridView1.DataSource = Rs.Fields[2];
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        //showing previous record in db
        private void button7_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is Empty!");
                return;
            }
            Rs.MovePrevious();
            if (Rs.BOF == true)
            {
                Rs.MoveFirst();
                MessageBox.Show("Passed Beginning of File");
            }
            ShowDataOnForm();
        }
        //showing next record in db
        private void button8_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is Empty!");
                return;
            }
            Rs.MoveNext();
            if (Rs.EOF == true)
            {
                Rs.MoveLast();
                MessageBox.Show("Passed End of File");
            }
            ShowDataOnForm();

        }
        //showing First record in db
        private void button6_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is Empty!");
                return;
            }
            Rs.MoveFirst();
            ShowDataOnForm();
        }
        public void ShowDataOnForm()
        {
            textBox1.Text = Convert.ToString(Rs.Fields["ContactID"].Value);
            textBox2.Text = Convert.ToString(Rs.Fields["CompanyName"].Value);
            textBox3.Text = Convert.ToString(Rs.Fields["Name"].Value);
            textBox4.Text = Convert.ToString(Rs.Fields["LastName"].Value);
            textBox5.Text = Convert.ToString(Rs.Fields["MiddleName"].Value);
            textBox6.Text = Convert.ToString(Rs.Fields["PhoneNumber"].Value);
            textBox7.Text = Convert.ToString(Rs.Fields["CellPhone"].Value);
            textBox8.Text = Convert.ToString(Rs.Fields["AlternativePhone"].Value);
            textBox9.Text = Convert.ToString(Rs.Fields["Email"].Value);
            textBox10.Text = Convert.ToString(Rs.Fields["Email2"].Value);
            textBox11.Text = Convert.ToString(Rs.Fields["Industry"].Value);
            textBox12.Text = Convert.ToString(Rs.Fields["SuiteNumber"].Value);
            textBox13.Text = Convert.ToString(Rs.Fields["StreetNumber"].Value);
            textBox14.Text = Convert.ToString(Rs.Fields["StreetName"].Value);
            textBox15.Text = Convert.ToString(Rs.Fields["City"].Value);
            textBox16.Text = Convert.ToString(Rs.Fields["Province"].Value);
            textBox17.Text = Convert.ToString(Rs.Fields["Country"].Value);
            textBox18.Text = Convert.ToString(Rs.Fields["PostalCode"].Value);
        }
        //showing last record in db
        private void button9_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is Empty!");
                return;
            }
            Rs.MoveLast();
            ShowDataOnForm();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is Empty!");
                return;
            }
            Rs.MoveNext();
            if (Rs.EOF == true)
            {
                Rs.MoveLast();
                MessageBox.Show("Passed End of File");
            }
            ShowDataOnForm();
        }
        
        private void button3_Click_1(object sender, EventArgs e)
        {
            ClearBoxes();
        }
        //function to clear the form
        public void ClearBoxes()
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            textBox13.Clear();
            textBox14.Clear();
            textBox15.Clear();
            textBox16.Clear();
            textBox17.Clear();
            textBox18.Clear();
        }
        //Adding new record into DB
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" ||
                textBox2.Text == "" ||
                textBox3.Text == "" ||
                textBox4.Text == "" ||
                textBox5.Text == "" ||
                textBox6.Text == "" ||
                textBox7.Text == "" ||
                textBox8.Text == "" ||
                textBox9.Text == "" ||
                textBox10.Text == "" ||
                textBox11.Text == "" ||
                textBox12.Text == "" ||
                textBox13.Text == "" ||
                textBox14.Text == "" ||
                textBox15.Text == "" ||
                textBox16.Text == "" ||
                textBox17.Text == "" ||
                textBox18.Text == "")
            {
                MessageBox.Show("Please Fill up all boxes");
                return;
            }
            String Criteria;
            Criteria = "ContactID =" + textBox1.Text;
            Rs.MoveFirst();
            //go to the beginning to start serach 
            Rs.Find(Criteria);
            //Either We find the record(s), which is the first record if there are more than one
            //If record is found the file pointer stays at it
            //if not found, the file pointer has passed eof meaning eof = true
            if (Rs.EOF == true)
            {
                //not found
                Rs.AddNew();
                SaveinTable();
                Rs.Update();
                MessageBox.Show("Record Added succesfully");
                ClearBoxes();
                return;
            }
            else
            {
                //found 
                MessageBox.Show("Duplicate Record, try another ISBN");
                return;
            }
        }
        private void SaveinTable()
        {
            Rs.Fields["ContactID"].Value = textBox1.Text;
            Rs.Fields["CompanyName"].Value = textBox2.Text;
            Rs.Fields["Name"].Value = textBox3.Text;
            Rs.Fields["LastName"].Value = textBox4.Text;
            Rs.Fields["MiddleName"].Value = textBox5.Text;
            Rs.Fields["PhoneNumber"].Value = textBox6.Text;
            Rs.Fields["CellPhone"].Value = textBox7.Text;
            Rs.Fields["AlternativePhone"].Value = textBox8.Text;
            Rs.Fields["Email"].Value = textBox9.Text;
            Rs.Fields["Email2"].Value = textBox10.Text;
            Rs.Fields["Industry"].Value = textBox11.Text;
            Rs.Fields["SuiteNumber"].Value = textBox12.Text;
            Rs.Fields["StreetNumber"].Value = textBox13.Text;
            Rs.Fields["StreetName"].Value = textBox14.Text;
            Rs.Fields["City"].Value = textBox15.Text;
            Rs.Fields["Province"].Value = textBox16.Text;
            Rs.Fields["Country"].Value = textBox17.Text;
            Rs.Fields["PostalCode"].Value = textBox18.Text;
        }
        //Modifying a existing record into DB
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" ||
                textBox2.Text == "" ||
                textBox3.Text == "" ||
                textBox4.Text == "" ||
                textBox5.Text == "" ||
                textBox6.Text == "" ||
                textBox7.Text == "" ||
                textBox8.Text == "" ||
                textBox9.Text == "" ||
                textBox10.Text == "" ||
                textBox11.Text == "" ||
                textBox12.Text == "" ||
                textBox13.Text == "" ||
                textBox14.Text == "" ||
                textBox15.Text == "" ||
                textBox16.Text == "" ||
                textBox17.Text == "" ||
                textBox18.Text == "")
            {
                MessageBox.Show("Please Fill up all boxes");
                return;
            }
            String Criteria;
            Criteria = "ContactID =" + textBox1.Text;
            Rs.MoveFirst();
            //go to the beginning to start serach 
            Rs.Find(Criteria);
            //Either We find the record(s), which is the first record if there are more than one
            // If record is found the file pointer stays at it
            //if not found, the file pointer has passed eof meaning eof = true
            if (Rs.EOF)
            {
                // it is impossible, if you refrain from changing the ID 
                MessageBox.Show("Record with this ContactID does not exist");
                return;
            }
            else
            {
                //found 
                SaveinTable();
                Rs.Update();
                MessageBox.Show("Record Modified succesfully");
            }
        }
        //searching one record into DB by Criteria
        private void button5_Click(object sender, EventArgs e)
        {
            String Criteria;
            Criteria = "";
            if (textBox1.Text != "")
            {
                Criteria = Criteria + "ContactID = " + textBox1.Text;
            }
            if (textBox2.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND CompanyName = '" + textBox2.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "CompanyName = '" + textBox2.Text + "'";
                }
            }
            if (textBox3.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND Name = '" + textBox3.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "Name = '" + textBox3.Text + "'";
                }
            }
            if (textBox4.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND LastName = '" + textBox4.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "LastName = '" + textBox4.Text + "'";
                }
            }
            if (textBox5.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND MiddleName = '" + textBox5.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "MiddleName = '" + textBox5.Text + "'";
                }
            }
            if (textBox6.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND PhoneNumber = '" + textBox6.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "PhoneNumber = '" + textBox6.Text + "'";
                }
            }
            if (textBox7.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND CellPhone = '" + textBox7.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "CellPhone = '" + textBox7.Text + "'";
                }
            }
            if (textBox8.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND AlternativePhone = '" + textBox8.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "AlternativePhone = '" + textBox8.Text + "'";
                }
            }
            if (textBox9.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND Email = '" + textBox9.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "Email = '" + textBox9.Text + "'";
                }
            }
            if (textBox10.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND Email2 = '" + textBox10.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "Email2 = '" + textBox10.Text + "'";
                }
            }
            if (textBox11.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND Industry = '" + textBox11.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "Industry = '" + textBox11.Text + "'";
                }
            }
            if (textBox12.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND SuiteNumber = '" + textBox12.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "SuiteNumber = '" + textBox12.Text + "'";
                }
            }
            if (textBox13.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND StreetNumber = '" + textBox13.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "StreetNumber = '" + textBox13.Text + "'";
                }
            }
            if (textBox14.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND StreetName = '" + textBox14.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "StreetName = '" + textBox14.Text + "'";
                }
            }
            if (textBox15.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND City = '" + textBox15.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "City = '" + textBox15.Text + "'";
                }
            }
            if (textBox16.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND Province = '" + textBox16.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "Province = '" + textBox16.Text + "'";
                }
            }
            if (textBox17.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND Country = '" + textBox17.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "Country = '" + textBox17.Text + "'";
                }
            }
            if (textBox18.Text != "")
            {
                if (Criteria != "")
                {
                    Criteria = Criteria + " AND PostalCode = '" + textBox18.Text + "'";
                }
                else
                {
                    Criteria = Criteria + "PostalCode = '" + textBox18.Text + "'";
                }
            }

            Rs.MoveFirst();
            Rs.Filter = Criteria;
            if (Rs.EOF == true)
            {
                //not found
                MessageBox.Show("Recod with your specific criteria not found");
                return;
            }
            else
            {
                ShowDataOnForm();
                Rs.Filter = "";
            }
        }
        //Deleting ane exsitind record into DB
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" ||
                textBox2.Text == "" ||
                textBox3.Text == "" ||
                textBox4.Text == "" ||
                textBox5.Text == "" ||
                textBox6.Text == "" ||
                textBox7.Text == "" ||
                textBox8.Text == "" ||
                textBox9.Text == "" ||
                textBox10.Text == "" ||
                textBox11.Text == "" ||
                textBox12.Text == "" ||
                textBox13.Text == "" ||
                textBox14.Text == "" ||
                textBox15.Text == "" ||
                textBox16.Text == "" ||
                textBox17.Text == "" ||
                textBox18.Text == "")
            {
                MessageBox.Show("Please Fill up all boxes");
                return;
            }
            String Criteria;
            Criteria = "ContactID =" + textBox1.Text;
            Rs.MoveFirst();
            //go to the beginning to start serach 
            Rs.Find(Criteria);
            //Either We find the record(s), which is the first record if there are more than one
            // If record is found the file pointer stays at it
            //if not found, the file pointer has passed eof meaning eof = true
            if (Rs.EOF)
            {
                // it is impossible, if you refrain from changing the ID 
                MessageBox.Show("Record with this ContactID does not exist");
                return;
            }
            else
            {
                //found 
                //confirm 
                DialogResult MsgbxResult;
                MsgbxResult = MessageBox.Show("Are you Sure?!", "Confirm Delete", MessageBoxButtons.YesNo);
                if (Convert.ToString(MsgbxResult) == "Yes")
                {
                    Rs.Delete();
                    Rs.Update();
                    MessageBox.Show("Record Deleted Successfully !!!");
                    ClearBoxes();

                }

            }
        }
    }
}
