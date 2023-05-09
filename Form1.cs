using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;

namespace pikpo_kp
{
    public partial class Form1 : Form
    {
        User user = new User();
        IDB db=new DB();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns[0].Visible = false;
        }
        private void mainwindowUI()
        {
            panel5.Visible = false;
            panel6.Visible = false;
            panel3.Visible = true;
            label3.Visible = false;
            label6.Visible = true;
            radioButton1.Visible = true;
            radioButton2.Visible = true;
            radioButton3.Visible = true;
            radioButton4.Visible = true;
            textBox1.Visible = false;
            textBox2.Visible = false;
            label4.Visible = false;
            label1.Visible = false;
            label2.Visible = false;
            button1.Visible = false;
            button7.Visible = true;
            dataGridView1.Visible = false;
            panel2.Visible = true;
            label10.Visible = false;
        }
        private void workwindowUI()
        {
            label10.Visible = true;
            panel2.Visible = false;
            panel3.Visible = true;
            panel6.Visible = true;
            dataGridView1.Visible = true;
            panel5.Visible = true;
            label5.Text = dataGridView1.Columns[1].HeaderText;
            label8.Text = dataGridView1.Columns[2].HeaderText;
            label9.Text = dataGridView1.Columns[3].HeaderText;
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox1.Text = "";
            comboBox2.Text = "";
            numericUpDown1.Value = 0;
            dataGridView1.ClearSelection();
            textBox3.Clear();
            db.GetValues(db.table,label5.Text,comboBox1);
            db.GetValues(db.table, label8.Text, comboBox2); 
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (db.AuthIn(textBox1.Text, textBox2.Text,ref user))
                mainwindowUI();
            else
                label4.Visible = true;
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            textBox3.Clear();
            textBox3.Text += "///////////////////////////////" + Environment.NewLine;
            try
            {
                db.SeeTable(db.table, dataGridView1);
                workwindowUI();
                textBox3.Text += "Выбранная таблица успешно открыта!" + Environment.NewLine;
            }
            catch(System.Exception ex)
            {
                textBox3.Text += "Невозможно открыть выбранную таблицу!"+Environment.NewLine;
            }
            textBox3.SelectionStart = textBox3.TextLength;
            textBox3.ScrollToCaret();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            db.table = radioButton1.Text;
            label10.Text="Таблица: "+ radioButton1.Text;
           
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            db.table = radioButton2.Text;
            label10.Text = "Таблица: " + radioButton2.Text;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            db.table = radioButton3.Text;
            label10.Text = "Таблица: " + radioButton3.Text;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            db.table = radioButton4.Text;
            label10.Text = "Таблица: " + radioButton4.Text;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            mainwindowUI();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox3.Text += "////////////////////////////" + Environment.NewLine;
            try
            {
                db.ChangeInTable(db.table, dataGridView1.Rows[dataGridView1.SelectedRows[0].Index].Cells[0].Value.ToString(),
                    comboBox1.Text, comboBox2.Text, numericUpDown1.Value.ToString(), label5.Text, label8.Text, label9.Text);
                db.GetValues(db.table, label5.Text, comboBox1);
                db.GetValues(db.table, label8.Text, comboBox2);
                textBox3.Text += "Строка успешно изменена!" + Environment.NewLine;
                db.SeeTable(db.table, dataGridView1);
                
            }
            catch (System.Exception ex)
            {
                textBox3.Text += "Ошибка, некорректные значения!" + Environment.NewLine;
            }
            textBox3.SelectionStart = textBox3.TextLength;
            textBox3.ScrollToCaret();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.ClearSelection();
                dataGridView1.Rows[db.CheckInTable(db.table, comboBox1.Text, comboBox2.Text, numericUpDown1.Value.ToString())].Selected = true;
            }
            catch (System.Exception ex)
            {
                textBox3.Text += "////////////////////////////" + Environment.NewLine;
                textBox3.Text += "Ошибка, в таблице нет строки с такими значениями!" + Environment.NewLine;
                textBox3.SelectionStart = textBox3.TextLength;
                textBox3.ScrollToCaret();
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                comboBox1.Text = dataGridView1.Rows[dataGridView1.SelectedRows[0].Index].Cells[1].Value.ToString();
                    comboBox2.Text = dataGridView1.Rows[dataGridView1.SelectedRows[0].Index].Cells[2].Value.ToString();
                    textBox3.Text += "////////////////////////////" + Environment.NewLine;
                    numericUpDown1.Value = Convert.ToInt32(dataGridView1.Rows[dataGridView1.SelectedRows[0].Index].Cells[3].Value.ToString());
                    textBox3.Text += "Строка выделена!" + Environment.NewLine;
                    textBox3.SelectionStart = textBox3.TextLength;
                    textBox3.ScrollToCaret();
                button3.Enabled = true;
                button4.Enabled = true;

            }
            catch (System.Exception ex)
            {
                button3.Enabled = false;
                button4.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                db.DeleteInTable(db.table, dataGridView1.Rows[dataGridView1.SelectedRows[0].Index].Cells[0].Value.ToString());
                textBox3.Text += "////////////////////////////" + Environment.NewLine;
                textBox3.Text += "Выделенная строка успешно удалена!" + Environment.NewLine;
                db.SeeTable(db.table,dataGridView1);
                textBox3.SelectionStart = textBox3.TextLength;
                db.GetValues(db.table, label5.Text,comboBox1);
                db.GetValues(db.table, label8.Text, comboBox2);
                textBox3.ScrollToCaret();
            }
            catch (System.Exception ex)
            {
                textBox3.Text += "////////////////////////////" + Environment.NewLine;
                textBox3.Text += "Ошибка, отсутсвует выделенная строка, которую можно было бы удалить!" + Environment.NewLine;
                textBox3.SelectionStart = textBox3.TextLength;
                textBox3.ScrollToCaret();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox3.Text += "////////////////////////////" + Environment.NewLine;
            try
            {
                db.PutInTable(db.table, comboBox1.Text, comboBox2.Text, numericUpDown1.Value.ToString(), label5.Text, label8.Text, label9.Text);
                textBox3.Text += "Cтрока успешно добавлена!" + Environment.NewLine;
                db.SeeTable(db.table, dataGridView1);
            }
            catch (System.Exception ex)
            {
                textBox3.Text += "Ошибка, некорректные значения!" + Environment.NewLine;
            }
            textBox3.SelectionStart = textBox3.TextLength;
            db.GetValues(db.table, label5.Text, comboBox1);
            db.GetValues(db.table, label8.Text, comboBox2);
            textBox3.ScrollToCaret();
        }
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            if (radioButton1.Checked) { comboBox1.Text = ""; }
        }

        private void comboBox2_TextUpdate(object sender, EventArgs e)
        {
            if (radioButton1.Checked) { comboBox2.Text = ""; }
        }
    }
}
class User
{
    public int Premissions { get; set; }
}
interface IDB
{
    public string table { get; set; }
    void SeeTable(string tablename,System.Windows.Forms.DataGridView DTF);
    void DeleteInTable(string tablename,string s1);
    void PutInTable(string tablename,string s1,string s2,string s3, string s4, string s5, string s6);
    void ChangeInTable(string tablename, string id, string s1, string s2, string s3, string s4, string s5, string s6);
    void GetFromTable(string tablename);
    int CheckInTable(string tablename, string s1, string s2,string s3);
    bool AuthIn(string login, string password,ref User user);
    void GetValues(string tablename,string columname,System.Windows.Forms.ComboBox cb);
   
}
class DB : IDB
{
    private string connectionstring = "Data Source=C:\\Users\\azamat\\Desktop\\pikpo\\pikpo.db;Cache=Shared;Mode=ReadWrite;";
    public string table { get; set; }
    public void SeeTable(string tablename, System.Windows.Forms.DataGridView DTF) 
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            SQLiteCommand cmd = new SQLiteCommand("select * from "+tablename, connection);
            List<string[]> data = new List<string[]>();
           
                DTF.Rows.Clear();
                SQLiteDataReader dr = cmd.ExecuteReader();
                DTF.Columns[0].HeaderText= dr.GetName(0);
                DTF.Columns[1].HeaderText = dr.GetName(1);
                DTF.Columns[2].HeaderText = dr.GetName(2);
                DTF.Columns[3].HeaderText = dr.GetName(3);
                while (dr.Read())
                {
                    data.Add(new string[4]);
                    data[data.Count - 1][0] = dr[0].ToString();
                    data[data.Count - 1][1] = dr[1].ToString();
                    data[data.Count - 1][2] = dr[2].ToString();
                    data[data.Count - 1][3] = dr[3].ToString();
                }
            foreach (string[] s in data)
            {
                DTF.Rows.Add(s);
               
            }
            
        }
       
    }
    public void DeleteInTable(string tablename, string s1)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            string comtext = $"DELETE from '{tablename}' WHERE id = {Convert.ToInt32(s1)}";
            SQLiteCommand cmd = new SQLiteCommand(comtext,connection);
            int dr = cmd.ExecuteNonQuery();
           
        }
    }
    public void PutInTable(string tablename,string s1,string s2,string s3,string s4, string s5, string s6)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            if (s1 == "" || s2 == "" || s3 == "")
                connection.Close();
            SQLiteCommand cmd = new SQLiteCommand("INSERT INTO " + tablename + " (" + s4 + "," + s5 + "," + s6 + ")" +
                " VALUES (:s1,:s2,:s3)", connection);
            cmd.Parameters.AddWithValue("s1", s1);
            cmd.Parameters.AddWithValue("s2", s2);
            cmd.Parameters.AddWithValue("s3", Convert.ToInt32(s3));
            int dr = cmd.ExecuteNonQuery();
        }
    }
    public void ChangeInTable(string tablename,string id,string s1,string s2,string s3,string s4, string s5,string s6)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            if(s1=="" || s2=="" || s3=="")
                connection.Close();
            string comtext = "", comtext1 = "", comtext2 = "";
            comtext = $"UPDATE '{table}' SET '{s4}' = '{s1}' WHERE id={Convert.ToInt32(id)}";
            comtext1 = $"UPDATE '{table}' SET '{s5}' = '{s2}' WHERE id={Convert.ToInt32(id)}";
            comtext2 = $"UPDATE '{table}' SET '{s6}' = {Convert.ToInt32(s3)} WHERE id={Convert.ToInt32(id)}";
            using (var command = new SQLiteCommand(comtext, connection))
            {
                command.ExecuteNonQuery();
            }
            using (var command = new SQLiteCommand(comtext1, connection))
            {
                command.ExecuteNonQuery();
            }
            using (var command = new SQLiteCommand(comtext2, connection))
            {
                command.ExecuteNonQuery();
            }
        }
    }

    public void GetFromTable(string tablename)
    {

    }
    public int CheckInTable(string tablename,string s1,string s2,string s3)
    {
        int index = -1,count=0;
        if (s1 == "" || s2 == "" || s3 == "")
        {
            return index;
        }
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            SQLiteCommand cmd = new SQLiteCommand("select * from "+tablename, connection);
            SQLiteDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (s1 == dr[1].ToString() && s2 == dr[2].ToString() && s3==dr[3].ToString() && index==-1)
                {
                    index = count;
                }
                count++;
            }
        }
        return index;
    }
    public bool AuthIn(string login,string password, ref User user)
    {
            using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
            {
                if (connection.State == ConnectionState.Open) { connection.Close(); }
                connection.Open();
                DataTable dt = new DataTable();
                SQLiteCommand cmd = new SQLiteCommand("select * from Пользователи", connection);
                SQLiteDataReader dr = cmd.ExecuteReader();
                int flag = 0;
                while (dr.Read())
                {
                    if (login == dr[1].ToString() && (password == dr[2].ToString()) && flag == 0)
                    {
                        flag = 1;
                        user.Premissions = Convert.ToInt32(dr[3]);
                    }
                }
                if (flag == 1)
                    return true;
            }
        
        return false;
    }
   public void GetValues(string tablename, string columname, System.Windows.Forms.ComboBox cb)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            cb.Items.Clear();
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            if (columname == "articule")
                tablename = "Товары";
            if(columname=="phone")
                tablename = "Клиенты";
            SQLiteCommand cmd = new SQLiteCommand("select * from "+tablename, connection);
            SQLiteDataReader dr = cmd.ExecuteReader();
            int count = 0;
            for(; count < 4; count++)
            {
                if (dr.GetName(count) == columname)
                {
                    break;
                }
                
            }
            while (dr.Read())
            {
                cb.Items.Add(dr[count].ToString());
            }
        }

    }

}


    
