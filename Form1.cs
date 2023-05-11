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
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


namespace pikpo_kp
{
    public partial class Form1 : Form
    {
        User user = new User();
        IDB db=new DB();
        Datas dbdata = new Datas();
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
            panel4.Visible = false;
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
           
            textBox3.Clear();
            
            
            db.GetValues(db.table,label5.Text,comboBox1);
            db.GetValues(db.table, label8.Text, comboBox2);
            db.GetValues(db.table, label5.Text, comboBox3);
            db.GetValues(db.table, label8.Text, comboBox4);
            panel4.Visible = true;
            checkBox1.Text = label5.Text;
            checkBox2.Text = label8.Text;
            checkBox3.Text = label9.Text;
            
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
            SetData();
            try
            {

                UpdateDGV();
               
                workwindowUI();
                textBox3.Text += "Выбранная таблица успешно открыта!" + Environment.NewLine;
            }
            catch (System.Exception ex)
            {
                textBox3.Text += "Невозможно открыть выбранную таблицу!" + Environment.NewLine;
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
           
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            db.table = radioButton3.Text;
            label10.Text = "Таблица: " + radioButton3.Text;
           
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            db.table = radioButton4.Text;
            label10.Text = "Таблица: " + radioButton4.Text;
           
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
                SetData();
                db.ChangeInTable(dbdata);
                db.GetValues(db.table, label5.Text, comboBox1);
                db.GetValues(db.table, label8.Text, comboBox2);
                textBox3.Text += "Строка успешно изменена!" + Environment.NewLine;
                button9_Click(sender, e);

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
                SetData();
                dataGridView1.Rows[db.CheckInTable(dbdata)].Selected = true;
                SetData();
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
                SetData();
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
                SetData();
                db.DeleteInTable(dbdata);
                textBox3.Text += "////////////////////////////" + Environment.NewLine;
                textBox3.Text += "Выделенная строка успешно удалена!" + Environment.NewLine;
                button9_Click(sender, e);
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
                SetData();
                db.PutInTable(dbdata);
                textBox3.Text += "Cтрока успешно добавлена!" + Environment.NewLine;
                button9_Click(sender, e);

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

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox3.Text += "////////////////////////////" + Environment.NewLine;
            textBox3.Text += "Таблица обновлена!" + Environment.NewLine;
            UpdateDGV();
            textBox3.SelectionStart = textBox3.TextLength;
            textBox3.ScrollToCaret();
        }
        private void UpdateDGV()
        {
            SetData();
            db.SeeTable(dbdata, dataGridView1);
            
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }
        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

                //NASSIM LOUCHANI
            }
        }

        private void SetData()
        {
            try
            {
                dbdata.id = dataGridView1.Rows[dataGridView1.SelectedRows[0].Index].Cells[0].Value.ToString();
            }
            catch(System.Exception ex)
            {
                dbdata.id = "-1";
            }
           
            dbdata.tablename = db.table;
            dbdata.s1 = comboBox1.Text.ToString();
            dbdata.s2 = comboBox2.Text.ToString();
            dbdata.s3 = numericUpDown1.Value.ToString();
            dbdata.s4 = label5.Text;
            dbdata.s5 = label8.Text;
            dbdata.s6 = label9.Text;
            try
            {
                db.GetValues(db.table, label5.Text, comboBox1);
                db.GetValues(db.table, label8.Text, comboBox2);
                db.GetValues(db.table, label5.Text, comboBox3);
                db.GetValues(db.table, label8.Text, comboBox4);
            }
            catch (System.Exception Ex)
            {

            }
        }
    }
}

class User
{
    public int Premissions { get; set; }
}
struct Datas
{
    public string tablename { get; set; }
    public string id { get; set; }
    public string s1 { get; set; }
    public string s2 { get; set; }
    public string s3 { get; set; }
    public string s4 { get; set; }
    public string s5 { get; set; }
    public string s6 { get; set; }
}
interface IDB
{
    public string table { get; set; }
    void SeeTable(Datas dbdata, System.Windows.Forms.DataGridView DTF);
    void DeleteInTable(Datas dbdata);
    void PutInTable(Datas dbdata);
    void ChangeInTable(Datas dbdata);
    void GetFromTable(Datas dbdata);
    int CheckInTable(Datas dbdata);
    bool AuthIn(string login,string password, ref User user);
    void GetValues(string tablename,string columname,System.Windows.Forms.ComboBox cb);
   

}
class DB : IDB
{
    private string connectionstring = "Data Source=C:\\Users\\azamat\\Desktop\\pikpo\\pikpo.db;Cache=Shared;Mode=ReadWrite;";
    public string table { get; set; }

    public void SeeTable(Datas dbdata, System.Windows.Forms.DataGridView DTF)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            SQLiteCommand cmd = new SQLiteCommand("select * from " + dbdata.tablename, connection);
            List<string[]> data = new List<string[]>();

            DTF.Rows.Clear();
            SQLiteDataReader dr = cmd.ExecuteReader();
            DTF.Columns[0].HeaderText = dr.GetName(0);
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
    public void DeleteInTable(Datas dbdata)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            string comtext = $"DELETE from '{dbdata.tablename}' WHERE id = {Convert.ToInt32(dbdata.id)}";
            SQLiteCommand cmd = new SQLiteCommand(comtext,connection);
            int dr = cmd.ExecuteNonQuery();
           
        }
    }
    public void PutInTable(Datas dbdata)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            if (dbdata.s1 == "" || dbdata.s2 == "" || dbdata.s3 == "")
                connection.Close();
            SQLiteCommand cmd = new SQLiteCommand("INSERT INTO " + dbdata.tablename + " (" + dbdata.s4 + "," + dbdata.s5 + "," + dbdata.s6 + ")" +
                " VALUES (:s1,:s2,:s3)", connection);
            cmd.Parameters.AddWithValue("s1", dbdata.s1);
            cmd.Parameters.AddWithValue("s2", dbdata.s2);
            cmd.Parameters.AddWithValue("s3", Convert.ToInt32(dbdata.s3));
            int dr = cmd.ExecuteNonQuery();
        }
    }
    public void ChangeInTable(Datas dbdata)
    {
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            if(dbdata.s1 =="" || dbdata.s2 =="" || dbdata.s3 =="")
                connection.Close();
            string comtext = "", comtext1 = "", comtext2 = "";
            comtext = $"UPDATE '{dbdata.tablename}' SET '{dbdata.s4}' = '{dbdata.s1}' WHERE id={Convert.ToInt32(dbdata.id)}";
            comtext1 = $"UPDATE '{table}' SET '{dbdata.s5}' = '{dbdata.s2}' WHERE id={Convert.ToInt32(dbdata.id)}";
            comtext2 = $"UPDATE '{table}' SET '{dbdata.s6}' = {Convert.ToInt32(dbdata.s3)} WHERE id={Convert.ToInt32(dbdata.id)}";
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

    public void GetFromTable(Datas dbdata)
    {

    }
    public int CheckInTable(Datas dbdata)
    {
        int index = -1,count=0;
        if (dbdata.s1 == "" || dbdata.s2 == "" || dbdata.s3 == "")
        {
            return index;
        }
        using (SQLiteConnection connection = new SQLiteConnection(connectionstring))
        {
            if (connection.State == ConnectionState.Open) { connection.Close(); }
            connection.Open();
            DataTable dt = new DataTable();
            SQLiteCommand cmd = new SQLiteCommand("select * from "+ dbdata.tablename, connection);
            SQLiteDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (dbdata.s1 == dr[1].ToString() && dbdata.s2 == dr[2].ToString() && dbdata.s3 ==dr[3].ToString() && index==-1)
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


    
