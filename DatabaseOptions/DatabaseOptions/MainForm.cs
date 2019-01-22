using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework;

namespace QuestInserter
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        OleDbConnection connect = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=../../../../WhoWantsToBeAEnglishMaster/bin/Debug/Content/Database/Quests.accdb");
        OleDbCommand query = new OleDbCommand();
        OleDbDataReader reader;
        public string Quest, OptionA, OptionB, OptionC, OptionD, AnswerOfQuest, path;
        public int QuestID, QuestMode, NumberOfSelected;
        private void Connect()
        {
            if (connect.State == ConnectionState.Open)
            {
                connect.Close();
                connect.Open();
            }
            else if (connect.State == ConnectionState.Closed)
            {
                connect.Open();
            }
        }
        private void Disconnect()
        {
            if (connect.State == ConnectionState.Open)
            {
                connect.Close();
            }
        }
        private void Clear()
        {
            QuestListView.Items.Clear();
        }
        private void List()
        {
            try
            {
                reader = query.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        ListViewItem Item = new ListViewItem(reader["questID"].ToString());
                        Item.SubItems.Add(reader["quest"].ToString());
                        Item.SubItems.Add(reader["optionA"].ToString());
                        Item.SubItems.Add(reader["optionB"].ToString());
                        Item.SubItems.Add(reader["optionC"].ToString());
                        Item.SubItems.Add(reader["optionD"].ToString());
                        Item.SubItems.Add(reader["questAnswer"].ToString());
                        QuestListView.Items.Add(Item);
                    }
                }
                NumberOfQuestsLabel.Text = QuestListView.Items.Count.ToString() + " Quest(s)";
            }
            catch (Exception Error)
            {
                MessageBox.Show("Error =>" + Error.Message.ToString());
            }
        }
        public MainForm() { InitializeComponent(); }
        private void MainForm_Load(object sender, EventArgs e)
        {
            Clear();
            Connect();
            query = new OleDbCommand("SELECT * FROM Words ORDER BY questID", connect);
            List();
            Disconnect();

            Mode1Button.Enabled = false;
        }
        private void Mode1Button_Click(object sender, EventArgs e)
        {
            path = Application.StartupPath.ToString() + "\\Content/Images/blue_plus.png";
            AddQuestButton.BackgroundImage = Image.FromFile(@path);

            Clear();
            Connect();
            query = new OleDbCommand("SELECT * FROM Words ORDER BY questID", connect);
            List();
            Disconnect();

            Mode1Button.Enabled = false;
            Mode2Button.Enabled = true;

            Mode1Button.Style = MetroColorStyle.Blue;
            Mode2Button.Style = MetroColorStyle.Blue;
            RefreshButton.Style = MetroColorStyle.Blue;
            DeleteButton.Style = MetroColorStyle.Blue;
            
            AddButton.Style = MetroColorStyle.Blue;
            QuestTextBox.Style = MetroColorStyle.Blue;
            QuestListView.Style = MetroColorStyle.Blue;
            OptionATextBox.Style = MetroColorStyle.Blue;
            OptionBTextBox.Style = MetroColorStyle.Blue;
            OptionCTextBox.Style = MetroColorStyle.Blue;
            OptionDTextBox.Style = MetroColorStyle.Blue;
            QuestModeComboBox.Style = MetroColorStyle.Blue;
            QuestAnswerComboBox.Style = MetroColorStyle.Blue;

            EditQuestButton.Style = MetroColorStyle.Blue;
            EditButton.Style = MetroColorStyle.Blue;
            QuestEditTextBox.Style = MetroColorStyle.Blue;
            OptionAEditTextBox.Style = MetroColorStyle.Blue;
            OptionBEditTextBox.Style = MetroColorStyle.Blue;
            OptionCEditTextBox.Style = MetroColorStyle.Blue;
            OptionDEditTextBox.Style = MetroColorStyle.Blue;
            QuestModeEditComboBox.Style = MetroColorStyle.Blue;
            QuestAnswerEditComboBox.Style = MetroColorStyle.Blue;

            this.Style = MetroColorStyle.Blue;
            this.Hide();
            this.Show();
        }
        private void Mode2Button_Click(object sender, EventArgs e)
        {
            path = Application.StartupPath.ToString() + "\\Content/Images/red_plus.png";
            AddQuestButton.BackgroundImage = Image.FromFile(@path);

            Clear();
            Connect();
            query = new OleDbCommand("SELECT * FROM Expressions ORDER BY questID", connect);
            List();
            Disconnect();

            Mode2Button.Enabled = false;
            Mode1Button.Enabled = true;

            Mode1Button.Style = MetroColorStyle.Red;
            Mode2Button.Style = MetroColorStyle.Red;
            RefreshButton.Style = MetroColorStyle.Red;
            DeleteButton.Style = MetroColorStyle.Red;
            
            AddButton.Style = MetroColorStyle.Red;
            QuestListView.Style = MetroColorStyle.Red;
            QuestTextBox.Style = MetroColorStyle.Red;
            OptionATextBox.Style = MetroColorStyle.Red;
            OptionBTextBox.Style = MetroColorStyle.Red;
            OptionCTextBox.Style = MetroColorStyle.Red;
            OptionDTextBox.Style = MetroColorStyle.Red;
            QuestModeComboBox.Style = MetroColorStyle.Red;
            QuestAnswerComboBox.Style = MetroColorStyle.Red;

            EditQuestButton.Style = MetroColorStyle.Red;
            EditButton.Style = MetroColorStyle.Red;
            QuestEditTextBox.Style = MetroColorStyle.Red;
            OptionAEditTextBox.Style = MetroColorStyle.Red;
            OptionBEditTextBox.Style = MetroColorStyle.Red;
            OptionCEditTextBox.Style = MetroColorStyle.Red;
            OptionDEditTextBox.Style = MetroColorStyle.Red;
            QuestModeEditComboBox.Style = MetroColorStyle.Red;
            QuestAnswerEditComboBox.Style = MetroColorStyle.Red;

            this.Style = MetroColorStyle.Red;
            this.Hide();
            this.Show();
        }
        private void RefreshButton_Click(object sender, EventArgs e)
        {
            Clear();
            Connect();
            List();
            Disconnect();
        }
        private void AddQuestButton_Click(object sender, EventArgs e)
        {
            this.Text = "Add Question";
            BackButton.Visible = true;
            ShowQuestsPanel.Visible = false;
            AddQuestPanel.Visible = true;
        }
        private void AddButton_Click(object sender, EventArgs e)
        {
            Quest = QuestTextBox.Text;
            OptionA = OptionATextBox.Text;
            OptionB = OptionBTextBox.Text;
            OptionC = OptionCTextBox.Text;
            OptionD = OptionDTextBox.Text;

            if (QuestModeComboBox.Text == "Expressions")
            {
                QuestMode = 2;
            }
            else if (QuestModeComboBox.Text == "Words")
            {
                QuestMode = 1;
            }

            if (QuestAnswerComboBox.Text == "Option A")
            {
                AnswerOfQuest = "A";
            }
            else if (QuestAnswerComboBox.Text == "Option B")
            {
                AnswerOfQuest = "B";
            }
            else if (QuestAnswerComboBox.Text == "Option C")
            {
                AnswerOfQuest = "C";
            }
            else if (QuestAnswerComboBox.Text == "Option D")
            {
                AnswerOfQuest = "D";
            }

            if (Quest == "" || OptionA == "" || OptionB == "" || OptionC == ""
                || OptionD == "" || AnswerOfQuest == "" || QuestModeComboBox.Text == "")
            {
                MetroMessageBox.Show(this, "Please fill the all blanks!", "WARNİNG!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                bool Exist = false;

                for (int i = 0; i < QuestListView.Items.Count; i++)
                {
                    if (QuestListView.Items[i].SubItems[1].Text == Quest)
                    {
                        MetroMessageBox.Show(this, "This question already exists!", "WARNİNG!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Exist = true;
                    }
                }

                if (Exist == false)
                {
                    Connect();
                    if (QuestMode == 1)
                    {
                        query = new OleDbCommand("INSERT INTO Words(quest, questAnswer, optionA, optionB, optionC, optionD) VALUES(@quest, @questAnswer, @optionA, @optionB, @optionC, @optionD)", connect);
                    }
                    else if (QuestMode == 2)
                    {
                        query = new OleDbCommand("INSERT INTO Expressions(quest, questAnswer, optionA, optionB, optionC, optionD) VALUES(@quest, @questAnswer, @optionA, @optionB, @optionC, @optionD)", connect);
                    }
                    query.Parameters.AddWithValue("@quest", Quest);
                    query.Parameters.AddWithValue("@questAnswer", AnswerOfQuest);
                    query.Parameters.AddWithValue("@optionA", OptionA);
                    query.Parameters.AddWithValue("@optionB", OptionB);
                    query.Parameters.AddWithValue("@optionC", OptionC);
                    query.Parameters.AddWithValue("@optionD", OptionD);
                    query.ExecuteNonQuery();
                    if (QuestMode == 1)
                    {
                        Clear();
                        query = new OleDbCommand("SELECT * FROM Words ORDER BY questID", connect);
                        List();
                    }
                    else if (QuestMode == 2)
                    {
                        Clear();
                        query = new OleDbCommand("SELECT * FROM Expressions ORDER BY questID", connect);
                        List();
                    }
                    SuccessPictureBox.Visible = true;
                    SuccessLabel.Visible = true;
                    SuccessTimer.Start();
                    Disconnect();
                }
            }
        }
        private void EditQuestButton_Click(object sender, EventArgs e)
        {
            NumberOfSelected = QuestListView.SelectedItems.Count;

            if (NumberOfSelected == 1)
            {
                foreach (ListViewItem DataOfSelected in QuestListView.SelectedItems)
                {
                    this.Text = "Edit Question";
                    BackButton.Visible = true;
                    ShowQuestsPanel.Visible = false;
                    QuestEditPanel.Visible = true;

                    QuestEditTextBox.Text = DataOfSelected.SubItems[1].Text;
                    OptionAEditTextBox.Text = DataOfSelected.SubItems[2].Text;
                    OptionBEditTextBox.Text = DataOfSelected.SubItems[3].Text;
                    OptionCEditTextBox.Text = DataOfSelected.SubItems[4].Text;
                    OptionDEditTextBox.Text = DataOfSelected.SubItems[5].Text;
                    QuestAnswerEditComboBox.Text = "Option " + DataOfSelected.SubItems[6].Text;

                    if (Mode1Button.Enabled == true)
                    {
                        QuestModeEditComboBox.Text = "Words";
                    }
                    else if (Mode2Button.Enabled == true)
                    {
                        QuestModeEditComboBox.Text = "Expressions";
                    }
                }
            }
            else if (NumberOfSelected > 1)
            {
                MetroMessageBox.Show(this, "You must select choose one!!", "WARNİNG!",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void EditButton_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem DataOfSelected in QuestListView.SelectedItems)
            {
                QuestID = Convert.ToInt32(DataOfSelected.SubItems[0].Text);
            }
                
            Quest = QuestEditTextBox.Text;
            OptionA = OptionAEditTextBox.Text;
            OptionB = OptionBEditTextBox.Text;
            OptionC = OptionCEditTextBox.Text;
            OptionD = OptionDEditTextBox.Text;

            if (QuestModeComboBox.Text == "Expressions")
            {
                QuestMode = 2;
            }
            else if (QuestModeComboBox.Text == "Words")
            {
                QuestMode = 1;
            }

            if (QuestAnswerComboBox.Text == "Option A")
            {
                AnswerOfQuest = "A";
            }
            else if (QuestAnswerComboBox.Text == "Option B")
            {
                AnswerOfQuest = "B";
            }
            else if (QuestAnswerComboBox.Text == "Option C")
            {
                AnswerOfQuest = "C";
            }
            else if (QuestAnswerComboBox.Text == "Option D")
            {
                AnswerOfQuest = "D";
            }

            if (Quest == "" || OptionA == "" || OptionB == "" || OptionC == ""
                || OptionD == "" || AnswerOfQuest == "" || QuestModeEditComboBox.Text == "")
            {
                MetroMessageBox.Show(this, "Please fill the all blanks!", "WARNİNG!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                bool Exist = false;

                for (int i = 0; i < QuestListView.Items.Count; i++)
                {
                    if (QuestListView.Items[i].SubItems[1].Text == Quest)
                    {
                        MetroMessageBox.Show(this, "This question already exists!", "WARNİNG!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Exist = true;
                    }
                }

                if (Exist == false)
                {
                    Connect();

                    if (QuestMode == 1)
                    {
                        query = new OleDbCommand("UPDATE Words SET quest = @quest, questAnswer = @questAnswer, optionA = @optionA, optionB = @optionB, optionC = @optionC, optionD = @optionD WHERE questID = @questID", connect);
                    }
                    else if (QuestMode == 2)
                    {
                        query = new OleDbCommand("UPDATE Expressions SET quest = @quest, questAnswer = @questAnswer, optionA = @optionA, optionB = @optionB, optionC = @optionC, optionD = @optionD WHERE questID = @questID", connect);
                    }
                    query.Parameters.AddWithValue("@quest", Quest);
                    query.Parameters.AddWithValue("@questAnswer", AnswerOfQuest);
                    query.Parameters.AddWithValue("@optionA", OptionA);
                    query.Parameters.AddWithValue("@optionB", OptionB);
                    query.Parameters.AddWithValue("@optionC", OptionC);
                    query.Parameters.AddWithValue("@optionD", OptionD);
                    query.Parameters.AddWithValue("@questID", QuestID);
                    query.ExecuteNonQuery();
                    if (QuestMode == 1)
                    {
                        Clear();
                        query = new OleDbCommand("SELECT * FROM Words ORDER BY questID", connect);
                        List();
                    }
                    else if (QuestMode == 2)
                    {
                        Clear();
                        query = new OleDbCommand("SELECT * FROM Expressions ORDER BY questID", connect);
                        List();
                    }

                    SuccessPictureBox.Visible = true;
                    SuccessLabel.Visible = true;
                    SuccessTimer.Start();
                    Disconnect();
                }
            }
        }
        private void DeleteButton_Click(object sender, EventArgs e)
        {
            NumberOfSelected = QuestListView.SelectedItems.Count;

            Connect();
            
            foreach (ListViewItem DataOfSelected in QuestListView.SelectedItems)
            {
                QuestID = Convert.ToInt32(DataOfSelected.SubItems[0].Text);
                
                if (Mode1Button.Enabled == false)
                {
                    query = new OleDbCommand("DELETE FROM Words WHERE questID = @QuestID", connect);
                }
                else if (Mode2Button.Enabled == false)
                {
                    query = new OleDbCommand("DELETE FROM Expressions WHERE questID = @QuestID", connect);
                }

                query.Parameters.AddWithValue("@questID", QuestID);
                query.ExecuteNonQuery();
            }

            if (Mode1Button.Enabled == false)
            {
                query = new OleDbCommand("SELECT * FROM Words ORDER BY questID", connect);
            }
            else if (Mode2Button.Enabled == false)
            {
                query = new OleDbCommand("SELECT * FROM Expressions ORDER BY questID", connect);
            }

            Clear();
            List();
            
            SuccessLabel1.Text = NumberOfSelected.ToString() + " Question(s) Deleted Successful!";
            SuccessLabel1.Visible = true;
            SuccessPictureBox1.Visible = true;
            SuccessTimer.Start();

            Disconnect();
        }
        private void BackButton_Click(object semder, EventArgs e)
        {
            BackButton.Visible = false;
            AddQuestPanel.Visible = false;
            QuestEditPanel.Visible = false;
            ShowQuestsPanel.Visible = true;
        }
        private void SuccessTimer_Tick(object sender, EventArgs e)
        {
            if (SuccessPictureBox.Visible == true)
            {
                SuccessPictureBox.Visible = false;
                SuccessLabel.Visible = false;
            }
            else if (SuccessPictureBox1.Visible == true)
            {
                SuccessPictureBox1.Visible = false;
                SuccessLabel1.Visible = false;
            }
            else if (SuccessPictureBox2.Visible == true)
            {
                SuccessPictureBox2.Visible = false;
                SuccessLabel2.Visible = false;

                BackButton.Visible = false;
                AddQuestPanel.Visible = false;
                QuestEditPanel.Visible = false;
                ShowQuestsPanel.Visible = true;
            }
            
            SuccessTimer.Stop();
        }
    }
}
