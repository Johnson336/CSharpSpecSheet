using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpSpecSheet
{
    public partial class AdminMenu : Form
    {

        SQLiteConnection m_dbConnection;

        public AdminMenu()
        {
            InitializeComponent();
            initializeDatabase();
        }

        private void initializeDatabase()
        {
            if (!File.Exists("db.sqlite"))
            {
                SQLiteConnection.CreateFile("db.sqlite");
            }

            m_dbConnection = new SQLiteConnection("Data Source=db.sqlite;Version=3;");
            m_dbConnection.Open();
        }

        private int executeSQL(string stmt)
        {
            SQLiteCommand cmd = new SQLiteCommand(stmt, m_dbConnection);

            //lblStatusBar2.Text = lblStatusBar.Text;
            //lblStatusBar.Text = stmt + ": " + cmd.ExecuteNonQuery() + " rows affected.";
            return cmd.ExecuteNonQuery();
        }

        private SQLiteDataReader executeSQLReader(string stmt)
        {
                SQLiteCommand command = new SQLiteCommand(stmt, m_dbConnection);
                return command.ExecuteReader();
        }

        private void buttonSQLExecute_Click(object sender, EventArgs e)
        {
            txtSQLResults.Text = "";
            if (txtSQLStmt.Text.Contains("SELECT") || (txtSQLStmt.Text.Contains("INSERT")) || (txtSQLStmt.Text.Contains("UPDATE")))
            {
                SQLiteDataReader result = executeSQLReader(txtSQLStmt.Text);
                do
                {
                    int count = result.FieldCount;
                    while (result.Read())
                    {
                        for (int i = 0; i < count; i++)
                        {
                            txtSQLResults.Text += (result.GetValue(i)) + "\t";
                        }
                        txtSQLResults.Text += System.Environment.NewLine;
                    }
                } while (result.NextResult());
            } else if (txtSQLStmt.Text.Contains("DELETE"))
            {
                txtSQLResults.Text += executeSQL(txtSQLStmt.Text) + " rows affected.";
            }
        }

        private void AdminMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_dbConnection.Close();
        }
    }
}
