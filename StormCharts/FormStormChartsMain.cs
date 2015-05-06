using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

namespace StormCharts
{
    public partial class FormStormChartsMain : Form
    {
        string decPlaces = "0.0000";
        
        int StormID = 1;
        string fileName = "";
        int resultsID = 0;

        string folder = "";
        private static string CONNECTION_STR = "Data Source=BESDBPROD2;Initial Catalog=NEPTUNE;Trusted_Connection = true;";

        public FormStormChartsMain()
        {
            InitializeComponent();
        }

        private void buttonCreateStormCharts_Click(object sender, EventArgs e)
        {
            SaveFileDialog theDialog = new SaveFileDialog();
            theDialog.DefaultExt = "xlsx";

            DialogResult theResult = theDialog.ShowDialog();
            
            if (theResult == DialogResult.OK)
            {
                folder = theDialog.FileName;
                pnlCancelBackgroundWorker.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                this.Refresh();
                backgroundWorkerSingle.RunWorkerAsync();
            }
        }

        private void backgroundWorkerSingle_DoWork(object sender, DoWorkEventArgs e)
        {
            //Create the query to retrive storm times
            string GetStormTimes =
                StormChartsQueries.GetStormTimes(
                                                  (int)numericUpDownH2Number.Value,
                                                  dateTimePickerStartTime.Value,
                                                  dateTimePickerEndTime.Value
                                                );

            //folder = System.IO.Path.GetDirectoryName(fileName);
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            
            //Get a connection to BESDBPROD2
            using (SqlConnection conn = new SqlConnection(CONNECTION_STR))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = GetStormTimes;
                cmd.CommandType = System.Data.CommandType.Text;
                
                conn.Open();
                cmd.CommandTimeout = 0;

                //We are returning a table, so we need to be prepared for this info
                //Create a dataRowReader, Call the query
                SqlDataReader reader = cmd.ExecuteReader();
                dt.Load(reader);
                reader.Close();
            }

            int StormNumber = 1;
            foreach (DataRow dr in dt.Rows)
            {
                //MessageBox.Show(((int)dr[0]).ToString());
                //Call the 5 minute stored procedure
                using (SqlConnection conn = new SqlConnection(CONNECTION_STR))
                {
                    dt2.Clear();

                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "[dbo].[USP_MODEL_RAIN]";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    SqlParameter start_date = cmd.Parameters.AddWithValue("@start_date", (RoundDown((DateTime)dr[1], TimeSpan.FromMinutes(5))));
                    SqlParameter end_date = cmd.Parameters.AddWithValue("@end_date", (RoundUp((DateTime)dr[2], TimeSpan.FromMinutes(5))));
                    SqlParameter interval = cmd.Parameters.AddWithValue("@interval", 5);
                    SqlParameter daypart = cmd.Parameters.AddWithValue("@daypart", "minute");
                    SqlParameter h2_number = cmd.Parameters.AddWithValue("@h2_number", (int)dr[0]);
                    SqlParameter limit_rows = cmd.Parameters.AddWithValue("@limit_rows", -1);

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    SqlDataReader reader = cmd.ExecuteReader();
                    dt2.Load(reader);
                    reader.Close();
                    dt2.ExportToExcel(RoundDown((DateTime)dr[1], TimeSpan.FromMinutes(5)).ToShortDateString(), (DateTime)dr[1], folder, StormNumber++, StormNumber);
                }
            }
        }

        private void backgroundWorkerSingle_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pnlCancelBackgroundWorker.Visible = false;

            if (e.Cancelled)
            {
                MessageBox.Show("StormCharts Creation Cancelled",
                            "StormCharts Canceled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error executing StormCharts: " + e.Error.Message,
                            "Error Executing StormCharts", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Successfully executed StormCharts!",
                            "StormCharts Executed Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        DateTime RoundUp(DateTime dt, TimeSpan d)
        {
            return new DateTime(((dt.Ticks + d.Ticks - 1) / d.Ticks) * d.Ticks);
        }

        DateTime RoundDown(DateTime dt, TimeSpan d)
        {
            return new DateTime((dt.Ticks / d.Ticks) * d.Ticks);
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {

        }

        private void buttonCreateChartOneStormManyGauges_Click(object sender, EventArgs e)
        {
            SaveFileDialog theDialog = new SaveFileDialog();
            theDialog.DefaultExt = "xlsx";

            DialogResult theResult = theDialog.ShowDialog();

            if (theResult == DialogResult.OK)
            {
                folder = theDialog.FileName;
                pnlCancelBackgroundWorker.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                this.Refresh();
                backgroundWorkerMultiple.RunWorkerAsync();
            }
        }

        private void backgroundWorkerMultiple_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pnlCancelBackgroundWorker.Visible = false;

            if (e.Cancelled)
            {
                MessageBox.Show("StormCharts Creation Cancelled",
                            "StormCharts Canceled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error executing StormCharts: " + e.Error.Message,
                            "Error Executing StormCharts", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Successfully executed StormCharts!",
                            "StormCharts Executed Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void backgroundWorkerMultiple_DoWork(object sender, DoWorkEventArgs e)
        {
            //Create the query to retrive storm times
            /*string GetStormTimes =
                StormChartsQueries.GetStormTimes(
                                                  (int)numericUpDownH2Number.Value,
                                                  dateTimePickerStartTime.Value,
                                                  dateTimePickerEndTime.Value
                                                );
            */

            //dt should hold the ids of the raingages we want
            List<int> dt = new List<int>();
            dt.Add(214);
            dt.Add(213);
            dt.Add(192);
            dt.Add(181);
            dt.Add(175);
            dt.Add(174);
            dt.Add(173);
            dt.Add(171);
            dt.Add(164);
            dt.Add(117);
            dt.Add(64);
            dt.Add(12);
            dt.Add(6);

            //dt2 holds the rainfall of the raingage in question
            DataTable dt2 = new DataTable();

            int StormNumber = 1;
            bool LastStorm;

            foreach (int dr in dt)
            {
                //MessageBox.Show(((int)dr[0]).ToString());
                //Call the 5 minute stored procedure
                using (SqlConnection conn = new SqlConnection(CONNECTION_STR))
                {
                    dt2.Clear();
                    LastStorm = StormNumber == dt.Count();

                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "[dbo].[USP_MODEL_RAIN]";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;            

                    SqlParameter start_date = cmd.Parameters.AddWithValue("@start_date", (RoundDown(dateTimePickerStartTime.Value, TimeSpan.FromMinutes(5))));
                    SqlParameter end_date = cmd.Parameters.AddWithValue("@end_date", (RoundUp(dateTimePickerEndTime.Value, TimeSpan.FromMinutes(5))));
                    SqlParameter interval = cmd.Parameters.AddWithValue("@interval", 5);
                    SqlParameter daypart = cmd.Parameters.AddWithValue("@daypart", "minute");
                    SqlParameter h2_number = cmd.Parameters.AddWithValue("@h2_number", (int)dr);
                    SqlParameter limit_rows = cmd.Parameters.AddWithValue("@limit_rows", -1);

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    SqlDataReader reader = cmd.ExecuteReader();
                    dt2.Load(reader);
                    reader.Close();
                    dt2.ExportToExcel(RoundDown(dateTimePickerStartTime.Value, TimeSpan.FromMinutes(5)).ToShortDateString(), (DateTime)dateTimePickerStartTime.Value, folder, (int)dr, StormNumber++, LastStorm);
                }
            }
        }
    }
}
