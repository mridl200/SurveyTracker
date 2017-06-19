using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Diagnostics;

/***************
 * Survey Inventory Program
 * Allows for new surveys to be electronically documented.
 * Scans in the barcode for each survey and then places them in a virtual box, which
 * should match the physical box number they are placed in.
 * ->May require review for SQL parameters
 * Feb 26/2015
 * *************/
namespace CoreInventory
{
    public partial class Form1 : Form
    {
        int id = 0; //Global var used to track surveys
        string imagePath = ""; //Global var to survey images
        DateTime date = DateTime.Now;
        string dbDate = "";
        int globalErrorCount = 0;
        int Barcode_int = 0;
        
        public Form1()
        {
            InitializeComponent();

        }

        private void dgDisplay_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = (int)comboBox1.SelectedValue;
            string ImagePath = getImagePath(id);
            string filename = "";
         
            if (e.ColumnIndex == 1)
            {
                filename = dgDisplay[e.ColumnIndex, e.RowIndex].Value.ToString();
                filename = @ImagePath + filename + ".tif";
                if (File.Exists(filename))
                {
                
                    Process.Start(filename);
                    
                }
                else
                {
                    MessageBox.Show(filename);
                }
            }
           //globalErrorCount++;
        }

        private void dgDisplay_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                string appt = Convert.ToString(e.Value);
                if (appt != "")
                {
                    dgDisplay.Rows[e.RowIndex].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                    dgDisplay.Rows[e.RowIndex].Cells[1].Style.Font = 
                new Font("Ariel", 8, FontStyle.Bold | FontStyle.Underline);  
                }
            }
        }

        /*************************
         * Checks Combo box for Survey Type.
         * Sets up inital vars required on load
         * @param EventArgs, Sender object
         * @return Nothing
         * ***********************/
        private void Form1_Load(object sender, EventArgs e)
        {
            
            /*Populate Drop Down List from Database */
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxx"].ConnectionString;
            SqlConnection conn = new SqlConnection(ConnString);
            
            conn.Open();
            //Querys the database to get both the SurveyDescription(used to populate Combobox) and SurveyID
            SqlCommand sc = new SqlCommand("select [SurveyDesc],[ID] from xxxxxxxxx", conn);
            SqlDataReader reader;

            reader = sc.ExecuteReader();
            DataTable dt = new DataTable();
    
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("SurveyDesc", typeof(string));
            dt.Load(reader);

            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "SurveyDesc";
            comboBox1.DataSource = dt;

            int surveyID = (int)comboBox1.SelectedValue;  //Returns ID of selected Survey
           
            conn.Close();
            
            lblDate.Text = DateTime.Today.ToShortDateString(); //Displays todays date
            string databaseDestination = FindDestination(surveyID); //Returns the database to save information into
            int highestBoxNumber = FindHighestBox(databaseDestination); //Returns the Higest(last) Box number used      
            int count = FindCount(highestBoxNumber, databaseDestination);  //Counts how many surveys are in the selected box
            txtBox.Text = highestBoxNumber.ToString(); //Required.  Takes the Highest Box Number and populates the Box number text box.      
            lblCounter.Text = count.ToString();  //Displays how many surveys in the box.
        }

        /*************************
         * Returns the maximum amount of physical surveys per box.  
         * This value is set in the database.
         * Different surveys may take up more, or less, physical room than another.
         * @Param  Survey ID
         * @Return Highest number of surveys allowed
         * ***********************/
        private int MaxSurveysPerBox(int ID)
        {
            int maxSurveys = 0;
            string query = "SELECT [MaxSurveyPerBox] FROM xxxxxxxxxxxxxx WHERE ID = " + @ID;
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxx"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                using (var command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@id", id); 
                    try
                    {
                        conn.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                maxSurveys = (int)reader[0];  //Stores the max amount of surveys in a physical box
                            }
                            if (reader.Read())
                            {
                                throw new Exception("Too many rows");
                            }

                            else
                            {
                                throw new Exception("No rows");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return maxSurveys;

        }


        /***********************
         * Querys the database to get the database that 
         * each scanned survey should be placed into.
         * @Param Survey ID
         * @Return Database name to store survey information
         * *********************/
        private string FindDestination(int ID)
        {
            string databaseDestination = "";

            string query = "SELECT [DestTable] FROM xxxxxxxxxxxxxx WHERE ID = " + @ID;
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxxx"].ConnectionString;


            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                using (var command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@id", id);
                    try
                    {
                        conn.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                databaseDestination = reader[0].ToString();  //Returns the Destination Table
                            }
                            if (reader.Read())
                            {
                                throw new Exception("Too many rows");
                            }

                            else
                            {
                                throw new Exception("No rows");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return databaseDestination;
        }

        /***************************
         * Querys the database to get the amount of physical surveys in a box
         * @Param  The Box number.  The Database location of the survey
         * @Return The total count of objects inside the selected box
         * *************************/
        private int FindCount(int box, string databaseDestination)
        {
            string dbSource = databaseDestination;
            int id = (int)comboBox1.SelectedValue; //Get Int read in from combo box
            int count = 0;

            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxx"].ConnectionString;
            string sql = "select count(*) from " + @dbSource +" where box_int=@box";

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@box", box);

                try
                {
                    conn.Open();
                    count = Convert.ToInt32(cmd.ExecuteScalar());

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            return count;

        }
    
        /*************************
         * Returns the largest(last) box in the selected database
         * @Param The database in question
         * @Return The highest numbered box that isnt full
         * ***********************/
        private int FindHighestBox(string databaseDestination)
        {
            int highestBox = 0;
            string textbox = databaseDestination;
            string sql = "";

            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxx"].ConnectionString;

            if (textbox == "boxtable")
                 sql = "SELECT MAX(box_int) FROM " + @databaseDestination + " WHERE box_int < 14000";
            else
                 sql = "SELECT MAX(box_int) FROM " + @databaseDestination ;

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@databaseDestination", databaseDestination);
                try
                {
                    conn.Open();
                    highestBox = (int)cmd.ExecuteScalar();
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            int bID = Convert.ToInt32(highestBox);

            return bID;
        }

        /***********************
         *Returns the selected Surveys Image Path
         *@Param Surveys ID
         *@Return Location of the surveys images
         ***********************/
        private string getImagePath(int id)
        {
            id = this.id;
            string query = "SELECT [ImagePath] FROM xxxxxxxxxxxxx WHERE ID = " + @id;
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxxx"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                using (var command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@id", id);
                    try
                    {
                        conn.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {

                                imagePath = reader[0].ToString();
                                
                                if (reader.Read())
                                {
                                    throw new Exception("Too many rows");
                                }
                            }
                            else
                            {
                                throw new Exception("No rows");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return imagePath;
        }
                

        /**************************
         * Refactoring required.  Vars are overwritten with the same information
         * */
        private void txtBarcode_Leave_1(object sender, EventArgs e)
        {
             
            //Init vars
            
            string imagePath = "";
            string surveyDescription = comboBox1.Text;
            string databaseSource = "";
            string databaseDestination = "";
            string textbox = databaseDestination;
           int id = (int)comboBox1.SelectedValue; //Get Int read in from combo box

            //Query database to get all required fields
            string query = "SELECT [ID], [SurveyDesc], [SourceTable], [DestTable], [ImagePath] FROM xxxxxxxxxxxxxxxxxx WHERE ID = " + @id;
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxx"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                using (var command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@id", id); //Double check that I did this right
                    try
                    {
                        conn.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                id = (int)reader[0];
                                surveyDescription = reader[1].ToString();
                                databaseSource = reader[2].ToString();
                                databaseDestination = reader[3].ToString();
                                imagePath = reader[4].ToString();

                                if (reader.Read())
                                {
                                    throw new Exception("Too many rows");
                                }
                            }
                            else
                            {
                                throw new Exception("No rows");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                
               
                
            }

      
            /*********************************/
            if (txtBarcode.Text.Length != 8 && txtBarcode.Text.Length>0 && databaseDestination != "tblRESH_Inventory" )
            {
                MessageBox.Show("Invalid Barcode");
                return;

            }
            else if (txtBarcode.Text == "")
            {
                //do nothing
            }
            else
            {

                int data = CheckHasData(databaseSource);
                    if (data < 1)
                {

                    MessageBox.Show("No Data in the Database");  
                    return;
                }

                int image = CheckHasImage(imagePath);

                if (image < 1)
                {
                    MessageBox.Show("No Image in the U drive");
                    return;
                }

                int sameDate = SameDate();

                if (sameDate == 1)
                {
                    int duplicate = CheckHasDuplicate(databaseDestination);
                    if ((databaseDestination == "tblCurrentConsent") && (duplicate > 0))
                    {
                        MessageBox.Show("Consent already in Inventory Database.  Adding duplicate");
                     
                       
                    }
                    if ((databaseDestination != "tblCurrentConsent") && (duplicate > 0))
                    {

                        MessageBox.Show("Record already exists in the Inventory Database");
                        return;
                    }

                }



                if (sameDate == 1)
                {
                int success = InsertIntoTable(databaseDestination, id);

               
                    if (success < 0)
                    {
                        MessageBox.Show("Record Failed to be inserted ");
                        
                      
                    }


                    DisplayGrid(databaseDestination, success);

                    this.dgDisplay.CellFormatting += new DataGridViewCellFormattingEventHandler(dgDisplay_CellFormatting);
                   // this.dgDisplay.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgDisplay_CellClick);

                }
                this.txtBarcode.Clear();
                this.txtBarcode.Focus();
            }
            txtBarcode.Text = "";
            
            
        }

        private int SameDate()
        {

            Barcode_int = int.Parse(txtBarcode.Text);
            string query = "SELECT [Inventory_Date_dtm] FROM xxxxxxxxxxxxxxxxx WHERE Barcode_int = " + @Barcode_int;
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxx"].ConnectionString;
            int success = 0;
            string dbDate = "";

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                using (var command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@Barcode_int", Barcode_int);
                    try
                    {
                        conn.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {

                                dbDate = reader[0].ToString();

                                if (reader.Read())
                                {
                                    throw new Exception("Too many rows");
                                }
                            }
                            else
                            {

                                throw new Exception("No rows");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }


                bool result = string.IsNullOrEmpty(dbDate);
                string dateWithFormat = date.ToString();
                DateTime newDate;
                if (result == false)
                {
                    newDate = Convert.ToDateTime(dbDate);

                    if (date.Date == newDate.Date)
                    {
                        MessageBox.Show("Same date.  Record not inserted.");
                        success = 0;

                    }
                    else
                        success = 1;

                }
                else success = 1;
            }
            return success;


          
        }
        private void DisplayGrid(string databaseDestination, int success)
        {
            int   highestBox = FindHighestBox(databaseDestination);
            int count = FindCount(highestBox, databaseDestination);
        
            string constring = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxx"].ConnectionString; ;
            using (SqlConnection con = new SqlConnection(constring))
            {
                using (SqlCommand cmd = new SqlCommand("select top(10) * from dbo." + @databaseDestination + " where box_int =@highestBox and inventory_date_dtm >= @date order by inventory_date_dtm desc", con))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@databaseDestination", databaseDestination);
                    cmd.Parameters.AddWithValue("@highestBox", Convert.ToInt32(txtBox.Text));
                    cmd.Parameters.AddWithValue("@date", (DateTime.Today));

                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);
                            dgDisplay.DataSource = dt;
                           
                        }
                    }
                }
            }
           
        }

        private int InsertIntoTable(string databaseDestination,int id)
        {
            int success = 0;
            int maxSurveyPerBox = MaxSurveysPerBox(id);
            int highestBox = FindHighestBox(databaseDestination);
            
            int count = FindCount(highestBox, databaseDestination);
            txtBox.Text = highestBox.ToString();

            lblCounter.Text = count.ToString();
            if (count < maxSurveyPerBox)
            {

                string ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxx"].ConnectionString;
                string sql = "insert into dbo." + @databaseDestination + " (barcode_int,box_int,inventory_date_dtm) values(@barcode,@highestBox,@date)";

                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@databaseDestination", databaseDestination);
                    cmd.Parameters.AddWithValue("@barcode", Convert.ToInt32(txtBarcode.Text));
                    cmd.Parameters.AddWithValue("@highestBox", Convert.ToInt32(txtBox.Text));
                    cmd.Parameters.AddWithValue("@date", (DateTime.Now));
                    try
                    {
                        conn.Open();
                        success = Convert.ToInt32(cmd.ExecuteNonQuery());

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else
            {
                string ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxxx"].ConnectionString;
                string sql = "insert into dbo." + @databaseDestination + " (barcode_int,box_int,inventory_date_dtm) values(@barcode,@highestBox,@date)";

                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@databaseDestination", databaseDestination);
                    cmd.Parameters.AddWithValue("@barcode", Convert.ToInt32(txtBarcode.Text));
                    cmd.Parameters.AddWithValue("@highestBox", (Convert.ToInt32(txtBox.Text)) + 1);
                    cmd.Parameters.AddWithValue("@date", (DateTime.Now));
                    try
                    {
                        conn.Open();
                        success = Convert.ToInt32(cmd.ExecuteNonQuery()) + 1;
                        txtBox.Text = (highestBox + 1).ToString();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }


            }
            return success;


        }

        private int CheckHasDuplicate(string databaseDestination)
        {
            
            int yes = 0;
            String ConnString = ConfigurationManager.ConnectionStrings["xxxxxxxxxxxxxxxxxxx"].ConnectionString;
            string sql = "select count(*) from dbo."+databaseDestination+"  where barcode_int=@barcode";

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@barcode", Convert.ToInt32(txtBarcode.Text));

                try
                {
                    conn.Open();
                    yes = Convert.ToInt32(cmd.ExecuteScalar());

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return yes;
        }


        private int CheckHasImage(string imagePath)
        {

            int yes = 0;
            string filename = @imagePath;
            string textbox = comboBox1.Text;
            try
            {
                // Only get files that are of type tif

                string[] dirs = Directory.GetFiles(@filename, "*.tif", SearchOption.AllDirectories);

                int count = dirs.Where(item => item.Contains(txtBarcode.Text)).Count();

                if (count > 1)
                {
                    if (textbox == "Consent")
                    {
                        MessageBox.Show("Multiple Images available for " + txtBarcode.Text + " Please check and delete");

                        //bool result=LaunchFolderView(filename);

                        string argument = @"/select, " + filename + txtBarcode.Text + ".tif";
                        //string argument1 = @"/select, " + filename + txtBarcode.Text + "_1.tif";
                        Process.Start("explorer.exe", argument);
                        //Process.Start("explorer.exe", argument1);
                        return 0;
                    }
                   /* MessageBox.Show("Multiple Images available for " + txtBarcode.Text + " Please check and delete");

                    //bool result=LaunchFolderView(filename);

                    string argument1 = @"/select, " + filename + txtBarcode.Text+".tif";
                    //string argument1 = @"/select, " + filename + txtBarcode.Text + "_1.tif";
                    Process.Start("explorer.exe",argument1);
                    //Process.Start("explorer.exe", argument1);
                    */
                    
                    return 0;
                   
                }
              

                foreach (string x in dirs)
                {
                  

                    if (x.Contains(txtBarcode.Text))
                    {
                        yes = 1;

                        return yes;
                    }
                }
                
               
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }


            return yes;

        }

        private bool LaunchFolderView(string p_Filename)
        {
            bool l_result = false;

            // Check the file exists
            if (File.Exists(p_Filename))
            {
                // Check the folder we get from the file exists
                // this function would just get "C:\Hello" from
                // an input of "C:\Hello\World.txt"
                //string l_folder = FileSystemHelpers.GetPathFromQualifiedPath(p_Filename);

                // Check the folder exists
                if (Directory.Exists(p_Filename))
                {
                    try
                    {
                        // Start a new process for explorer
                        // in this location     
                        ProcessStartInfo l_psi = new ProcessStartInfo();
                        l_psi.FileName = "explorer";
                        l_psi.Arguments = string.Format("");
                        l_psi.UseShellExecute = true;

                        Process l_newProcess = new Process();
                        l_newProcess.StartInfo = l_psi;
                        l_newProcess.Start();

                        // No error
                        l_result = true;
                    }
                    catch (Exception exception)
                    {
                        throw new Exception("Error in 'LaunchFolderView'.", exception);
                    }
                }
            }

            return l_result;
        }

        private int CheckHasData(string databaseSource)
        {
            int yes = 0;

           

            String ConnString = ConfigurationManager.ConnectionStrings["ATP_Portal1"].ConnectionString;
            string sql = "select count(*) from "+@databaseSource+" where barcode=@barcode";

            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@barcode", Convert.ToInt32(txtBarcode.Text));

                try
                {
                    conn.Open();
                    yes = Convert.ToInt32(cmd.ExecuteScalar());

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            return yes;
        }

        /*******************
         * Refactor later (lots of reused calls
         * When ComboBox is switched this method is called.
         * Will initalize variables required for the rest of the program
         * *****************/
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            id = (int)comboBox1.SelectedValue;
            int surveyID = (int)comboBox1.SelectedValue;
            string databaseDestination = FindDestination(surveyID);
            int highestBox = FindHighestBox(databaseDestination);
            int count = FindCount(highestBox, databaseDestination);
            txtBox.Text = highestBox.ToString();
            lblCounter.Text = count.ToString();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void txtBox_TextChanged(object sender, EventArgs e)
        {

        }

       

        private void txtBarcode_TextChanged(object sender, EventArgs e)
        {

        }

       
    }
}
