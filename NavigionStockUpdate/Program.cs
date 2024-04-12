using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net.Mime;
using ClosedXML.Excel;

namespace NavigionStockUpdate
{
    class Program
    {
        // Get connection string from configuration
        static string con = ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString;

        // Get testing SKU from configuration
        static readonly string testingSKU = ConfigurationManager.AppSettings["TestingSKU"];

        // Get base URL from configuration
        static readonly string baseurl = ConfigurationManager.AppSettings["baseurl"];

        // Initialize variable to store faulty object
        static Value exceptFaultyObject;

        // Initialize variable to store exception log
        static string ExceptionLog = string.Empty;

        // Get API username from configuration
        static readonly string apiusername = ConfigurationManager.AppSettings["apiusername"];

        // Get API password from configuration
        static readonly string apipassword = ConfigurationManager.AppSettings["apipassword"];

        static void Main(string[] args)
        {
            string maxValue = "0";
            bool lastValue = false;
            //string inputJson = (new JavaScriptSerializer()).Serialize(input);
            DataTable dt = new DataTable();
            dt.Columns.Add("SKU");
            dt.Columns.Add("Quantity");

            DataTable dtEXAMPLE1 = dt.Clone();
            DataTable dtEXAMPLE2 = dt.Clone();

            List<string> Etag = new List<string>();
            string apiUrl = string.Empty;
            // Loop until the lastValue flag is true
            while (!lastValue)
            {
                try
                {
                    // Construct the SKU filter string
                    var skuList = "";
                    foreach (var sku in testingSKU.Split(','))
                    {
                        skuList += "Item_No eq '" + sku + "' or ";
                    }
                    skuList = skuList.Remove(skuList.Length - 3);

                    // Construct the API URL with filters and sorting
                    apiUrl = baseurl + "?$format=json&$filter=" +
                        "(Location_Code eq 'TOCW' or " +
                        "Location_Code eq 'EXAMPLE2') " +
                        " and (" + skuList + ") & $orderby= Item_No asc ";

                    // Set the lastValue flag to true
                    lastValue = true;

                    // Initialize variables for API response and retry count
                    string reply = string.Empty;
                    int retrycount = 0;

                    // Retry loop to handle network errors
                    Retry:
                    try
                    {
                        // Create a WebClient to make API request
                        WebClient client = new WebClient();
                        client.Credentials = new NetworkCredential(apiusername, apipassword);

                        // Download the API response
                        reply = client.DownloadString(apiUrl);
                    }
                    catch (Exception ex)
                    {
                        // Retry logic for network errors
                        if (retrycount < 5)
                        {
                            retrycount++;
                            goto Retry;
                        }
                        else
                        {
                            // Send email notification after 5 retries
                            // Code for sending email notification goes here
                        }
                    }

                    // Deserialize the API response into Response object
                    Responce responce = JsonConvert.DeserializeObject<Responce>(reply);

                    // Log the response retrieval time
                    Console.WriteLine("Got Response :" + DateTime.Now);

                    // Process the response if there are items
                    if (responce.value.Count > 0)
                    {
                        // Update the maxValue variable
                        maxValue = responce.value[responce.value.Count - 1].Item_No;

                        // Iterate through each item in the response
                        foreach (Value items in responce.value)
                        {
                            try
                            {
                                exceptFaultyObject = items;

                                // Check if the ETag is not already processed
                                if (!Etag.Contains(items.ETag))
                                {
                                    Etag.Add(items.ETag); // Add ETag to the list

                                    // Add item to DataTable based on Location_Code
                                    if (items.Location_Code == "TOCW")
                                    {
                                        DataRow allrow = dtEXAMPLE1.NewRow();
                                        allrow["SKU"] = items.Item_No;
                                        allrow["Quantity"] = Convert.ToInt32(items.Net_Inventory);
                                        dtEXAMPLE1.Rows.Add(allrow);
                                    }
                                    else if (items.Location_Code == "EXAMPLE2")
                                    {
                                        DataRow kwtrow = dtEXAMPLE2.NewRow();
                                        kwtrow["SKU"] = items.Item_No;
                                        kwtrow["Quantity"] = Convert.ToInt32(items.Net_Inventory);
                                        dtEXAMPLE2.Rows.Add(kwtrow);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Skiped");
                                }

                                // Check for loop termination conditions
                                if (responce.value.Count <= 0 || maxValue == responce.value[0].Item_No)
                                {
                                    lastValue = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                // Handle exceptions during item processing
                                // Code for exception handling goes here
                            }
                        }
                        Console.WriteLine("-----------------------------");
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions during main loop execution
                    // Code for exception handling goes here
                }
            }

            try
            {
                // Create a new SqlConnection object using the connection string
                using (SqlConnection conn = new SqlConnection(con))
                {
                    // Specify the name of the stored procedure to be executed
                    string procedurename = "ProcedureForUpdatingEntriesInSQLDatabase";

                    // Create a new SqlCommand object for executing the stored procedure
                    using (SqlCommand cmd = new SqlCommand(procedurename, conn))
                    {
                        cmd.CommandTimeout = 0; // Set command timeout to unlimited
                        cmd.CommandType = CommandType.StoredProcedure; // Specify command type as stored procedure

                        // Add parameters to the SqlCommand for passing DataTables to the stored procedure
                        cmd.Parameters.AddWithValue("@dtEX_1", dtEXAMPLE1);
                        cmd.Parameters.AddWithValue("@dtEX_2", dtEXAMPLE2);

                        // Open the database connection
                        conn.Open();

                        // Log the start time of the update process
                        Console.WriteLine("Updating :" + DateTime.Now);

                        // Execute the SqlCommand to update the database
                        cmd.ExecuteNonQuery();

                        // Log the end time of the update process
                        Console.WriteLine("Ends :" + DateTime.Now);
                    }
                }

                // Retrieve the file path of the generated report from the database
                string generatedFilePath = GetDataFromDataBase();

                // Check if the file path is not empty
                if (!string.IsNullOrEmpty(generatedFilePath))
                {
                    // Send an email with the attachment
                    // Code for sending email with attachment goes here
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions by logging or sending email notification
                // Code for sending email notification or logging exceptions goes here
            }

        }


        public static string GetDataFromDataBase()
        {
            SqlConnection con = null; // Declare SqlConnection object outside try-catch block for proper disposal in finally block

            try
            {
                // Create a new SqlConnection object using the connection string from configuration
                con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString);

                // Open the database connection
                con.Open();

                // Create a new SqlCommand object for executing the stored procedure
                SqlCommand cmd = new SqlCommand("ProcedureForGettingStockuUdateReport", con);
                cmd.CommandTimeout = 0; // Set command timeout to unlimited
                cmd.CommandType = CommandType.StoredProcedure; // Specify command type as stored procedure

                // Create a new SqlDataAdapter object to fill the DataTable with query results
                SqlDataAdapter adp = new SqlDataAdapter(cmd);

                // Create a new DataTable to hold the query results
                DataTable dt = new DataTable();
                dt.TableName = "StockUpdateTableName"; // Set DataTable name

                // Fill the DataTable with data from the stored procedure
                adp.Fill(dt);

                // Execute the SqlCommand to ensure any pending changes are committed
                cmd.ExecuteNonQuery();

                // Create response by generating an Excel file from the DataTable
                return GenerateFile(dt);
            }
            catch (SqlException ex)
            {
                // Handle SQL exceptions by displaying error message
                Console.WriteLine("\n *** SQL Exception occurred : " + ex.Message.ToString());
                return ""; // Return an empty string in case of exception
            }
            finally
            {
                // Close the database connection in the finally block to ensure proper cleanup
                if (con != null && con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }
        }


        public static string GenerateFile(DataTable QueryOutput)
        {
            try
            {
                // Get the directory where the application is executing
                string appDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                // Generate a unique filename based on the current date and time
                string filename = "StockUpdate_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".xlsx";

                // Combine directory path and filename to create the complete file path
                string filepath = System.IO.Path.Combine(appDirectory, filename);

                // Convert DataTable to Excel file using ClosedXML library
                using (XLWorkbook wb = new XLWorkbook())
                {
                    // Add DataTable contents to a new worksheet in the Excel workbook
                    wb.Worksheets.Add(QueryOutput, QueryOutput.TableName);

                    // Apply formatting to the worksheet
                    wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Align text horizontally to center
                    wb.Style.Font.Bold = true; // Make text bold

                    // Save the Excel workbook to the specified filepath
                    wb.SaveAs(filepath);
                }

                // Return the filepath of the generated Excel file
                return filepath;
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during file generation
                Console.WriteLine("An error occurred while generating the Excel file: " + ex.Message);
                return string.Empty; // Return an empty string if file generation fails
            }
        }
    }
}
