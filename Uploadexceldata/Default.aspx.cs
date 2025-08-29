using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Uploadexceldata
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            

        }


        protected void UploadExcelFile(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                try
                {
                    // Save the uploaded file
                    string filePath = Server.MapPath("~/Uploads/") + FileUpload1.FileName;
                   //int indexToAccess = 0;

                    // Open the Excel file using EPPlus
                    using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
                    {
                        FileUpload1.SaveAs(filePath);
                        if (package.Workbook.Worksheets.Count == 0)
                        {
                            throw new Exception("The workbook contains no worksheets.");
                        }
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                        int totalRows = worksheet.Dimension.Rows;

                        // Create a DataTable to hold the Excel data
                        DataTable dataTable = new DataTable();
                        dataTable.Columns.Add("Column1", typeof(string));
                        dataTable.Columns.Add("Column2", typeof(string));
                        dataTable.Columns.Add("Column3", typeof(int));

                        for (int row = 2; row <= totalRows; row++) // Assuming the first row contains headers
                        {
                            DataRow newRow = dataTable.NewRow();
                            newRow["Column1"] = worksheet.Cells[row, 1].Text; // First column
                            newRow["Column2"] = worksheet.Cells[row, 2].Text; // Second column
                            newRow["Column3"] = int.Parse(worksheet.Cells[row, 3].Text); // Third column (integer)
                            dataTable.Rows.Add(newRow);
                        }

                        // Pass the DataTable to the stored procedure
                        //string connectionString = "MyDB";
                        string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["MyDB"].ConnectionString;

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            using (SqlCommand command = new SqlCommand("InsertExcelData", connection))
                            {
                                command.CommandType = CommandType.StoredProcedure;

                                // Add the table-valued parameter
                                SqlParameter tvpParam = command.Parameters.AddWithValue("@ExcelData", dataTable);
                                tvpParam.SqlDbType = SqlDbType.Structured;
                                tvpParam.TypeName = "ExcelDataType";

                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions (e.g., log errors or show a message)
                    Response.Write("Error: " + ex.Message);
                }
            }
            else
            {
                Response.Write("Please upload a file.");
            }
        }

    }
}