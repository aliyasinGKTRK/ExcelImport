using DataAccess;
using Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Data.OleDb;

namespace ExcelImport.Controllers
{
    public class ExcelController : Controller
    {
        Product Product = new Product();
        private readonly IConfiguration configuration;
        private readonly ILogger<ExcelController> logger;
        ExcelImportContext context = new ExcelImportContext();
        public ExcelController(IConfiguration configuration, ILogger<ExcelController> logger)
        {
            this.configuration = configuration;
            this.logger = logger;
        }



        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(IFormFile formFile)
        {

            try
            {
                if (formFile.Length > 0)
                {

                    var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcelFile");
                    if (!Directory.Exists(mainPath))
                    {
                        Directory.CreateDirectory(mainPath);
                    }

                    var fileName = formFile.FileName;
                    string extension = Path.GetExtension(fileName);
                    string newName = new DateTime().Date.ToString("dd-mm-yyyy-hh-MM-ss") + extension;
                    var filePath = Path.Combine(mainPath, newName);
                    using (FileStream stream = new FileStream(filePath, FileMode.Create))
                    {
                        formFile.CopyTo(stream);
                    }
                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".xls":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=2\\'";
                            break;
                        case ".xlsx":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=2\\'";
                            break;
                    }

                    DataTable dt = new DataTable();
                    conString = string.Format(conString, filePath);
                    using (OleDbConnection conExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = new OleDbCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = conExcel;
                                conExcel.Open();
                                DataTable dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                conExcel.Close();

                            }
                        }
                    }


                    conString = configuration.GetConnectionString("ExcelImportContext");
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            sqlBulkCopy.DestinationTableName = "Products";
                            sqlBulkCopy.ColumnMappings.Add("ProductName", "ProductName");
                            sqlBulkCopy.ColumnMappings.Add("QuantityPerUnit", "QuantityPerUnit");
                            sqlBulkCopy.ColumnMappings.Add("Color", "Color");
             

                            con.Open();


                            sqlBulkCopy.WriteToServer(dt);
                            con.Close();
                        }
                    }

                    ViewBag.massage = "File Import successfully, data saved into database";

                    return View();
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, ex.Message);
                string msg = ex.Message;
            }

            return View();
        }
    }
}
