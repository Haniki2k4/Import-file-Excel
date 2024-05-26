using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Import_file_Excel.Data;
using OfficeOpenXml;
using System.IO;
using System.Data.Entity.Validation;

namespace Import_file_Excel.Controllers
{
    public class ImportController : Controller
    {
        // GET: Import
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase excelFile)
        {
            string message = string.Empty;
            int count = 0;

            // Kiểm tra tệp tải lên
            if (excelFile != null && excelFile.ContentLength > 0 &&
                (excelFile.FileName.EndsWith(".xls") || excelFile.FileName.EndsWith(".xlsx")))
            {
                try
                {
                    string uploadsDir = Server.MapPath("~/App_Data/uploads");

                    // Kiểm tra và tạo thư mục nếu nó không tồn tại
                    if (!Directory.Exists(uploadsDir))
                    {
                        Directory.CreateDirectory(uploadsDir);
                    }

                    // Đường dẫn tệp
                    string path = Path.Combine(uploadsDir, Path.GetFileName(excelFile.FileName));

                    // Lưu tệp vào thư mục
                    excelFile.SaveAs(path);

                    // Gọi phương thức ImportData để xử lý tệp Excel
                    bool importResult = ImportData(path, out count);
                    if (importResult)
                    {
                        message = $"Successfully imported {count} records.";
                    }
                    else
                    {
                        message = "Failed to import data.";
                    }
                }
                catch (Exception ex)
                {

                }
            }
            else
            {
                message = "Invalid file format. Please upload an Excel file.";
            }

            ViewBag.Message = message;
            return View();
        }

        public bool ImportData(string filePath, out int count)
        {
            var result = false;
            count = 0;
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var package = new ExcelPackage(new FileInfo(filePath));
                int startColumn = 1;
                int startRow = 2;
                ExcelWorksheet workSheet = package.Workbook.Worksheets[0]; //sheet 1
                object data = null;

                Db_ThucTapEntities db = new Db_ThucTapEntities();

                do
                {
                    data = workSheet.Cells[startRow, startColumn].Value;
                    object MaXa = workSheet.Cells[startRow, startColumn].Value;
                    if (data != null && MaXa != null)
                    {
                        // Ghi log dữ liệu được đọc từ tệp Excel
                        System.Diagnostics.Debug.WriteLine($"Row {startRow}: Data={data}, MaXa={MaXa}");

                        //import data
                        var isSuccess = saveXa(MaXa.ToString(), db);
                        if (isSuccess)
                        {
                            count++;
                        }
                    }
                    else
                    {
                        // Ghi log nếu không có dữ liệu
                        System.Diagnostics.Debug.WriteLine($"Row {startRow}: No data found");
                    }
                    startRow++;
                }
                while (data != null);
                result = true;
            }
            catch (Exception ex)
            {
                // Log exception hoặc xử lý lỗi tương ứng
                System.Diagnostics.Debug.WriteLine($"Exception: {ex.Message}");
            }
            return result;
        }


        public bool saveXa(String MaXa, Db_ThucTapEntities db)
        {
            var result = false;
            try
            {
                // check exists
                if (db.XAs.Any(m => m.MaXa == MaXa))
                {
                    // Ghi log nếu dữ liệu đã tồn tại
                    System.Diagnostics.Debug.WriteLine($"Mã xã {MaXa} đã tồn tại.");
                }
                else
                {
                    var item = new XA();
                    item.MaXa = MaXa;
                    item.MaHuyen = MaHuyen;
                    item.MaXa = MaXa;
                    item.MaXa = MaXa;
                    db.XAs.Add(item);
                    db.SaveChanges();
                    result = true;
                }
            }
            catch (DbEntityValidationException ex)
            {
                foreach (var validationErrors in ex.EntityValidationErrors)
                {
                    foreach (var validationError in validationErrors.ValidationErrors)
                    {
                        System.Diagnostics.Debug.WriteLine($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log exception hoặc xử lý lỗi tương ứng
                System.Diagnostics.Debug.WriteLine($"Exception in saveXa: {ex.Message}");
            }
            return result;
        }

    }
}

