using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.IO;
using WebAppExportExcel.Models;
using System.Globalization;

namespace WebAppExportExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            FileViewModel model = new FileViewModel();
            return View(model);
        }

        public IActionResult Privacy()
        {
            return View();
        }
        public IActionResult Clean()
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files");
            if (Directory.Exists(path))
                Directory.Delete(path, true);
            return View("Index", new FileViewModel());
        }

        [HttpPost]
        public IActionResult ExportExcel(FileViewModel model)
        {
            if (!ModelState.IsValid)
            {
                ModelState.AddModelError("InvalidAuth", "Chưa chọn file");
                return View("Index", model);
            }
            try
            {

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);


            string fileNameWithPath = Path.Combine(path, model.filetick.FileName);
            using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
            {
                model.filetick.CopyTo(stream);
            }
            fileNameWithPath = Path.Combine(path, model.filecong.FileName);
            using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
            {
                model.filecong.CopyTo(stream);
            }

            
                string ticktay = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/"+model.filetick.FileName);


                string chamcong = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + model.filecong.FileName);
                Stream memoryStream = new MemoryStream();
                Console.WriteLine(ticktay);
                Console.WriteLine(chamcong);
                FileInfo filepathticktay = new FileInfo(ticktay);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                List<ticktay> listticktay = new List<ticktay>();
                using (var excelPack = new ExcelPackage(filepathticktay))
                {
                    //Load excel stream
                    using (var stream = System.IO.File.OpenRead(ticktay))
                    {
                        excelPack.Load(stream);
                        var ws = excelPack.Workbook.Worksheets[0];
                        var rowCount = ws.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            ticktay item = new ticktay();
                            item.Name = ws.Cells[row, 3].Value.ToString().Trim();
                            item.dateTime = ws.Cells[row, 4].Value.ToString().Trim();
                            item.NewState = ws.Cells[row, 6].Value != null ? ws.Cells[row, 6].Value.ToString().Trim() : string.Empty;
                            item.Exception = ws.Cells[row, 7].Value.ToString().Trim();
                            item.Date = item.dateTime.Substring(0, 10);
                            item.Time = item.dateTime.Substring(11);
                            DateTime dateTime = DateTime.ParseExact(item.dateTime, "dd/MM/yyyy h:mm tt", CultureInfo.InvariantCulture);
                            listticktay.Add(item);
                        }
                    }
                }
          
                FileInfo filepathchamcong = new FileInfo(chamcong);
                using (var excelPack = new ExcelPackage(filepathchamcong))
                {
                    //Load excel stream
                    using (var stream = System.IO.File.OpenRead(chamcong))
                    {
                        excelPack.Load(stream);
                        var ws = excelPack.Workbook.Worksheets[0];
                        var rowCount = ws.Dimension.Rows;

                        for (int row = 3; row <= rowCount; row++)
                        {
                            if (ws.Cells[row, 2].Value == null) break;
                            var name = ws.Cells[row, 2].Value.ToString().Trim();
                            var dateValue = ws.Cells[row, 4].Value.ToString().Trim();
                            int dateInt = 0;
                            int.TryParse(dateValue, out dateInt);
                            var date = DateTime.FromOADate(dateInt);
                            DateTime dt;
                            name = name.NonUnicode();
                            var obj = listticktay.Where(x => x.Name == name).Where(x => x.Date == date.ToString("dd/MM/yyyy")).ToList();
                            if (obj.Count == 4)
                            {
                                int i = 4;
                                foreach (var item in obj)
                                {
                                    i++;
                                    ws.Cells[row, i].Value = item.Time.ToString();
                                }
                                var dateTimeChuan = DateTime.ParseExact(obj[1].Date + " 7:00 AM", "dd/MM/yyyy h:mm tt", CultureInfo.InvariantCulture);
                                var dateTime = DateTime.ParseExact(obj[0].dateTime, "dd/MM/yyyy h:mm tt", CultureInfo.InvariantCulture);
                                var time = dateTimeChuan - dateTime;
                                var min = time.TotalMinutes;
                                if (min < 0)
                                {
                                    min = min * -1;
                                    ws.Cells[row, 11].Value = min;
                                }
                                else ws.Cells[row, 11].Value = string.Empty;


                            }
                            else if (obj.Count == 3)
                            {
                                int i = 5;
                                var itemNull = obj.Where(x => x.NewState != "OverTime Out" && x.NewState != "OverTime In").FirstOrDefault();
                                var index = obj.IndexOf(itemNull);
                                if (index == 0)
                                {
                                    ws.Cells[row, i + 2].Value = obj[1].Time.ToString();
                                    ws.Cells[row, i + 3].Value = obj[2].Time.ToString();
                                }
                                else
                                {
                                    ws.Cells[row, i].Value = obj[0].Time.ToString();
                                    ws.Cells[row, i + 1].Value = obj[1].Time.ToString();
                                    var dateTimeChuan = DateTime.ParseExact(obj[1].Date + " 7:00 AM", "dd/MM/yyyy h:mm tt", CultureInfo.InvariantCulture);
                                    var dateTime = DateTime.ParseExact(obj[0].dateTime, "dd/MM/yyyy h:mm tt", CultureInfo.InvariantCulture);
                                    var time = dateTimeChuan - dateTime;
                                    var min = time.TotalMinutes;
                                    if (min < 0)
                                    {
                                        min = min * -1;
                                        ws.Cells[row, 11].Value = min;
                                    }
                                }
                            }
                        }
                    }
                    excelPack.Save();
                }

                var streamreturn = System.IO.File.OpenRead(chamcong);
                return File(streamreturn, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",model.filecong.FileName);
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                //  WriteLogFile.WriteLog(string.Format("{0}{1}", "MoMo", DateTime.Now.ToString("ddMMyyyy")), string.Format("Ip: {0}. DataResponse {1}" , ex.Message.ToString(), DateTime.Now), "MoMo");
                ModelState.AddModelError("InvalidAuth", ex.ToString());
                return View("Index", model);
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}