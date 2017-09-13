using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Net.Http.Headers;
using System.IO;
using OfficeOpenXml;
using System.Text;
using System.Collections;
using ExcelRead.Models;

namespace ExcelRead.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View(new List<ExcelViewModel>());
        }

        [HttpPost]
        public async Task<IActionResult> Index(IFormFile file)
        {
            IList<ExcelViewModel> list = new List<ExcelViewModel>();
            var validTypes = new string[] { "application/octet-stream", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel" };
            if (file == null)
            {
                ModelState.AddModelError("file", "Please select excel (.xlsx) file.");
            }

            if (file != null)
            {
                if (file.Length > 0)
                {
                    if (!validTypes.Contains(file.ContentType))
                    {
                        ModelState.AddModelError("file", "Only the following file types are allowed: .xlsx");
                    }
                }
            }

            if (!ModelState.IsValid)
            {
                return View(list);
            }
            else
            {
                if (file != null)
                {
                    if (file.Length > 0)
                    {
                        if (validTypes.Contains(file.ContentType))
                        {
                            string guid = Guid.NewGuid().ToString();
                            var parsedContentDisposition = ContentDispositionHeaderValue.Parse(file.ContentDisposition);
                            string uploadPath = "Uploads\\" + guid + "";
                            string path = Path.Combine(_hostingEnvironment.WebRootPath, uploadPath);
                            Directory.CreateDirectory(path);
                            var filePath = Path.Combine(path, parsedContentDisposition.FileName.Trim('"'));
                            FileInfo fileInfo = new FileInfo(filePath);

                            using (var stream = new FileStream(filePath, FileMode.Create))
                            {
                                await file.CopyToAsync(stream);
                            }

                            using (ExcelPackage package = new ExcelPackage(fileInfo))
                            {
                                var worksheets = package.Workbook.Worksheets;
                                foreach (ExcelWorksheet item in worksheets)
                                {
                                    StringBuilder sb = new StringBuilder();
                                    int rowCount = item.Dimension.Rows;
                                    int ColCount = item.Dimension.Columns;

                                    ExcelViewModel sheet = new ExcelViewModel();
                                    sheet.SheetName = item.Name;

                                    for (int row = 1; row <= rowCount; row++)
                                    {
                                        sb.Append("<tr>");
                                        for (int col = 1; col <= ColCount; col++)
                                        {
                                            if (row == 1)
                                            {
                                                sb.Append("<th>");
                                                sb.Append(Convert.ToString(item.Cells[row, col].Value));
                                                sb.Append("</th>");
                                            }
                                            else
                                            {
                                                sb.Append("<td>");
                                                sb.Append(Convert.ToString(item.Cells[row, col].Value));
                                                sb.Append("</td>");
                                            }
                                        }
                                        sb.Append("</tr>");
                                    }
                                    sheet.Data = sb.ToString();
                                    list.Add(sheet);
                                }
                            }
                        }
                    }
                }
            }

            return View(list);
        }

        public IActionResult CSVRead()
        {
            return View(new ExcelViewModel());
        }

        [HttpPost]
        public async Task<IActionResult> CSVRead(IFormFile file)
        {
            ExcelViewModel model = new ExcelViewModel();
            var validTypes = new string[] { "text/csv", "application/csv", "application/vnd.ms-excel", "application/octet-stream" };
            if (file == null)
            {
                ModelState.AddModelError("file", "Please select excel (.csv) file.");
            }

            if (file != null)
            {
                if (file.Length > 0)
                {
                    if (!validTypes.Contains(file.ContentType))
                    {
                        ModelState.AddModelError("file", "Only the following file types are allowed: .csv");
                    }
                }
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }
            else
            {
                if (file != null)
                {
                    if (file.Length > 0)
                    {
                        if (validTypes.Contains(file.ContentType))
                        {
                            string guid = Guid.NewGuid().ToString();
                            var parsedContentDisposition = ContentDispositionHeaderValue.Parse(file.ContentDisposition);
                            string uploadPath = "Uploads\\" + guid + "";
                            string path = Path.Combine(_hostingEnvironment.WebRootPath, uploadPath);
                            Directory.CreateDirectory(path);
                            var filePath = Path.Combine(path, parsedContentDisposition.FileName.Trim('"'));
                            FileInfo fileInfo = new FileInfo(filePath);

                            using (var stream = new FileStream(filePath, FileMode.Create))
                            {
                                await file.CopyToAsync(stream);
                            }

                            using (var streamReader = System.IO.File.OpenText(filePath))
                            {
                                List<string> data = new List<string>();
                                StringBuilder sb = new StringBuilder();

                                while (!streamReader.EndOfStream)
                                {
                                    string line = streamReader.ReadLine();
                                    data = new List<string>();
                                    data = line.Split(new[] { ',' }).ToList();
                                    sb.Append("<tr>");
                                    foreach (var item in data)
                                    {
                                        sb.Append("<td>");
                                        sb.Append(item);
                                        sb.Append("</td>");
                                    }
                                    sb.Append("</tr>");
                                }

                                model.Data = sb.ToString();
                            }
                        }
                    }
                }
            }

            return View(model);
        }

        public IActionResult TextRead()
        {
            return View(new ExcelViewModel());
        }

        [HttpPost]
        public async Task<IActionResult> TextRead(IFormFile file)
        {
            ExcelViewModel model = new ExcelViewModel();
            var validTypes = new string[] { "text/plain" };
            if (file == null)
            {
                ModelState.AddModelError("file", "Please select text (.txt) file.");
            }

            if (file != null)
            {
                if (file.Length > 0)
                {
                    if (!validTypes.Contains(file.ContentType))
                    {
                        ModelState.AddModelError("file", "Only the following file types are allowed: .txt");
                    }
                }
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }
            else
            {
                if (file != null)
                {
                    if (file.Length > 0)
                    {
                        if (validTypes.Contains(file.ContentType))
                        {
                            string guid = Guid.NewGuid().ToString();
                            var parsedContentDisposition = ContentDispositionHeaderValue.Parse(file.ContentDisposition);
                            string uploadPath = "Uploads\\" + guid + "";
                            string path = Path.Combine(_hostingEnvironment.WebRootPath, uploadPath);
                            Directory.CreateDirectory(path);
                            var filePath = Path.Combine(path, parsedContentDisposition.FileName.Trim('"'));
                            FileInfo fileInfo = new FileInfo(filePath);

                            using (var stream = new FileStream(filePath, FileMode.Create))
                            {
                                await file.CopyToAsync(stream);
                            }

                            using (var streamReader = System.IO.File.OpenText(filePath))
                            {
                                model.Data = streamReader.ReadToEnd();
                            }
                        }
                    }
                }
            }

            return View(model);
        }

        public IActionResult Error()
        {
            return View();
        }
    }
}
