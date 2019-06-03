using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTimetableGenerator.Models;
using ExcelTimetableGenerator.Shared;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelTimetableGenerator.Pages
{
    public class ExportTimetablesModel : PageModel
    {
        private readonly ExcelTimetableGenerator.Data.ApplicationDbContext _context;
        private IHostingEnvironment _hostingEnvironment;
        public ExportTimetablesModel(ExcelTimetableGenerator.Data.ApplicationDbContext context, IHostingEnvironment hostingEnvironment)
        {
            _context = context;
            _hostingEnvironment = hostingEnvironment;
        }

        public string SavePath { get; set; }

        public int NumFilesExported { get; set; }

        public IList<Customer> Customer { get; set; }

        public IList<Programme> Programme { get; set; }

        public IList<Programme> ProgrammesAndCourses { get; set; }

        public IList<Course> Course { get; set; }

        public async Task<IActionResult> OnGetAsync(string academicYear, int planRevisionID)
        {
            planRevisionID = 66;

            string CurrentAcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);
            var academicYearParam = new SqlParameter("@AcademicYear", CurrentAcademicYear);
            var planRevisionIDParam = new SqlParameter("@PlanRevisionID", planRevisionID);

            Course = await _context.Course
                .FromSql("EXEC SPR_ETG_CourseData @AcademicYear, @PlanRevisionID", academicYearParam, planRevisionIDParam)
                .ToListAsync();

            Programme = await _context.Programme
                .FromSql("EXEC SPR_ETG_ProgrammeData @AcademicYear, @PlanRevisionID", academicYearParam, planRevisionIDParam)
                .ToListAsync();

            SavePath = _hostingEnvironment.WebRootPath + @"\ExportedTimetables";
            string filePath = null;
            string fileName = null;
            string fileURL = null;
            int rowNum = 0;
            NumFilesExported = 0;

            foreach (var programme in Programme)
            {
                NumFilesExported += 1;

                //Save timetables into departmental and team folders (creates a folder if it doesn't exist)
                filePath = SavePath + @"\" + programme.FacCode + @"\" + programme.TeamCode + @"\";
                Directory.CreateDirectory(filePath);

                //Generate each Excel workbook
                fileName = @"" + programme.ProgCode + " - " + programme.ProgTitle + ".xlsx";
                fileName = FileFunctions.MakeValidFileName(fileName); //Sanitize
                fileURL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, fileName);
                FileInfo file = new FileInfo(Path.Combine(filePath, fileName));
                var memory = new MemoryStream();
                using (var fs = new FileStream(Path.Combine(filePath, fileName), FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook;
                    workbook = new XSSFWorkbook();

                    //Cell formats
                    CellType ctString = CellType.String;
                    CellType ctNumber = CellType.Numeric;
                    CellType ctFormula = CellType.Formula;
                    CellType ctBlank = CellType.Blank;

                    //Cell Colours
                    XSSFColor cLightBlue = new XSSFColor(Color.LightSteelBlue);
                    XSSFColor cBlue = new XSSFColor(Color.RoyalBlue);
                    XSSFColor cMainAim = new XSSFColor(Color.LemonChiffon);
                    XSSFColor cEngMaths = new XSSFColor(Color.PaleGreen);
                    XSSFColor cWex = new XSSFColor(Color.Plum);
                    XSSFColor cTut = new XSSFColor(Color.LightCyan);
                    XSSFColor cOther = new XSSFColor(Color.NavajoWhite);

                    //Cell Borders
                    BorderStyle border = BorderStyle.Medium;

                    //Fonts
                    //Tab Header
                    IFont fHeader = workbook.CreateFont();
                    fHeader.FontHeight = 20;
                    fHeader.FontName = "Arial";
                    fHeader.IsBold = true;

                    //Table Header
                    IFont fTableHeader = workbook.CreateFont();
                    fTableHeader.IsBold = true;

                    //Cell formats
                    //Tab Header
                    XSSFCellStyle sHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                    sHeader.SetFont(fHeader);

                    //Table Header
                    XSSFCellStyle sTableHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                    sTableHeader.SetFont(fTableHeader);
                    sTableHeader.SetFillForegroundColor(cLightBlue);
                    sTableHeader.FillPattern = FillPattern.SolidForeground;
                    sTableHeader.BorderTop = border;
                    sTableHeader.BorderBottom = border;
                    sTableHeader.BorderLeft = border;
                    sTableHeader.BorderRight = border;

                    //Date
                    XSSFCellStyle sDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    ICreationHelper createHelper = workbook.GetCreationHelper();
                    sDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));
                    sDate.BorderTop = border;
                    sDate.BorderBottom = border;
                    sDate.BorderLeft = border;
                    sDate.BorderRight = border;

                    //Types of courses
                    XSSFCellStyle sMainAim = (XSSFCellStyle)workbook.CreateCellStyle();
                    sMainAim.SetFillForegroundColor(cMainAim);
                    sMainAim.FillPattern = FillPattern.SolidForeground;
                    sMainAim.BorderTop = border;
                    sMainAim.BorderBottom = border;
                    sMainAim.BorderLeft = border;
                    sMainAim.BorderRight = border;

                    XSSFCellStyle sEngMaths = (XSSFCellStyle)workbook.CreateCellStyle();
                    sEngMaths.SetFillForegroundColor(cEngMaths);
                    sEngMaths.FillPattern = FillPattern.SolidForeground;
                    sEngMaths.BorderTop = border;
                    sEngMaths.BorderBottom = border;
                    sEngMaths.BorderLeft = border;
                    sEngMaths.BorderRight = border;

                    XSSFCellStyle sWex = (XSSFCellStyle)workbook.CreateCellStyle();
                    sWex.SetFillForegroundColor(cWex);
                    sWex.FillPattern = FillPattern.SolidForeground;
                    sWex.BorderTop = border;
                    sWex.BorderBottom = border;
                    sWex.BorderLeft = border;
                    sWex.BorderRight = border;

                    XSSFCellStyle sTut = (XSSFCellStyle)workbook.CreateCellStyle();
                    sTut.SetFillForegroundColor(cTut);
                    sTut.FillPattern = FillPattern.SolidForeground;
                    sTut.BorderTop = border;
                    sTut.BorderBottom = border;
                    sTut.BorderLeft = border;
                    sTut.BorderRight = border;

                    XSSFCellStyle sOther = (XSSFCellStyle)workbook.CreateCellStyle();
                    sOther.SetFillForegroundColor(cOther);
                    sOther.FillPattern = FillPattern.SolidForeground;
                    sOther.BorderTop = border;
                    sOther.BorderBottom = border;
                    sOther.BorderLeft = border;
                    sOther.BorderRight = border;

                    //Border Only
                    XSSFCellStyle sBorder = (XSSFCellStyle)workbook.CreateCellStyle();
                    sBorder.BorderTop = border;
                    sBorder.BorderBottom = border;
                    sBorder.BorderLeft = border;
                    sBorder.BorderRight = border;

                    //Create Index Sheet
                    ISheet indexSheet = workbook.CreateSheet("Index");
                    IRow indexRow = indexSheet.CreateRow(0);
                    indexRow = indexSheet.CreateRow(1);
                    indexRow = indexSheet.CreateRow(2);
     
                    ICell cell = indexRow.CreateCell(0);
                    
                    cell.SetCellValue(programme.ProgCode + " - " + programme.ProgTitle);
                    cell.CellStyle = sHeader;

                    indexRow = indexSheet.CreateRow(3);

                    indexRow = indexSheet.CreateRow(4);

                    cell = indexRow.CreateCell(0);
                    cell.SetCellValue("Course Code");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(1);
                    cell.SetCellValue("Course Title");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(2);
                    cell.SetCellValue("Aim Code");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(3);
                    cell.SetCellValue("Award Body");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(4);
                    cell.SetCellValue("Weeks");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(5);
                    cell.SetCellValue("PLH 16-18");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(6);
                    cell.SetCellValue("EEP 16-18");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(7);
                    cell.SetCellValue("PLH 19+");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(8);
                    cell.SetCellValue("EEP 19+");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(9);
                    cell.SetCellValue("Start Date");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(10);
                    cell.SetCellValue("End Date");
                    cell.CellStyle = sTableHeader;

                    cell = indexRow.CreateCell(11);
                    cell.SetCellValue("Status");
                    cell.CellStyle = sTableHeader;

                    indexSheet.SetColumnWidth(0, 16 * 256);
                    indexSheet.SetColumnWidth(1, 40 * 256);
                    indexSheet.SetColumnWidth(2, 16 * 256);
                    indexSheet.SetColumnWidth(3, 12 * 256);
                    indexSheet.SetColumnWidth(4, 8 * 256);
                    indexSheet.SetColumnWidth(5, 10 * 256);
                    indexSheet.SetColumnWidth(6, 10 * 256);
                    indexSheet.SetColumnWidth(7, 10 * 256);
                    indexSheet.SetColumnWidth(8, 10 * 256);
                    indexSheet.SetColumnWidth(9, 12 * 256);
                    indexSheet.SetColumnWidth(10, 12 * 256);
                    indexSheet.SetColumnWidth(11, 20 * 256);

                    //Line for each course
                    rowNum = 4;

                    //Generate each Excel worksheet
                    if (programme.Course != null && programme.Course.Count > 0)
                    {
                        foreach (var course in programme.Course)
                        {
                            rowNum += 1;
                            indexRow = indexSheet.CreateRow(rowNum);
                            XSSFCellStyle cellStyle;
                            XSSFCellStyle cellStyleDate;

                            switch(course.CourseOrder)
                            {
                                case 0:
                                    cellStyle = sOther;
                                    cellStyleDate = sDate;
                                    break;
                                case 1:
                                    cellStyle = sMainAim;
                                    cellStyleDate = sDate;
                                    break;
                                case 2:
                                    cellStyle = sOther;
                                    cellStyleDate = sDate;
                                    break;
                                case 3:
                                    cellStyle = sEngMaths;
                                    cellStyleDate = sDate;
                                    break;
                                case 4:
                                    cellStyle = sEngMaths;
                                    cellStyleDate = sDate;
                                    break;
                                case 5:
                                    cellStyle = sTut;
                                    cellStyleDate = sDate;
                                    break;
                                case 6:
                                    cellStyle = sWex;
                                    cellStyleDate = sDate;
                                    break;
                                case 7:
                                    cellStyle = sOther;
                                    cellStyleDate = sDate;
                                    break;
                                default:
                                    cellStyle = sMainAim;
                                    cellStyleDate = sDate;
                                    break;
                            }

                            cell = indexRow.CreateCell(0, ctString);
                            cell.SetCellValue(course.CourseCode);
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(1, ctString);
                            cell.SetCellValue(course.CourseTitle);
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(2, ctString);
                            cell.SetCellValue(course.AimCode);
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(3, ctString);
                            cell.SetCellValue(course.AwardBody);
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(4, ctNumber);
                            if(course.Weeks >= 0)
                            {
                                cell.SetCellValue((double)course.Weeks);
                            }
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(5, ctNumber);
                            if (course.PLH1618 >= 0)
                            {
                                cell.SetCellValue((double)course.PLH1618);
                            }
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(6, ctNumber);
                            if (course.EEP1618 >= 0)
                            {
                                cell.SetCellValue((double)course.EEP1618);
                            }
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(7, ctNumber);
                            if (course.PLH19 >= 0)
                            {
                                cell.SetCellValue((double)course.PLH19);
                            }
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(8, ctNumber);
                            if (course.EEP19 >= 0)
                            {
                                cell.SetCellValue((double)course.EEP19);
                            }
                            cell.CellStyle = cellStyle;

                            cell = indexRow.CreateCell(9, ctNumber);
                            cell.CellStyle = sDate;
                            cell.SetCellValue((DateTime)course.StartDate);

                            cell = indexRow.CreateCell(10, ctNumber);
                            cell.CellStyle = sDate;
                            cell.SetCellValue((DateTime)course.EndDate);

                            cell = indexRow.CreateCell(11, ctString);
                            cell.SetCellValue(course.CourseStatus.ToString());
                            cell.CellStyle = cellStyle;
                        }

                        ////Set widths(does not seem to work very well)
                        //int numberOfColumns = indexSheet.GetRow(4).PhysicalNumberOfCells;
                        //for (int i = 1; i <= numberOfColumns; i++)
                        //{
                        //    indexSheet.AutoSizeColumn(i);
                        //    GC.Collect(); // Add this line
                        //}
                    }
                    else
                    {
                        indexRow = indexSheet.CreateRow(rowNum);

                        cell = indexRow.CreateCell(0, ctString);
                        cell.SetCellValue("NO COURSES ATTACHED");
                    }

                    

                    if (programme.Course != null && programme.Course.Count > 0)
                    {
                        foreach (var course in programme.Course)
                        {
                            ISheet excelSheet = workbook.CreateSheet(course.CourseCode);
                            IRow row = excelSheet.CreateRow(0);

                            row.CreateCell(0).SetCellValue("ID");
                            row.CreateCell(1).SetCellValue("Name");
                            row.CreateCell(2).SetCellValue("Age");

                            row = excelSheet.CreateRow(1);
                            row.CreateCell(0).SetCellValue(1);
                            row.CreateCell(1).SetCellValue("Kane Williamson");
                            row.CreateCell(2).SetCellValue(29);

                            row = excelSheet.CreateRow(2);
                            row.CreateCell(0).SetCellValue(2);
                            row.CreateCell(1).SetCellValue("Martin Guptil");
                            row.CreateCell(2).SetCellValue(33);

                            row = excelSheet.CreateRow(3);
                            row.CreateCell(0).SetCellValue(3);
                            row.CreateCell(1).SetCellValue("Colin Munro");
                            row.CreateCell(2).SetCellValue(23);
                        }
                    }
                    else
                    {
                        ISheet excelSheet = workbook.CreateSheet("No Courses");
                        IRow row = excelSheet.CreateRow(0);

                        row.CreateCell(0).SetCellValue("ID");
                        row.CreateCell(1).SetCellValue("Name");
                        row.CreateCell(2).SetCellValue("Age");

                        row = excelSheet.CreateRow(1);
                        row.CreateCell(0).SetCellValue(1);
                        row.CreateCell(1).SetCellValue("Kane Williamson");
                        row.CreateCell(2).SetCellValue(29);

                        row = excelSheet.CreateRow(2);
                        row.CreateCell(0).SetCellValue(2);
                        row.CreateCell(1).SetCellValue("Martin Guptil");
                        row.CreateCell(2).SetCellValue(33);

                        row = excelSheet.CreateRow(3);
                        row.CreateCell(0).SetCellValue(3);
                        row.CreateCell(1).SetCellValue("Colin Munro");
                        row.CreateCell(2).SetCellValue(23);
                    }

                    workbook.Write(fs);
                }
                using (var stream = new FileStream(Path.Combine(filePath, fileName), FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
            }

            return Page();
        }
    }
}