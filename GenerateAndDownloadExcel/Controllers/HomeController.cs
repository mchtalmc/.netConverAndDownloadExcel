using ClosedXML.Excel;
using GenerateAndDownloadExcel.Models;
using Microsoft.AspNetCore.Mvc;

namespace GenerateAndDownloadExcel.Controllers
{
	public class HomeController : Controller
	{
		List<Student> _students = new List<Student>();

		public HomeController()
		{
			for (int i = 0; i < 10; i++)
			{
				_students.Add(new Student()
				{
				StudentId= i,
				Name= "Student"+i,
				Surname= "Student Surname"+i
				});
			}
		}
		

		public IActionResult Index()
		{
			using (var workbook = new XLWorkbook())
			{
				#region Header

			
				var worksheet = workbook.Worksheets.Add("Students");
				var currentRow = 1;
				worksheet.Cell(currentRow, 1).Value = "StudentId";
				worksheet.Cell(currentRow, 2).Value = "Name";
				worksheet.Cell(currentRow, 3).Value = "Surname";
				#endregion


				#region Body
				foreach (var student in _students)
				{
					currentRow++;
					worksheet.Cell(currentRow, 1).Value = student.StudentId;
					worksheet.Cell(currentRow, 2).Value = student.Name;
					worksheet.Cell(currentRow, 3).Value = student.Surname;
				}
				#endregion

				using (var stream = new MemoryStream())
				{
					workbook.SaveAs(stream);
					var content = stream.ToArray();
					return File(
											content,
											"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
											"ExcelConvert.xlsx"
											);
				}


			}
		}

		
	}
}
