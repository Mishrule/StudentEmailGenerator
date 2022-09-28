
using OfficeOpenXml;

namespace StudentEmailGenerator
{
	public class Program
	{
		public static async Task Main()
		{
			Console.WriteLine("Initialing Package");
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			Console.WriteLine("Loading Excel File");
			var file = new FileInfo(@"C:\Users\mishr\Desktop\Testing\destinationFolder\Emails.xlsx");

			List<StudentModel> students = await LoadExcelFile(file);

			foreach (var record in students)
			{
				Console.WriteLine($"Ref: {record.Reference} |FirstName: {record.Firstname} |Othername: {record.Othername} |Surname: {record.Surname} |Email: {record.Email}| Password: {record.Password}");

			}

			Console.WriteLine("Completed");
		}

		public static async Task<List<StudentModel>> LoadExcelFile(FileInfo file)
		{
			Console.WriteLine("Calling Method");
			List<StudentModel> output = new();
			using var package = new ExcelPackage(file);
			await package.LoadAsync(file);
			var ws = package.Workbook.Worksheets[0];

			var code = string.Empty;
			int row = 2;
			int col = 1;

			while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
			{
				StudentModel st = new();
				//st.Reference = ws.Cells[row, col + 0].Value?.ToString().Substring(0, 3).Trim();
				var reff = ws.Cells[row, col + 0].Value?.ToString().Trim();
				st.Reference = reff;
				var prefix = reff.Substring(reff.Length - 4).Trim();


				st.Firstname = ws.Cells[row, col + 1].Value?.ToString().Substring(0, 1).Trim();
				var othername = ws.Cells[row, col + 2].Value?.ToString().Substring(0, 1).Trim();
				if (othername != null)
				{
					st.Othername = othername.Substring(0, 1).Trim();
				}

				st.Surname = ws.Cells[row, col + 3].Value?.ToString().Trim();
				if (st.Surname.Contains(" "))
				{
					st.Surname.Replace(" ", String.Empty);
				}


				code = ws.Cells[row, col + 4].Value.ToString().ToLower().Trim();
				st.DeptCode = code;

				var email  = $"{code}-{st.Firstname}{st.Othername}{st.Surname}{prefix}@st.umat.edu.gh".Replace(" ",string.Empty).ToLower().Trim();

				ws.Cells[row, col+ 5].Value = email;
				st.Email = ws.Cells[row, col + 5].Value?.ToString().ToLower().Trim();


				

				ws.Cells[row, col+ 6].Value = $"{st.Surname}{prefix}".Replace(" ",string.Empty).ToLower();
				

				 st.Password = $"{st.Surname}{prefix}".Replace(" ", string.Empty).ToLower();
			 
				
				
				
				
				
				
				output.Add(st);
				row += 1;
			}


			await package.SaveAsync();

			return output;
		}

	}
}

