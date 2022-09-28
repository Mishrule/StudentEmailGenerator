using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentEmailGenerator
{
	public class StudentModel
	{
		public string Reference { get; set; }
		public string Surname { get; set; }
		public string Firstname { get; set; }
		public string Othername { get;  set; } = string.Empty;
		public string Email { get; set; } = String.Empty;
		public string DeptCode { get; set; }
		public string Password { get; set; }
	}
}
