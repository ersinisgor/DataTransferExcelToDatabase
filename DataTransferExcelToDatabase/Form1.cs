using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataTransferExcelToDatabase
{
	public partial class Form1 : Form
	{
		SqlConnection connection = new SqlConnection(@"Data Source=DESKTOP-1LT3S9C\SQL2022;Initial Catalog=C#_Projects_Udemy;Integrated Security=True;");
		public Form1()
		{
			InitializeComponent();
		}

		private void DatabaseToExcel_Click(object sender, EventArgs e)
		{
			Excel.Application excelApplication = new Excel.Application();
			excelApplication.Visible = true;
			Excel.Workbook workbook = excelApplication.Workbooks.Add(System.Reflection.Missing.Value);
			Excel.Worksheet worksheet = workbook.Sheets[1];

			string[] Titles = { "Personel No", "First Name", "Last Name", "District", "City" };
			Excel.Range range;

			for (int i = 0; i < Titles.Length; i++)
			{
				range = worksheet.Cells[1, i + 1];
				range.Value2 = Titles[i];
			}

			try
			{
				connection.Open();
				string sql = "SELECT PersonelNo, FirstName, LastName, District, City FROM Personel";
				SqlCommand command = new SqlCommand(sql, connection);
				SqlDataReader reader = command.ExecuteReader();

				while (reader.Read())
				{
					string personelNo = reader[0].ToString();
					string firstName = reader[1].ToString();
					string lastName = reader[2].ToString();
					string district = reader[3].ToString();
					string city = reader[4].ToString();
					richTextBox1.Text += personelNo + " " + firstName + " " + lastName + " " + district + " " + city + "\n";

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("An error occurred while reading data from SQL, ErrorCode: SQLREAD01 \n" + ex.ToString());
			}
			finally
			{
				if (connection != null)
					connection.Close();
			}



		}
	}
}
