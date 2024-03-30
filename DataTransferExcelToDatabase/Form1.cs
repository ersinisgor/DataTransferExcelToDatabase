using System.Collections;
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

				int row = 2;
				while (reader.Read())
				{
					string personelNo = reader[0].ToString();
					string firstName = reader[1].ToString();
					string lastName = reader[2].ToString();
					string district = reader[3].ToString();
					string city = reader[4].ToString();
					richTextBox1.Text += personelNo + " " + firstName + " " + lastName + " " + district + " " + city + "\n";


					range = worksheet.Cells[row, 1];
					range.Value2 = personelNo;
					range = worksheet.Cells[row, 2];
					range.Value2 = firstName;
					range = worksheet.Cells[row, 3];
					range.Value2 = lastName;
					range = worksheet.Cells[row, 4];
					range.Value2 = district;
					range = worksheet.Cells[row, 5];
					range.Value2 = city;
					row++;

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

		private void ExcelToDatabase_Click(object sender, EventArgs e)
		{
			Excel.Application excelApplication;
			Excel.Workbook workbook;
			Excel.Worksheet worksheet;
			Excel.Range range;

			int rowCounter;
			int columnCounter;

			excelApplication = new Excel.Application();
			workbook = excelApplication.Workbooks.Open(@"C:\Users\Ersin\Desktop\test.xlsx");
			worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
			range = worksheet.UsedRange;

			richTextBox2.Clear();

			for (rowCounter = 2; rowCounter <= range.Rows.Count; rowCounter++)
			{
				ArrayList list = new ArrayList();

				for (columnCounter = 1; columnCounter <= range.Columns.Count; columnCounter++)
				{
					string value = Convert.ToString((range.Cells[rowCounter, columnCounter] as Excel.Range).Value2);
					richTextBox2.Text += value + " ";
					list.Add(value);
				}

				richTextBox2.Text += "\n";

				try
				{
					connection.Open();
					SqlCommand sqlCommand = new SqlCommand("INSERT INTO Personel (PersonelNo, FirstName, LastName, District, City) "
																								 + "VALUES (@PersonelNo, @FirstName, @LastName, @District, @City)", connection);
					sqlCommand.Parameters.AddWithValue("@PersonelNo", list[0]);
					sqlCommand.Parameters.AddWithValue("@FirstName", list[1]);
					sqlCommand.Parameters.AddWithValue("@LastName", list[2]);
					sqlCommand.Parameters.AddWithValue("@District", list[3]);
					sqlCommand.Parameters.AddWithValue("@City", list[4]);
					sqlCommand.ExecuteNonQuery();

				}
				catch (Exception ex)
				{
					MessageBox.Show("An error occurred while writing data to SQL, ErrorCode: SQLWRITE01 \n" + ex.ToString());
				}
				finally
				{
					if (connection != null)
						connection.Close();
				}


			}

			excelApplication.Quit();
			ReleaseObject(worksheet);
			ReleaseObject(workbook);
			ReleaseObject(excelApplication);
		}

		private void ReleaseObject(Object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{

				obj = null;
			}
			finally
			{
				GC.Collect();

			}
		}
	}
}
