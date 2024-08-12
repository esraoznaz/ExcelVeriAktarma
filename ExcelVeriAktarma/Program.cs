using System;
using System.Data;
using System.Data.SqlClient;
using ClosedXML.Excel;

class Program
{
	static void Main()
	{
		// Excel dosya yolu ve SQL bağlantı dizesi
		string excelFilePath = @"C:\Users\ASUS\OneDrive\Masaüstü\ilceler.xlsx"; // Excel dosya yolunu buraya yazın
		string connectionString = "Server=LAPTOP-3H9G77VD\\SQLEXPRESS;Database=Kres_Basvuru;Integrated Security=True"; // SQL bağlantı dizesini buraya yazın

		// Excel dosyasını okuyup DataTable'a yükleme
		DataTable dataTable = ReadExcelFile(excelFilePath);

		// DataTable'ı SQL veritabanına kaydetme
		SaveToDatabase(dataTable, connectionString);

		Console.WriteLine("Veriler başarıyla SQL veritabanına kaydedildi.");
	}

	static DataTable ReadExcelFile(string filePath)
	{
		DataTable dt = new DataTable();

		using (var workbook = new XLWorkbook(filePath))
		{
			var worksheet = workbook.Worksheet(1); // İlk sayfayı seçiyoruz
			bool firstRow = true;

			foreach (var row in worksheet.RowsUsed())
			{
				if (firstRow)
				{
					// Kolonları DataTable'a ekle
					foreach (var cell in row.Cells())
					{
						dt.Columns.Add(cell.Value.ToString());
					}
					firstRow = false;
				}
				else
				{
					// Satırları DataTable'a ekle
					dt.Rows.Add();
					int i = 0;
					foreach (var cell in row.Cells())
					{
						dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
						i++;
					}
				}
			}
		}
		return dt;
	}

	static void SaveToDatabase(DataTable dataTable, string connectionString)
	{
		using (SqlConnection conn = new SqlConnection(connectionString))
		{
			conn.Open();
			using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
			{
				bulkCopy.DestinationTableName = "ILCELER"; // SQL tablosunun adını buraya yazın

				try
				{
					// Verileri SQL tablosuna kopyala
					bulkCopy.WriteToServer(dataTable);
				}
				catch (Exception ex)
				{
					Console.WriteLine("Hata: " + ex.Message);
				}
			}
		}
	}
}
