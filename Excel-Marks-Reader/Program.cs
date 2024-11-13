using System;
using System.Text;
using Aspose.Cells;

class TestExcel
{
	struct Coordinate
	{
		public int column;
		public int row;

		public Coordinate(int column, int row)
		{
			this.column = column;
			this.row = row;
		}
	}

	static void Main(string[] args)
	{
		try
		{
			Console.InputEncoding = UTF8Encoding.UTF8;
			Console.OutputEncoding = UTF8Encoding.UTF8;

			Console.WriteLine("Hello, world!");
			try
			{
				StreamReader outputPathTxt = new StreamReader(Path.Combine(Directory.GetCurrentDirectory(), "OutputPath.txt"));
				string outputPath = outputPathTxt.ReadLine().Replace(@"\", @"\\");
				outputPathTxt.Close();

				FileStream file = new FileStream(Path.Combine(outputPath, "file.xlsx"), FileMode.Open);
				Workbook workbook = new Workbook(file);
				Worksheet worksheet = workbook.Worksheets[0];

				Coordinate start = new Coordinate(1 - 1, 10 - 1),
				   end = new Coordinate(19, 35);


				for (int row = start.row + 2; row < end.row; row++)
				{
					try
					{
						string studentName = worksheet.Cells[row, start.column].Value.ToString();
						FileStream student = new FileStream(Path.Combine(outputPath, studentName + ".txt"), FileMode.Create);
						student.Write(new UTF8Encoding(true).GetBytes($"Оцінки {studentName}.\n\n"));
						List<double> marks = new List<double>();
						for (int column = start.column; column < end.column; column++)
						{
							if (worksheet.Cells[start.row, column].Type != CellValueType.IsNull)
							{
								Cell mark = worksheet.Cells[row, column],
									lesson = worksheet.Cells[start.row, column];
								if (mark.IsNumericValue)
									marks.Add(mark.FloatValue);
								else
									mark.Value = "-";
								student.Write(new UTF8Encoding(true).GetBytes($"    {lesson.Value}: {mark.Value}\n"));
							}
						}
						student.Write(new UTF8Encoding(true).GetBytes($"\nСередній бал: {(marks.Sum() / marks.Count): 0.00}.\n\n"));
						student.Write(new UTF8Encoding(true).GetBytes("Позначення:" +
							"    \n\"з.в.\" - Звільнен(-а) з предмету;" +
							"    \n\"н.а.\" - Не атестован(-а). Занадто низький бал;" +
							"    \n\"-\" - Оцінку не виставлено (з невідомої причини).\n"));
						student.Close();
						Console.WriteLine($"Оценки ученика {studentName} записаны!");
					}
					catch (Exception exception)
					{
						Console.WriteLine(exception);
					}

				}
				file.Close();
			}
			catch (Exception exception)
			{
				Console.WriteLine(exception);
			}
			Console.ReadKey();
		}
		catch (Exception exception)
		{
			Console.WriteLine(exception);
		}
		Console.ReadKey();
	}
}
