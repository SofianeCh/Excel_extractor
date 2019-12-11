using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Nest;
using Elasticsearch.Net;

namespace Aram
{
	class Program
	{
		/// <summary>
		/// We open the xlsx file in the desktop and we'll work colum by colum. Each row will be put in an object that will be inserted into elasticsearch.
		/// </summary>
		/// <param name="args"></param>
		static void Main(string[] args)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

			var filePath = @"C:\Users\sofiane\Desktop\Aram.xlsx";
			using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
			{
				Uri node = new Uri("http://localhost:9200");
				SingleNodeConnectionPool pool = new SingleNodeConnectionPool(node);
				var config = new ConnectionConfiguration(pool);
				var settings = new ConnectionSettings(node)
					.DefaultIndex("aram");
				var clientP = new ElasticClient(new ConnectionSettings(pool));
				ElasticClient client = new ElasticClient(settings);

				using (var reader = ExcelReaderFactory.CreateReader(stream))
				{
					// Use the AsDataSet extension method
					DataSet result = reader.AsDataSet();
					var rowlist = new List<Rows>();
					int id = 0;

					// The result of each spreadsheet is in result.Tables
					Fill_object(result, 0, client);
					Console.WriteLine();
				}
				Console.WriteLine("Data extracted");
			}
		}
		/// <summary>
		/// Parse evry field in the excel colum and put them one by one in the appropriate property.
		/// </summary>
		/// 
		/// 
		/// --> DateTime to check !!!!!
		/// 
		/// 
		/// <returns></returns>

		public static int Fill_object(DataSet result, int id, ElasticClient client)
		{
			foreach (DataTable table in result.Tables)
			{
				foreach (DataRow row in table.Rows)
				{
					int colom = 0;
					Rows romObj = new Rows();
					romObj.Id = id.ToString();
					foreach (DataColumn column in table.Columns)
					{
						try
						{
							if (id == 0)
							{
								DateTime value = new DateTime(2018, 4, 1);
								romObj.start = value;
								romObj.end = value;

							}
							if (colom == 0 && id > 0)
							{
								var time = row[colom];
								romObj.start = (DateTime)row[colom];
								romObj.weekday = (int)romObj.start.DayOfWeek;
							}
							try
							{
								if (colom == 1 && id > 0)
								{
									romObj.end = (DateTime)row[colom];
									TimeSpan duration = romObj.end - romObj.start;
									romObj.Duration = (int)duration.TotalSeconds;
								}
							}
							catch
							{
								Console.WriteLine("No end date for this row");
							}
							if (colom == 2 && id > 0)
							{
								romObj.Proces = (string)row[colom];
								int index = romObj.Proces.IndexOf('-');
								int lastindexof = romObj.Proces.LastIndexOf('-');
								if (lastindexof > 0)
								{
									romObj.Proces = romObj.Proces.Remove(lastindexof);
									if (index != lastindexof)
										romObj.Proces = romObj.Proces.Substring(index + 2);
								}
							}
							if (colom == 3 && id > 0)
								romObj.Status = (string)row[colom];
							if (colom == 4 && id > 0)
							{
								romObj.started_from = (string)row[colom];
								if (romObj.started_from.Contains("_debug"))
									romObj.started_from = romObj.started_from.Remove(romObj.started_from.Length - 6);
							}
							if (colom == 5 && id > 0)
							{
								romObj.run_on = (string)row[colom];
								if (romObj.run_on.Contains("_debug"))
									romObj.run_on = romObj.run_on.Remove(romObj.run_on.Length - 6);

							}
							//Console.WriteLine(row[column] + "   " + colom);
							colom++;
						}
						catch (Exception ex)
						{
							Console.WriteLine(ex);
							return (0);
						}
					}
					id++;
					colom = 0;
					Console.WriteLine(romObj.Id + "  " + romObj.Proces + "  " + romObj.Status);
					var indexResponse1 = client.IndexDocument(romObj);
					//rowlist.Add(romObj);
				}
			}
			return (1);
		}
	}
}
