using System;
using System.Text.Json.Serialization;
using Newtonsoft.Json.Converters;

namespace Aram
{
	public class MyDateTimeConverter : IsoDateTimeConverter
	{
		public MyDateTimeConverter()
		{
			DateTimeFormat = "dd/MM/yyyy HH:mm:ss";
		}
	}

	class Rows
	{
		public String Id { get; set; }

		[JsonConverter(typeof(MyDateTimeConverter))]
		public DateTime start { get; set; }

		[JsonConverter(typeof(MyDateTimeConverter))]
		public DateTime end { get; set; }

		public int weekday { get; set; }
		public int Duration { get; set; }
		public String Proces { get; set; }
		public String Status { get; set; }
		public String started_from { get; set; }
		public String run_on { get; set; }
	}
}
