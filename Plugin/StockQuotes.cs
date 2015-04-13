using System;
using System.Drawing;
using System.Collections.Generic;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rainmeter;

namespace StockQuotes
{
	internal class Measure
	{
		internal string Name;
		internal IntPtr Skin;

		internal string webParserString = "";

		internal string Source;

		List<double> valueOpen = new List<double>();
		List<double> valueClose = new List<double>();
		List<double> valueHigh = new List<double>();
		List<double> valueLow = new List<double>();

		List<int> scaledOpen = new List<int>();
		List<int> scaledClose = new List<int>();
		List<int> scaledHigh = new List<int>();
		List<int> scaledLow = new List<int>();


		int count = 0;
		string clrHigh, clrLow, clrBorder, clrBg;
		Color colorHigh, colorLow, colorBorder, colorBg;

		int ReverseOrder;

		int FromDay, FromMonth, FromYear;
		int ToDay, ToMonth, ToYear;
		string Ticker;

		string imageAddress;
		string Url;
		string CustomPattern;

		int width, height;

		// loads settings
		internal void Reload(Rainmeter.API api, ref double maxValue)
		{
			Name = api.GetMeasureName();
			Skin = api.GetSkin();

			clrBorder = api.ReadString("ColorBorder", "0,0,0,0");
			clrHigh = api.ReadString("ColorHigh", clrBorder);
			clrLow = api.ReadString("ColorLow", "0,0,0,255");
			clrBg = api.ReadString("ColorSolid", "0,0,0,255");

			colorBorder = ColorFromRGBA(clrBorder);
			if (colorBorder == Color.Empty) colorBorder = Color.FromArgb(0, 0, 0, 0);

			colorHigh = ColorFromRGBA(clrHigh);
			if (colorHigh == Color.Empty) colorHigh = colorBorder;

			colorLow = ColorFromRGBA(clrLow);
			if (colorLow == Color.Empty) colorLow = Color.FromArgb(255, 0, 0, 0);

			colorBg = ColorFromRGBA(clrBg);
			if (colorBg == Color.Empty) colorBg = Color.FromArgb(255, 0, 0, 0);

			Source = api.ReadString("Source", "");
			Ticker = api.ReadString("Ticker", "");
			
			width = api.ReadInt("W", 400);
			height = api.ReadInt("H", 300);

			imageAddress = api.ReadString("ImageAddress", "");

			Url = api.ReadString("Url", "");
			CustomPattern = api.ReadString("RegEx", "");

			ReverseOrder = api.ReadInt("ReverseOrder", 1);
		}

		internal virtual void Dispose()
		{
		}

		internal string BuildUrl(string source)
		{
			string url;
			if (string.Compare(source, "moex", true) == 0)
			{
				url = "http://moex.com/iss/history/engines/stock/markets/index/securities/" +
					Ticker +
					".xml?iss.only=history&iss.json=extended&callback=JSON_CALLBACK" +
					"&from=" + FromYear + "-" + FromMonth + "-" + FromDay +
					"&till=" + ToYear + "-" + ToMonth + "-" + ToDay +
					"&limit=100&start=0&sort_order=TRADEDATE&sort_order_desc=desc";
			}
			else if (string.Compare(source, "yahoo", true) == 0)
			{
				url = "http://ichart.yahoo.com/table.csv?s=" +
					Ticker +
					"&a=" + (FromMonth - 1) + "&b=" + FromDay + "&c=" + FromYear +
					"&d=" + (ToMonth - 1) + "&e=" + ToDay + "&f=" + ToYear +
					"&g=d&ignore=.csv";
			}
			else if (string.Compare(source, "finam", true) == 0)
			{
				url = "http://195.128.78.52/" + Ticker + ".csv?market=1&em=16842&code=" + Ticker + "&df=10&mf=2" +
					"&yf=" + FromYear + "from=" + FromDay + "." + FromMonth + "." + FromYear +
					"&dt=10&mt=3" +
					"&yt=" + ToYear + "&to=" + ToDay + "." + ToMonth + "." + ToYear +
					"&p=8&f=" + Ticker + "&e=.csv&cn=" + Ticker + "&dtf=1&tmf=1&MSOR=1&mstime=on&mstimever=1&sep=1&sep2=2&datf=5&at=1";
			}
			else
			{
				API.Log(API.LogType.Debug, "StockQuotes.dll: Source \"" + Source + "\" is not valid.");
				url = string.Empty;
			}
			API.Log(API.LogType.Debug, "StockQuotes.dll: url = " + url);
			return url;
		}

		// converts string from "R,G,B,A" format to .NET "Color" type
		internal Color ColorFromRGBA(string clr)
		{
			string _pattern = "\\s*(?<1>\\d*)\\s*,\\s*(?<2>\\d*)\\s*,\\s*(?<3>\\d*)\\s*,\\s*(?<4>\\d*)\\s*";
			Regex regex = new Regex(_pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
			Match match;
			match = regex.Match(clr);
			if (match.Success)
			{
				int clrR;
				int clrG;
				int clrB;
				int clrA;
				try
				{
					clrR = Int32.Parse(match.Groups[1].ToString());
					clrG = Int32.Parse(match.Groups[2].ToString());
					clrB = Int32.Parse(match.Groups[3].ToString());
					clrA = Int32.Parse(match.Groups[4].ToString());
				}
				catch (Exception)
				{
					API.Log(API.LogType.Error, "StockQuotes: Color \"" + clr + "\" is not valid.");
					return Color.Empty;
				}
				return Color.FromArgb(clrA, clrR, clrG, clrB);
			}
			return Color.Empty;
		}


		internal bool ParseData(string parseString, string source)
		{
			string pattern = "";
			

			if (string.Compare(source, "moex", true) == 0)
			{
				pattern = "CLOSE=\"(?<1>\\S+)\"\\s*" + "OPEN=\"(?<2>\\S+)\"\\s*" + "HIGH=\"(?<31>\\S+)\"\\s*" + "LOW=\"(?<4>\\S+)\"\\s*";
			}
			else if (string.Compare(source, "yahoo", true) == 0)
			{
				pattern = "\\d*-\\d*-\\d*,(?<2>\\d*\\.\\d*),(?<3>\\d*\\.\\d*),(?<4>\\d*\\.\\d*),(?<1>\\d*\\.\\d*)\\s*";
			}
			else if (string.Compare(source, "finam", true) == 0)
			{
				pattern = "\\d*,\\d*,(?<2>\\d*\\.\\d*),(?<3>\\d*\\.\\d*),(?<4>\\d*\\.\\d*),(?<1>\\d*\\.\\d*),\\S+";
			}
			else if (string.Compare(source, "custom", true) == 0)
			{
				pattern = CustomPattern;
			}

			Regex r = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
			Match m;

			string sVal = "";
			double dVal;
			count = 0;
			for (m = r.Match(parseString); m.Success; m = m.NextMatch())
			{
				try
				{
					sVal = m.Groups[1].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
					valueClose.Add(dVal);

					sVal = m.Groups[2].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
					valueOpen.Add(dVal);

					sVal = m.Groups[3].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
					valueHigh.Add(dVal);

					sVal = m.Groups[4].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
					valueLow.Add(dVal);

					// number of given values for each keyword
					count++;
				}
				catch (Exception e)
				{
					API.Log(API.LogType.Debug, "StockQuotes.dll: Parser exception: " + e.Message);
					return false;
				}
			}
			if (count == 0)
			{
				API.Log(API.LogType.Debug, "StockQuotes.dll: Parser (Finam) cant find matches");
				return false;
			}
			return true;
		}

		// scales listed values to pixel-size of candlesticks
		internal void Scale()
		{
			API.Log(API.LogType.Debug, "StockQuotes.dll: Scale");

			scaledOpen = new List<int>();
			scaledClose = new List<int>();
			scaledHigh = new List<int>();
			scaledLow = new List<int>();

			double maxVal = 0.0;
			double minVal = Double.MaxValue;

			foreach (double val in valueLow)
			{
				if (val < minVal) minVal = val;
			}

			foreach (double val in valueHigh)
			{
				if (val > maxVal) maxVal = val;
			}

			double k = Convert.ToDouble(height) / (maxVal - minVal);

			foreach (double val in valueOpen)
			{
				scaledOpen.Add(Convert.ToInt32((val - minVal) * k));
			}

			foreach (double val in valueClose)
			{
				scaledClose.Add(Convert.ToInt32((val - minVal) * k));
			}

			foreach (double val in valueHigh)
			{
				scaledHigh.Add(Convert.ToInt32((val - minVal) * k));
			}

			foreach (double val in valueLow)
			{
				scaledLow.Add(Convert.ToInt32((val - minVal) * k));
			}
		}

		internal Measure()
		{
			API.Log(API.LogType.Debug, "StockQuotes.dll: Measure constructor called");
			//Reload();
			//Update();
			//ParseData("");
		}

		// downloads Url
		internal string GetQuotes(string url)
		{
			string result;
			using (var client = new WebClient())
			{
				try
				{
					result = client.DownloadString(Url);
					return result;
				}
				catch (Exception e)
				{
					API.Log(API.LogType.Debug, "StockQuotes.dll: WebClient exception: " + e.Message);
				}
			}
			return string.Empty;
		}


		// draws a graph, makes a picture and returns its full name
		internal string DrawGraph()
		{
			//API.Log(API.LogType.Debug, "StockQuotes.dll: DrawGraph");

			int _width = 0;
			if (count != 0)
			{
				_width = width / count;
			}

			int _height = height;
			int _elementWidth = Convert.ToInt32(_width * 0.8);

			if ((_width - _elementWidth) % 2 == 0)
			{
				_elementWidth--;
			}

			int _padding = (_width - _elementWidth) / 2;

			Pen pen = new Pen(colorBorder, 1);
			Point p1 = new Point();
			Point p2 = new Point();

			using (Image image = new Bitmap(width, height))
			{
				Graphics graphics = Graphics.FromImage(image);
				graphics.FillRectangle(new SolidBrush(colorBg), new Rectangle(0, 0, width, height));

				for (int index = 0; index < count; index++)
				{
					Image img = new Bitmap(_width, _height);
					Graphics _singleGraphics = Graphics.FromImage(img);
					SolidBrush brush;

					if (scaledOpen[index] < scaledClose[index])
					{
						p1.X = p2.X = _padding + _elementWidth / 2;
						p1.Y = _height - scaledHigh[index];
						p2.Y = _height - scaledClose[index];
						_singleGraphics.DrawLine(pen, p1, p2);
						p1.Y = _height - scaledOpen[index];
						p2.Y = _height - scaledLow[index];
						_singleGraphics.DrawLine(pen, p1, p2);

						p1.X = _padding;
						p2.X = _elementWidth;
						pen.Color = colorBorder;
						brush = new SolidBrush(colorHigh);
						p1.Y = _height - scaledClose[index];
						p2.Y = scaledClose[index] - scaledOpen[index];
						if (p2.Y == 0) p2.Y = 1;
						_singleGraphics.DrawRectangle(pen, p1.X, p1.Y, p2.X, p2.Y);
						_singleGraphics.FillRectangle(brush, p1.X + 1, p1.Y + 1, p2.X - 1, p2.Y - 1);
						//API.Log(API.LogType.Debug, "StockQuotes.dll: Rect draw");
					}
					else
					{
						p1.X = p2.X = _padding + _elementWidth / 2;
						p1.Y = _height - scaledHigh[index];
						p2.Y = _height - scaledOpen[index];
						_singleGraphics.DrawLine(pen, p1, p2);
						p1.Y = _height - scaledClose[index];
						p2.Y = _height - scaledLow[index];
						_singleGraphics.DrawLine(pen, p1, p2);

						p1.X = _padding;
						p2.X = _elementWidth;
						pen.Color = colorBorder;
						brush = new SolidBrush(colorLow);
						p1.Y = _height - scaledOpen[index];
						p2.Y = scaledOpen[index] - scaledClose[index];
						if (p2.Y == 0) p2.Y = 1;
						_singleGraphics.DrawRectangle(pen, p1.X, p1.Y, p2.X, p2.Y);
						_singleGraphics.FillRectangle(brush, p1.X + 1, p1.Y + 1, p2.X - 1, p2.Y - 1);
						//API.Log(API.LogType.Debug, "StockQuotes.dll: Rect draw");
					}
					if (ReverseOrder != 0)
					{
						graphics.DrawImage(img, new Point(_width * (count - index - 1), 0));
					}
					else
					{
						graphics.DrawImage(img, new Point(_width * index, 0));
					}

				}
				image.Save(imageAddress, System.Drawing.Imaging.ImageFormat.Png);
			}

			//API.Log(API.LogType.Debug, "StockQuotes.dll: DrawGraph = " + imageAddress);
			return imageAddress;
		}

		internal void GetValue()
		{
			//API.Log(API.LogType.Debug, "StockQuotes.dll: GetValue");
			return;
		}

		internal double Update()
		{
			//API.Log(API.LogType.Debug, "StockQuotes.dll: Update");

			DateTime nowDate = DateTime.Now;

			ToDay = nowDate.Day;
			ToMonth = nowDate.Month;
			ToYear = nowDate.Year;

			FromDay = nowDate.Day;
			FromMonth = nowDate.AddMonths(-1).Month;
			FromYear = nowDate.Year;

			Url = BuildUrl(Source);

			webParserString = GetQuotes(Url);

			ParseData(webParserString, Source);

			Scale();

			imageAddress = DrawGraph();

			return 0.0;
		}

		internal string GetString()
		{

			return imageAddress;
		}
	}

	public static class Plugin
	{

		static IntPtr StringBuffer = IntPtr.Zero;

		[DllExport]
		public static void Initialize(ref IntPtr data, IntPtr rm)
		{
			Rainmeter.API api = new Rainmeter.API(rm);
			Measure measure = new Measure();
			data = GCHandle.ToIntPtr(GCHandle.Alloc(measure));
		}

		[DllExport]
		public static void Finalize(IntPtr data)
		{
			Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
			measure.Dispose();
			GCHandle.FromIntPtr(data).Free();


			if (StringBuffer != IntPtr.Zero)
			{
				Marshal.FreeHGlobal(StringBuffer);
				StringBuffer = IntPtr.Zero;
			}
		}

		[DllExport]
		public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
		{
			Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
			measure.Reload(new Rainmeter.API(rm), ref maxValue);
		}

		[DllExport]
		public static double Update(IntPtr data)
		{
			Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
			return measure.Update();
		}


		[DllExport]
		public static IntPtr GetString(IntPtr data)
		{
			Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;

			if (StringBuffer != IntPtr.Zero)
			{
				Marshal.FreeHGlobal(StringBuffer);
				StringBuffer = IntPtr.Zero;
			}

			string stringValue = measure.GetString();
			if (stringValue != null)
			{
				StringBuffer = Marshal.StringToHGlobalUni(stringValue);
			}

			return StringBuffer;
		}
	}
}

