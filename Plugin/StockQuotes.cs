//#define DEBUG

using Rainmeter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace StockQuotes
{
//	internal abstract class Stock
//	{
//		internal string Ticker { get; set; }

//		//internal DateTime nowDate = DateTime.Now;
//		internal int FromDay, FromMonth, FromYear;
//		internal int ToDay, ToMonth, ToYear;

//		internal List<double> ValueOpen = new List<double>();
//		internal List<double> ValueClose = new List<double>();
//		internal List<double> ValueHigh = new List<double>();
//		internal List<double> ValueLow = new List<double>();

//		internal List<int> ScaledOpen = new List<int>();
//		internal List<int> ScaledClose = new List<int>();
//		internal List<int> ScaledHigh = new List<int>();
//		internal List<int> ScaledLow = new List<int>();

//		private string GetQuotes(string url)
//		{
//			string result = string.Empty;
//			using (var client = new WebClient())
//			{
//				try
//				{
//					result = client.DownloadString(url);
//				}
//				catch (Exception e)
//				{
//					API.Log(API.LogType.Debug, "StockQuotes.dll: WebClient exception: " + e.Message);
//				}
////#if DEBUG
////				finally
////				{
////					// save downloaded data to txt
////					using (FileStream fs = new FileStream(ImageAddress + ".txt", FileMode.Create))
////					{
////						byte[] buffer = new byte[result.Length * sizeof(char)];
////						System.Buffer.BlockCopy(result.ToCharArray(), 0, buffer, 0, buffer.Length);
////						fs.Write(buffer, 0, buffer.Length);
////					}
////				}
////#endif
//			}
//			return result;
//		}

//		internal Stock()
//		{

//		}

//		internal Stock(string ticker)
//		{

//		}

//		private abstract bool ParseQuotes(string quotes)
//		{
//			return false;
//		}

//		internal void GetValues()
//		{
//			ToDay = DateTime.Now.Day;
//			ToMonth = DateTime.Now.Month;
//			ToYear = DateTime.Now.Year;
//			FromDay = DateTime.Now.Day;
//			FromMonth = DateTime.Now.AddMonths(-1).Month;
//			FromYear = DateTime.Now.Year;

//			string url = this.BuildUrl();

//			string webQuotes = this.GetQuotes(url);

//			ParseQuotes(webQuotes);

//			Scale();
//		}

//		private abstract string BuildUrl();


//		private void Scale()
//		{
//			ScaledOpen.Clear();
//			ScaledClose.Clear();
//			ScaledHigh.Clear();
//			ScaledLow.Clear();

//			double maxVal = Double.MinValue;
//			double minVal = Double.MaxValue;

//			foreach (double val in ValueLow)
//			{
//				if (val < minVal) minVal = val;
//			}

//			foreach (double val in ValueHigh)
//			{
//				if (val > maxVal) maxVal = val;
//			}

//			double k = Convert.ToDouble(Height) / (maxVal - minVal);

//			foreach (double val in ValueOpen)
//			{
//				ScaledOpen.Add(Convert.ToInt32((val - minVal) * k));
//			}

//			foreach (double val in ValueClose)
//			{
//				ScaledClose.Add(Convert.ToInt32((val - minVal) * k));
//			}

//			foreach (double val in ValueHigh)
//			{
//				ScaledHigh.Add(Convert.ToInt32((val - minVal) * k));
//			}

//			foreach (double val in ValueLow)
//			{
//				ScaledLow.Add(Convert.ToInt32((val - minVal) * k));
//			}
//		}
//	}

//	internal class StockFinam : Stock
//	{
//		internal int Id { get; set; }
//		internal int MarketId { get; set; }
//		internal StockFinam(string ticker)
//		{
//			Ticker = ticker;
//			Id = Finam.GetId(ticker);
//			MarketId = Finam.GetMarketId(ticker);
//		}

//		internal override string BuildUrl()
//		{
//			string result = "http://195.128.78.52/" + Ticker + ".csv?" +
//					"market=" + MarketId +
//					"&em=" + Id +
//					"&code=" + Ticker +
//					"&df=" + FromDay +
//					"&mf=" + (FromMonth - 1) +
//					"&yf=" + FromYear +
//					"&from=" + FromDay + "." + FromMonth + "." + FromYear +
//					"&dt=" + ToDay +
//					"&mt=" + (ToMonth - 1) +
//					"&yt=" + ToYear +
//					"&to=" + ToDay + "." + ToMonth + "." + ToYear +
//					"&p=8" + // period: 8 for one-day ticks
//					"&f=" + Ticker + "&e=.csv" +
//					"&cn=" + Ticker +
//					"&dtf=1&tmf=1&MSOR=1&mstime=on&mstimever=1&sep=1&sep2=2&datf=5&at=1";
//			//result = String.Format("http://195.128.78.52/" + "{0}.{1}?d=d&market={2}&em={3}&p={4}&df={5}&mf={6}&yf={7}&dt={8}&mt={9}&yt={10}&f={11}&e=.{12}&datf={13}&cn={14}&dtf=1&tmf=1&MSOR=0&sep=3&sep2=1&at=1",
//			//	Ticker,					// 0
//			//	"csv",					// 1
//			//	MarketId,				// 2
//			//	Id,						// 3
//			//	1,//settings.period,	// 4
//			//	FromDay,				// 5
//			//	FromMonth - 1,			// 6
//			//	FromYear,				// 7
//			//	ToDay,					// 8
//			//	ToMonth - 1,			// 9
//			//	ToYear,					// 10
//			//	Ticker,					// 11
//			//	"csv",					// 12
//			//	11,//format, if (settings.period == 1) format = 11; else format = 5;
//			//	Ticker					// 14
//			//	);
//			return result;
//		}

//		private override bool ParseQuotes(string quotes)
//		{
//			string pattern = "\\d*,\\d*,(?<2>\\d*\\.\\d*),(?<3>\\d*\\.\\d*),(?<4>\\d*\\.\\d*),(?<1>\\d*\\.\\d*),\\S+";
//			ValueClose.Clear();
//			ValueOpen.Clear();
//			ValueHigh.Clear();
//			ValueLow.Clear();
//			//Outdated = false;

//			Regex r = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
//			Match m;

//			double dVal;

//			for (m = r.Match(quotes); m.Success; m = m.NextMatch())
//			{
//				try
//				{
//					dVal = Double.Parse(m.Groups[1].ToString(), System.Globalization.CultureInfo.InvariantCulture);
//					ValueClose.Add(dVal);

//					dVal = Double.Parse(m.Groups[2].ToString(), System.Globalization.CultureInfo.InvariantCulture);
//					ValueOpen.Add(dVal);

//					dVal = Double.Parse(m.Groups[3].ToString(), System.Globalization.CultureInfo.InvariantCulture);
//					ValueHigh.Add(dVal);

//					dVal = Double.Parse(m.Groups[4].ToString(), System.Globalization.CultureInfo.InvariantCulture);
//					ValueLow.Add(dVal);
//				}
//				catch (Exception e)
//				{
//					API.Log(API.LogType.Debug, "StockQuotes.dll: Parser exception: " + e.Message);
//					ValueClose.Clear();
//					ValueOpen.Clear();
//					ValueHigh.Clear();
//					ValueLow.Clear();
//					return false;
//				}
//			}
//			if (ValueClose.Capacity == 0)
//			{
//				API.Log(API.LogType.Debug, "StockQuotes.dll: Parser cant find matches");
//				return false;
//			}
//			return true;
//		}
//	}

//	internal class StockYahoo : Stock
//	{
//		internal StockYahoo(string ticker)
//		{

//		}

//		internal override string BuildUrl()
//		{
//			string result = "http://ichart.yahoo.com/table.csv?s=" +
//							Ticker +
//							"&a=" + (FromMonth - 1) + "&b=" + FromDay + "&c=" + FromYear +
//							"&d=" + (ToMonth - 1) + "&e=" + ToDay + "&f=" + ToYear +
//							"&g=d&ignore=.csv";
//			return result;
//		}
//	}

	internal class Measure
	{
		#region InnerProperties
		internal string WebParserString = "";

		internal List<double> ValueOpen = new List<double>();
		internal List<double> ValueClose = new List<double>();
		internal List<double> ValueHigh = new List<double>();
		internal List<double> ValueLow = new List<double>();

		internal List<int> ScaledOpen = new List<int>();
		internal List<int> ScaledClose = new List<int>();
		internal List<int> ScaledHigh = new List<int>();
		internal List<int> ScaledLow = new List<int>();

		internal int Count = 0;
		internal string Url;

		internal bool Outdated = false;
		#endregion

		#region ConfigProperties
		internal string Name;
		internal IntPtr Skin;
		internal Color ColorHigh, ColorLow, ColorBorder, ColorBackground;
		internal string Source;
		internal string Ticker;
		internal int Width, Height;
		internal string ImageAddress;

		//internal string CustomRegEx;
		internal int ReverseOrder;
		#endregion

		// just constructor
		internal Measure()
		{
#if DEBUG
			API.Log(API.LogType.Debug, "StockQuotes.dll: Measure constructor called");
#endif
		}

		// just destructor
		internal void Dispose()
		{
#if DEBUG
			API.Log(API.LogType.Debug, "StockQuotes.dll: Measure dnstructor called");
#endif
		}

		// loading settings from ini file
		internal void Reload(Rainmeter.API api, ref double maxValue)
		{
			Name = api.GetMeasureName();
			Skin = api.GetSkin();
#if DEBUG
			API.Log(API.LogType.Debug, "StockQuotes.dll: Name: " + Name);
#endif
			string colorHigh, colorLow, colorBorder, colorBackground;
			colorBorder = api.ReadString("ColorBorder", "0,0,0,255");
			colorHigh = api.ReadString("ColorHigh", "255,255,255,255");
			colorLow = api.ReadString("ColorLow", "0,0,0,255");
			colorBackground = api.ReadString("ColorSolid", "255,255,255,25");

			ColorBorder = ColorFromRGBA(colorBorder);
			if (ColorBorder == Color.Empty) ColorBorder = Color.FromArgb(0, 0, 0, 0);

			ColorHigh = ColorFromRGBA(colorHigh);
			if (ColorHigh == Color.Empty) ColorHigh = ColorBorder;

			ColorLow = ColorFromRGBA(colorLow);
			if (ColorLow == Color.Empty) ColorLow = Color.FromArgb(255, 0, 0, 0);

			ColorBackground = ColorFromRGBA(colorBackground);
			if (ColorBackground == Color.Empty) ColorBackground = Color.FromArgb(255, 0, 0, 0);

			Source = api.ReadString("Source", "");
			Ticker = api.ReadString("Ticker", "");

			Width = api.ReadInt("W", 400);
			Height = api.ReadInt("H", 300);

			ImageAddress = api.ReadString("ImageAddress", "") + @"\" + Name + ".png";

			ReverseOrder = api.ReadInt("ReverseOrder", 1);

			Outdated = false;
		}

		//internal string DownloadFinam(StockFinam stock)
		//{

		//}

		//internal string DownloadYahoo(StockYahoo stock)
		//{

		//}

		// building (or rebuilding) Url with current dates and ticker
		internal string BuildUrl(string source)
		{
			string url;
			int FromDay, FromMonth, FromYear;
			int ToDay, ToMonth, ToYear;
			DateTime nowDate = DateTime.Now;
			ToDay = nowDate.Day;
			ToMonth = nowDate.Month;
			ToYear = nowDate.Year;
			FromDay = nowDate.AddMonths(-1).Day;
			FromMonth = nowDate.AddMonths(-1).Month;
			FromYear = nowDate.AddMonths(-1).Year;

			if (string.Compare(source, "moex", true) == 0)
			{
				url = "http://www.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/" +
					Ticker +
					".csv" +
					"?from=" + FromYear + "-" + FromMonth + "-" + FromDay +
					"&till=" + ToYear + "-" + ToMonth + "-" + ToDay +
					"&lang=RU";
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
				url = "http://195.128.78.52/" + Ticker + ".csv?" +
					"market=" + Finam.GetMarketId(Ticker) +
					"&em=" + Finam.GetId(Ticker) +
					"&code=" + Ticker +
					"&df=" + FromDay +
					"&mf=" + (FromMonth - 1) +
					"&yf=" + FromYear +
					"&from=" + FromDay + "." + FromMonth + "." + FromYear +
					"&dt=" + ToDay +
					"&mt=" + (ToMonth - 1) +
					"&yt=" + ToYear +
					"&to=" + ToDay + "." + ToMonth + "." + ToYear +
					"&p=8" + // period: 8 for one-day ticks
					"&f=" + Ticker + "&e=.csv" +
					"&cn=" + Ticker +
					"&dtf=1&tmf=1&MSOR=1&mstime=on&mstimever=1&sep=1&sep2=1&datf=5&at=1";
			}
			else
			{
				API.Log(API.LogType.Error, "StockQuotes.dll: Source \"" + Source + "\" is not valid.");
				url = string.Empty;
			}
#if DEBUG
			API.Log(API.LogType.Debug, "StockQuotes.dll: url = " + url);
#endif
			return url;
		}

		// converting string from "R,G,B,A" format to .NET "Color" type
		internal Color ColorFromRGBA(string clr)
		{
			string _pattern = "\\s*(?<1>\\d*)\\s*,\\s*(?<2>\\d*)\\s*,\\s*(?<3>\\d*)\\s*,\\s*(?<4>\\d*)\\s*";
			Regex regex = new Regex(_pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
			Match match;
			match = regex.Match(clr);
			if (match.Success)
			{
				short clrR;
				short clrG;
				short clrB;
				short clrA;
				try
				{
					clrR = short.Parse(match.Groups[1].ToString());
					clrG = short.Parse(match.Groups[2].ToString());
					clrB = short.Parse(match.Groups[3].ToString());
					clrA = short.Parse(match.Groups[4].ToString());
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
			ValueClose.Clear();
			ValueOpen.Clear();
			ValueHigh.Clear();
			ValueLow.Clear();

			Count = 0;

			// index of column with corresponding data
			int OpenIdx = 0;
			int CloseIdx = 0;
			int HighIdx = 0;
			int LowIdx = 0;

			string[] separator = { ";" };
			//System.Globalization.NumberFormatInfo nfi = new System.Globalization.CultureInfo("en-US", false).NumberFormat;
			//nfi.NumberDecimalSeparator = ".";

			if (string.Compare(source, "moex", true) == 0)
			{
				//0      ;1        ;2        ;3    ;4        ;5    ;6   ;7  ;8   ;9
				//BOARDID;TRADEDATE;SHORTNAME;SECID;NUMTRADES;VALUE;OPEN;LOW;HIGH;LEGALCLOSEPRICE
				OpenIdx = 6;
				CloseIdx = 9;
				HighIdx = 8;
				LowIdx = 7;

				separator[0] = ";";
			}
			else if (string.Compare(source, "yahoo", true) == 0)
			{
				//0   ;1   ;2   ;3  ;4    ;5     ;6
				//Date,Open,High,Low,Close,Volume,Adj Close
				OpenIdx = 1;
				CloseIdx = 4;
				HighIdx = 2;
				LowIdx = 3;

				separator[0] = ",";
			}
			else if (string.Compare(source, "finam", true) == 0)
			{
				//0     ;1     ;2     ;3     ;4    ;5      ;6
				//<DATE>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOL>
				OpenIdx = 2;
				CloseIdx = 5;
				HighIdx = 3;
				LowIdx = 4;

				separator[0] = ",";
			}
			
			

			using (TextReader csvFile = new StringReader(parseString))
			{
				// skip header
				string line = csvFile.ReadLine();
				if (string.Compare(source, "moex", true) == 0)
				{
					line = csvFile.ReadLine();
					line = csvFile.ReadLine();
				}

				while ((line = csvFile.ReadLine()) != null)
				{
					string[] attr = line.Split(separator, StringSplitOptions.None);
					try
					{
						ValueOpen.Add(Double.Parse(attr[OpenIdx], System.Globalization.CultureInfo.InvariantCulture));
						ValueLow.Add(Double.Parse(attr[LowIdx], System.Globalization.CultureInfo.InvariantCulture));
						ValueHigh.Add(Double.Parse(attr[HighIdx], System.Globalization.CultureInfo.InvariantCulture));
						ValueClose.Add(Double.Parse(attr[CloseIdx], System.Globalization.CultureInfo.InvariantCulture));
						Count++;

#if DEBUG
						API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " Open = " + Double.Parse(attr[OpenIdx], System.Globalization.CultureInfo.InvariantCulture));
						API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " Close = " + Double.Parse(attr[CloseIdx], System.Globalization.CultureInfo.InvariantCulture));
						API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " High = " + Double.Parse(attr[HighIdx], System.Globalization.CultureInfo.InvariantCulture));
						API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " Low = " + Double.Parse(attr[LowIdx], System.Globalization.CultureInfo.InvariantCulture));
						API.Log(API.LogType.Debug, "StockQuotes: Count = " + Count);
#endif
					}
					catch (Exception e)
					{
						API.Log(API.LogType.Warning, "StockQuotes.dll: Double.Parse exception: " + e.Message);
						//return false;
					}
				}
				if (Count == 0)
				{
					API.Log(API.LogType.Warning, "StockQuotes.dll: (" + Ticker + ") Parser cant find matches");
					return false;
				}
			}
			return true;
		}

		// just parsing values from the string we got with GetQuotes()
		internal bool ParseDataOld(string parseString, string source)
		{
			string pattern = string.Empty;
			ValueClose.Clear();
			ValueOpen.Clear();
			ValueHigh.Clear();
			ValueLow.Clear();
			//Outdated = false;

			if (string.Compare(source, "moex", true) == 0)
			{
				//pattern = "CLOSE=\"(?<1>\\S+)\"\\s*" + "OPEN=\"(?<2>\\S+)\"\\s*" + "HIGH=\"(?<3>\\S+)\"\\s*" + "LOW=\"(?<4>\\S+)\"\\s*";
				pattern = @"\w*;\d{4}-\d{2}-\d{2};\S*\s*\S*;\S+;\d*;\d*\.?\d*;(?<2>\d*\.?\d*);(?<4>\d*\.?\d*);(?<3>\d*\.?\d*);(?<1>\d*\.?\d*)";
			}
			else if (string.Compare(source, "yahoo", true) == 0)
			{
				pattern = @"\d*-\d*-\d*,(?<2>\d*\.\d*),(?<3>\d*\.\d*),(?<4>\d*\.\d*),(?<1>\d*\.\d*)\s*";
			}
			else if (string.Compare(source, "finam", true) == 0)
			{
				pattern = @"\d*,\d*,(?<2>\d*\.?\d*\.\d{7}),(?<3>\d*\.?\d*\.\d{7}),(?<4>\d*\.?\d*\.\d{7}),(?<1>\d*\.?\d*\.\d{7}),\s*";
			}

			Regex r = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
			Match m;

			string sVal = "";
			double dVal;
			Count = 0;
			for (m = r.Match(parseString); m.Success; m = m.NextMatch())
			{
				try
				{
					sVal = m.Groups[1].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
#if DEBUG
					API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " dVal = " + dVal);
#endif
					ValueClose.Add(dVal);

					sVal = m.Groups[2].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
#if DEBUG
					API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " dVal = " + dVal);
#endif
					ValueOpen.Add(dVal);

					sVal = m.Groups[3].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
#if DEBUG
					API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " dVal = " + dVal);
#endif
					ValueHigh.Add(dVal);

					sVal = m.Groups[4].ToString();
					dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);
#if DEBUG
					API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " dVal = " + dVal);
#endif
					ValueLow.Add(dVal);

					// number of given values for each keyword
					Count++;
#if DEBUG
					API.Log(API.LogType.Debug, "StockQuotes: Count = " + Count);
#endif
				}
				catch (Exception e)
				{
					Count = 0;
					API.Log(API.LogType.Error, "StockQuotes.dll: (" + Ticker + ") Parser exception: " + e.Message);
					return false;
				}
			}
			if (Count == 0)
			{
				API.Log(API.LogType.Warning, "StockQuotes.dll: (" + Ticker + ") Parser cant find matches");
				return false;
			}
			return true;
		}

		// scaling listed values to pixel-size of candlesticks
		internal void Scale()
		{
			ScaledOpen.Clear();
			ScaledClose.Clear();
			ScaledHigh.Clear();
			ScaledLow.Clear();

			double maxVal = Double.MinValue;
			double minVal = Double.MaxValue;

			foreach (double val in ValueLow)
			{
				if (val < minVal) minVal = val;
			}

			foreach (double val in ValueHigh)
			{
				if (val > maxVal) maxVal = val;
			}
#if DEBUG
			API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " minVal = " + minVal);
			API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " maxVal = " + maxVal);
#endif

			double k = Convert.ToDouble(Height) / (maxVal - minVal);

			foreach (double val in ValueOpen)
			{
				ScaledOpen.Add(Convert.ToInt32((val - minVal) * k));
			}

			foreach (double val in ValueClose)
			{
				ScaledClose.Add(Convert.ToInt32((val - minVal) * k));
			}

			foreach (double val in ValueHigh)
			{
				ScaledHigh.Add(Convert.ToInt32((val - minVal) * k));
			}

			foreach (double val in ValueLow)
			{
				ScaledLow.Add(Convert.ToInt32((val - minVal) * k));
			}
#if DEBUG
			API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " ScaledOpen = " + ScaledOpen.Count);
			API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " ScaledClose = " + ScaledClose.Count);
			API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " ScaledHigh = " + ScaledHigh.Count);
			API.Log(API.LogType.Debug, "StockQuotes: " + Ticker + " ScaledLow = " + ScaledLow.Count);
#endif
		}

		// downloading Url
		internal string GetQuotes(string url)
		{
			string result = string.Empty;
			using (var client = new WebClient())
			{
				try
				{
					result = client.DownloadString(Url);
				}
				catch (Exception e)
				{
					API.Log(API.LogType.Error, "StockQuotes.dll: WebClient exception: " + e.Message);
				}
#if DEBUG
				finally
				{
					// save downloaded data to txt
					using (FileStream fs = new FileStream(ImageAddress + ".txt", FileMode.Create))
					{
						byte[] buffer = new byte[result.Length * sizeof(char)];
						System.Buffer.BlockCopy(result.ToCharArray(), 0, buffer, 0, buffer.Length);
						fs.Write(buffer, 0, buffer.Length);
					}
				}
#endif
			}
			return result;
		}

		// marking a graph if we can't download new values with GetQuotes()
		internal void MarkOutdated()
		{
			try
			{
				MemoryStream ms = new MemoryStream();
				using (FileStream fs = new FileStream(ImageAddress, FileMode.Open))
				{
					byte[] buffer = new byte[fs.Length];
					fs.Read(buffer, 0, buffer.Length);
					ms.Write(buffer, 0, buffer.Length);
				}
				Image imageOld = new Bitmap(ms);
				Graphics graphics = Graphics.FromImage(imageOld);

				// mark here

				imageOld.Save(ImageAddress, System.Drawing.Imaging.ImageFormat.Png);
				Outdated = true;

			}
			catch (Exception e)
			{
				API.Log(API.LogType.Error, "StockQuotes.dll: MarkOutdated exception: " + e.ToString());
			}

		}

		// drawing a graph, makes a picture and returns its full name
		internal void DrawGraph()
		{
			int _width = 0;
			if (Count != 0)
			{
				_width = Width / Count;
			}

			int _height = Height;
			int _elementWidth = Convert.ToInt32(_width * 0.8);

			if ((_width - _elementWidth) % 2 == 0)
			{
				_elementWidth--;
			}

			int _padding = (_width - _elementWidth) / 2;

			Pen pen = new Pen(ColorBorder, 1);
			Point p1 = new Point();
			Point p2 = new Point();

			using (Image image = new Bitmap(Width, Height))
			{

				Graphics graphics = Graphics.FromImage(image);
				graphics.FillRectangle(new SolidBrush(ColorBackground), new Rectangle(0, 0, Width, Height));

				for (int index = 0; index < Count; index++)
				{
					using (Image img = new Bitmap(_width, _height))
					{
						Graphics singleGraphics = Graphics.FromImage(img);
						SolidBrush brush;

						if (ScaledOpen[index] < ScaledClose[index])
						{
							p1.X = p2.X = _padding + _elementWidth / 2;
							p1.Y = _height - ScaledHigh[index];
							p2.Y = _height - ScaledClose[index];
							singleGraphics.DrawLine(pen, p1, p2);
							p1.Y = _height - ScaledOpen[index];
							p2.Y = _height - ScaledLow[index];
							singleGraphics.DrawLine(pen, p1, p2);

							p1.X = _padding;
							p2.X = _elementWidth;
							pen.Color = ColorBorder;
							brush = new SolidBrush(ColorHigh);
							p1.Y = _height - ScaledClose[index];
							p2.Y = ScaledClose[index] - ScaledOpen[index];
							if (p2.Y == 0) p2.Y = 1;
							singleGraphics.DrawRectangle(pen, p1.X, p1.Y, p2.X, p2.Y);
							singleGraphics.FillRectangle(brush, p1.X + 1, p1.Y + 1, p2.X - 1, p2.Y - 1);
						}
						else
						{
							p1.X = p2.X = _padding + _elementWidth / 2;
							p1.Y = _height - ScaledHigh[index];
							p2.Y = _height - ScaledOpen[index];
							singleGraphics.DrawLine(pen, p1, p2);
							p1.Y = _height - ScaledClose[index];
							p2.Y = _height - ScaledLow[index];
							singleGraphics.DrawLine(pen, p1, p2);

							p1.X = _padding;
							p2.X = _elementWidth;
							pen.Color = ColorBorder;
							brush = new SolidBrush(ColorLow);
							p1.Y = _height - ScaledOpen[index];
							p2.Y = ScaledOpen[index] - ScaledClose[index];
							if (p2.Y == 0) p2.Y = 1;
							singleGraphics.DrawRectangle(pen, p1.X, p1.Y, p2.X, p2.Y);
							singleGraphics.FillRectangle(brush, p1.X + 1, p1.Y + 1, p2.X - 1, p2.Y - 1);
						}
						if (ReverseOrder != 0)
						{
							graphics.DrawImage(img, new Point(_width * (Count - index - 1), 0));
						}
						else
						{
							graphics.DrawImage(img, new Point(_width * index, 0));
						}
						singleGraphics.Dispose();
						brush.Dispose();
					}
				}
				image.Save(ImageAddress, System.Drawing.Imaging.ImageFormat.Png);
				graphics.Dispose();
			}
			pen.Dispose();
		}

		//// just get rest
		//internal void GetValue()
		//{
		//	return;
		//}

		// updating the measure - download and parse new values, redraw an imge with graph
		internal double Update()
		{
			//API.Log(API.LogType.Debug, "StockQuotes.dll: Update");
			Url = BuildUrl(Source);

			if (Url != string.Empty)
			{
				WebParserString = GetQuotes(Url);
			}
			else if (!Outdated)
			{
				//MarkOutdated();
			}

			if (WebParserString != string.Empty)
			{
				if (ParseData(WebParserString, Source) == true)
				{
					Scale();
					DrawGraph();
				}
			}
			else if (!Outdated)
			{
				//MarkOutdated();
			}

			return 0.0;
		}

		// return an address of the image we draw
		internal string GetString()
		{
			Update();
			return ImageAddress;
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

