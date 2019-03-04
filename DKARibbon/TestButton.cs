using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Microsoft.Office.Interop.Excel;
using XL = Microsoft.Office.Interop.Excel.Application;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;

using DKARibbon;
using System.Net;
using System.Text.RegularExpressions;
using System.Globalization;

namespace DKAExcelStuff
{
    public class TestButton
    {
        public static void TestM(KAXLApp kaxlApp)
        {
            CurrencyConversion(1, "CAD", "USD");
            
            
            // Key to api.forex/users/dashboard: ab12608f-3a25-457e-9689-a72db5290948









            //CurrencyConversion(1, "CAD", "USD");

            //WebClient web = new WebClient();
            //string url = string.Format("https://www.google.com/finance/converter?from=CAD&to=USD&a=1");
            //string response = web.DownloadString(url);

            //string url = string.Format("http://finance.yahoo.com/d/quotes.csv?s=CADUSD=X&f=11");
            //string response = new WebClient().DownloadString(url);

            // this attempt returns hundreds of lines of stuff, and does not find exrate
            //string url = string.Format("http://www.google.co.in/ig/calculator?h1=en&q=CADUSD%3D%3F1");
            //string blah = new WebClient().DownloadString(url);
            //decimal x = decimal.Parse(blah, System.Globalization.CultureInfo.InvariantCulture); // this line errors out, "not correct format"

            //string url = string.Format("https://www.google.com/finance/converter?a={0}&from={1}&to={1}", 
            //    Convert.ToString(1), "CAD", "USD");
            //WebRequest request = WebRequest.Create(url);            
            //request.UseDefaultCredentials = true;
            //StreamReader sr = new StreamReader(request.GetResponse().GetResponseStream(), System.Text.Encoding.ASCII); Error 403 - forbidden
            //string result = Regex.Matches(sr.ReadToEnd(), @"<span class="\" bld="">([^<]+)</span>")[0].Groups[1].Value;

            //string result = sr.ReadToEnd();

            // the below doesn't work....doesn't scrub exrate from massive string returned
            //NumberFormatInfo nfi = NumberFormatInfo.CurrentInfo;
            //string pattern;
            //pattern = @"^\s*[";
            //// Get the positive and negative sign symbols.
            //pattern += Regex.Escape(nfi.PositiveSign + nfi.NegativeSign) + @"]?\s?";
            //// Get the currency symbol.
            //pattern += Regex.Escape(nfi.CurrencySymbol) + @"?\s?";
            //// Add integral digits to the pattern.
            //pattern += @"(\d*";
            //// Add the decimal separator.
            //pattern += Regex.Escape(nfi.CurrencyDecimalSeparator) + "?";
            //// Add the fractional digits.
            //pattern += @"\d{";
            //// Determine the number of fractional digits in currency values.
            //pattern += nfi.CurrencyDecimalDigits.ToString() + "}?){1}$";

            //Regex rgx = new Regex(pattern);

            //double exRate;

            //exRate = rgx.IsMatch(result) ? Convert.ToDouble(result) : Convert.ToDouble(1);
        }

        public static string CurrencyConversion(decimal amount, string fromCurrency, string toCurrency)
        {    
            //string urlPattern = "http://finance.yahoo.com/d/quotes.csv?s={0}{1}=X&f=l1"; // System.Net.WebException: 'The remote name could not be resolved: 'download.finance.yahoo.com''
            string urlPattern = "http://www.google.co.in/ig/calculator?h1=en&q={0}{1}%3D%3F1"; 

            string url = string.Format(urlPattern, fromCurrency, toCurrency);
            
            using (var wc = new WebClient())
            {
                var response = wc.DownloadString(url);
                decimal exchangeRate = decimal.Parse(response, CultureInfo.InvariantCulture); // with google - data not in correct format

                return (amount * exchangeRate).ToString("N3");
            }
        }
    }  
}