using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Specialized;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadFileLog
{
    class Program
    {
        static string _path_folder = @"D:\Documents\Cencosud\Logs Web Privada\";
        static string _file_name = "autorizador_security-2020-10-31";
        static string _pattern_search = @"         000023";
        static string _first_pattern = @"_____";

        static void Main(string[] args)
        {
            string _path = string.Format("{0}{1}.{2}", _path_folder, _file_name, "log");
            List<KeyValuePair<string, string>> _resultData = new List<KeyValuePair<string, string>>();

            using (FileStream _file = new FileStream(_path, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader _sr = new StreamReader(_file))
                {
                    string line = String.Empty;

                    while (!_sr.EndOfStream)
                    {
                        line = _sr.ReadLine();

                        if (line.Contains(_pattern_search)) {
                            bool is_valid = false;
                            string _date = Get_Time(line);
                            string _dni = Get_DNI(line, out is_valid);

                            if (is_valid) _resultData.Add(new KeyValuePair<string, string>(_date, _dni));
                        }
                    }
                }
            }

            Export_Excel(_resultData);
        }

        static string Get_DNI(string line, out bool is_valid)
        {
            string _DNI = string.Empty;
            is_valid = Is_valid_HBK(line);
            if(is_valid) _DNI = Get_Format_DNI(line);
            return _DNI;
        }

        static bool Is_valid_HBK(string line)
        {
            string _pattern = "HBK";
            int _first_position = line.IndexOf(_pattern);
            string _sub_string = line.Substring(_first_position, 15);

            _sub_string = _sub_string.Substring(0, _sub_string.IndexOf(" "));
            string _new_string = _sub_string.Substring((_sub_string.Length - 2));

            return (_new_string.Equals("00"));
        }

        static string Get_Format_DNI(string line)
        {
            int _default_split = 30;
            int _first_point = line.IndexOf(_first_pattern);
            _first_point = (_first_point == -1) ? line.IndexOf("___") : _first_point;   //en algunos casos el patron es menor.
            string _sub_string = line.Substring(_first_point, _default_split);
            int _init_position_docuemnt = (_sub_string.IndexOf("D") > -1 )? _sub_string.IndexOf("D") : ((_sub_string.IndexOf("C") > -1) ? _sub_string.IndexOf("C") : 0);

            _sub_string = _sub_string.Substring(_init_position_docuemnt, (_sub_string.Length - _init_position_docuemnt));
            _sub_string = _sub_string.Substring(0, _sub_string.IndexOf("_"));
            return _sub_string.Trim();
        }

        static string Get_Time(string line)
        {
            int _last_position = (line.IndexOf(",") - 1);
            return line.Substring(1, _last_position);
        }

        static void Export_Excel(List<KeyValuePair<string, string>> data)
        {            
            string _path = string.Format("{0}{1}.{2}", _path_folder, _file_name, "xls");

            Excel.Application _excel = new Excel.Application();
            Excel.Workbook workbook = _excel.Workbooks.Add();
            Excel._Worksheet sheet = workbook.Sheets[1];
            Excel.Range range = sheet.UsedRange;

            //Graba la data la hoja excel.
            Set_DATA(data, sheet);

            workbook.SaveAs(_path);
            workbook.Close();
            _excel.Quit();
        }

        static void Set_DATA(List<KeyValuePair<string, string>> data, Excel._Worksheet sheet)
        {
            KeyValuePair<string, string> _header = new KeyValuePair<string, string>("Fecha", "Documento");
            int _currentrow = 1;

            //agrega los encabezados
            sheet.Cells[_currentrow, 1] = _header.Key;
            sheet.Cells[_currentrow, 2] = _header.Value;

            foreach (KeyValuePair<string, string> item in data)
            {
                _currentrow++;
                sheet.Cells[_currentrow, 1] = item.Key;
                sheet.Cells[_currentrow, 2] = item.Value;
            }
        }
    }
}
