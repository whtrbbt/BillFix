using System;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace BillFix
{
    class Program
    {
        static void Main(string[] args)
        {
            string dirpathIN = @ConfigurationManager.AppSettings.Get("dirpathIN");
            string dirpathOUT = @ConfigurationManager.AppSettings.Get("dirpathOUT");
                       
            FixDir(dirpathIN, dirpathOUT);
            Console.WriteLine("Готово!");
            Console.ReadKey();
        }

        static void FixDir(string inDir, string outDir)
        {
            var dirIN = new DirectoryInfo(@inDir); //папка с входящими файлами 
            var dirOUT = new DirectoryInfo(@outDir); //папка с исходящими файлами  
            string dirName = "";

            foreach (DirectoryInfo dir in dirIN.GetDirectories()) // ищем все подкаталоги в каталоге dirIN
            {
                dirName = Path.GetFileName(dir.FullName); // получаем имя текущего подкаталога
                Console.WriteLine(dirName);
                dirName = dirOUT + @"\" + dirName;
                if (!Directory.Exists(dirName))
                    Directory.CreateDirectory(dirName);
                FixDir(dir.FullName, dirName);

            }
            FixFiles(dirIN.FullName, dirOUT.FullName);
        }

        static void FixFiles(string inDir, string outDir)
        {
            var dirIN = new DirectoryInfo(@inDir); // папка с входящими файлами 
            var dirOUT = new DirectoryInfo(@outDir); // папка с исходящими файлами 
            DirectoryInfo tmpDir; // временная папка для создания архива
            DirectoryInfo tmpUnZipDir; //временная папка для распаковки архива
            string fileName = "";

            foreach (FileInfo file in dirIN.GetFiles())
            {
                fileName = Path.GetFileName(file.FullName);
                Console.WriteLine(fileName);
                if (Path.GetExtension(fileName) == ".zip")
                {
                    tmpDir = CreateTempDir();
                    tmpUnZipDir = UnzipFileToTempDir(file.FullName);
                    FixDir(tmpUnZipDir.FullName, tmpDir.FullName);
                    ZipFile.CreateFromDirectory(tmpDir.FullName, @outDir + @"\" + fileName);
                    tmpDir.Delete(true);
                    tmpUnZipDir.Delete(true);
                }
                else //Path.GetExtension(fileName) <> ".zip"
                {
                    fileName = RemoveInvalidFilePathCharacters(fileName, "");
                    FixBill(@file.FullName, @outDir + @"\" + fileName);
                }
            }
        }

        static void FixBill(string inFileName, string outFileName = "")
        {
            // Объявляем приложение
            Excel.Application exc = new Microsoft.Office.Interop.Excel.Application();

            Excel.XlReferenceStyle RefStyle = exc.ReferenceStyle;

            Excel.Workbook wb = null;



            try
            {
                wb = exc.Workbooks.Add(inFileName); // !!! 
            }
            catch (System.Exception ex)
            {
                throw new Exception("Не удалось загрузить файл! " + inFileName + "\n" + ex.Message);
            }
            //Console.WriteLine("Файл найден, начинаю работу. Это может занять несколько минут.");
            //Excel.Sheets excelsheets;

            // Выбираем 1 лист
            Excel.Worksheet wsh = wb.Worksheets.get_Item(1) as Excel.Worksheet;

            Excel.Range excelcells;

            excelcells = wsh.get_Range("C19", "C19");
            excelcells.Value2 = "";

            if (outFileName != "")
                wb.SaveAs(outFileName);
            else
                wb.SaveAs(inFileName);
            exc.Quit();

        }

        public static string RemoveInvalidFilePathCharacters(string filename, string replaceChar)
        // Удаляет запрещенные символы в именах файлов      
        {
                return Regex.Replace(filename, "[\\[\\]]+", replaceChar, RegexOptions.Compiled);
        }

        public static DirectoryInfo UnzipFileToTempDir (string fileName)
        // Распаковывает архив во временную папку и возвращает ссылку на неё.      
        {
            DirectoryInfo tempDir;
            tempDir = CreateTempDir();
            ZipFile.ExtractToDirectory(fileName, @tempDir.FullName);
            return tempDir;
        }

        public static DirectoryInfo CreateTempDir ()
        // Создает временную папку и возращает ссылку на нее
        {
            DirectoryInfo tempDir;
            tempDir = Directory.CreateDirectory(Path.GetTempPath() + Path.GetRandomFileName());
            return tempDir;
        }
    }
}
