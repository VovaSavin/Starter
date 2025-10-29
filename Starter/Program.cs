// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using System.IO.Compression; // Для розпакування архівів
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics; // Для запуску провідника
using System.Management.Automation; // Для роботи з PowerShell
using System.Reflection; // Для роботи з PowerShell
using Microsoft.Win32;
using Json.More;
using Microsoft.VisualBasic;
using System.Collections.ObjectModel; // Для перевірки реєстру на наявність Excel

namespace Starter
{
    public static class Attributes
    {
        // Клас налаштувань та параметрів
        public static string AppDirAll { get; } = "Artilery_3027";
        public static string WorkFileDir { get; } = "WorkFile";
        public static string ReportDir { get; } = "Backups";
        public static string BackUpDir { get; } = "Reports";
        public static string NameDirStarter { get; } = "StarterShootCounter";

    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (IsExcelInstalled())
            {
                // Зчитуємо налаштування з конфігурації
                string nameAppDir = Attributes.AppDirAll;
                string workFile = Attributes.WorkFileDir;
                string reportDir = Attributes.ReportDir;
                string backUpDir = Attributes.BackUpDir;
                string nameDirStarter = Attributes.NameDirStarter;
                string workedDrive = SelectDrive();
                CreateDirectories(
                    workedDrive, 
                    nameAppDir, 
                    reportDir, 
                    backUpDir, 
                    workFile, 
                    nameDirStarter
                    );

                UnpackageArchive(
                    Path.Combine(
                        workedDrive, 
                        nameAppDir, 
                        reportDir
                    ),
                    Path.Combine(
                        workedDrive, 
                        nameAppDir, 
                        backUpDir
                    ),
                    Path.Combine(
                        workedDrive, 
                        nameAppDir, 
                        workFile, 
                        nameDirStarter
                    ),
                    Path.Combine(
                        workedDrive, 
                        nameAppDir, 
                        workFile
                     )
                );
            }
            else
            {
                MessageBox.Show(
                    "Microsoft Excel не встановлено на цьому комп'ютері.\n" +
                    "Будь ласка, встановіть Microsoft Excel та повторіть спробу.",
                    "Excel не знайдено",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

        }

        protected static void UnpackageArchive(
            string pathReport,
            string pathBackups,
            string thisDirectory,
            string workFile
        )
        // Розпаковує архів з основним файлом та допоміжними файлами
        {
            string myPath = AppDomain.CurrentDomain.BaseDirectory;
            string zipName = "packed.zip"; 
            // Отримуємо батьківську директорію
            DirectoryInfo parentDirectory = Directory.GetParent(
                Directory.GetParent(myPath).FullName
            );
            // Якщо батьківська директорія існує, продовжуємо
            if (parentDirectory != null)
            {
                string toZip = Path.Combine(parentDirectory.FullName, zipName); // Шлях до архіву
                string pathDestination = workFile;
                // Перевіряємо, чи існує архів
                if (File.Exists(toZip))
                {
                    Console.WriteLine("Архів знайдено.");
                    // Перевіряємо, чи є в директорії робочий файл
                    // Якщо директорія порожня, розпаковуємо архів
                    if (GetFirstFile(pathDestination) == "")
                    {
                        ZipFile.ExtractToDirectory(toZip, pathDestination);
                        Console.WriteLine(
                            $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                        );
                    }

                }
                else
                {
                    if (GetFirstFile(pathDestination) == "")
                    {
                        MessageBox.Show(
                            $"Архіву не знайдено за цією адресою: {toZip}.\nВидобуття не відбулося.\nВиберіть файл.",
                            toZip
                        );

                        // Вибираємо та перевіряємо тип вибраного файлу
                        string openedFile = CheckTypeFile("zip");

                        // Якщо файл не пустий
                        if (openedFile != "")
                        {
                            // Дивимось, що є в архіві
                            List<ZipArchiveEntry> inArchive = GetDataArchive(openedFile);
                            if (inArchive.Count == 0 || inArchive.Count > 1)
                            {
                                MessageBox.Show(
                                    "Виберіть архів з одним файлом типу *.xlsm.",
                                    "Даний архів містить більше ніж один файл або пустий.",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                                return;
                            }

                            // Якщо файл в архіві є xlsm
                            if (OnlyCheckTypeFile("xlsm", inArchive.First().ToString()))
                            {
                                ZipFile.ExtractToDirectory(openedFile, pathDestination);
                                Console.WriteLine(
                                    $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                                );
                            } else
                            {
                                MessageBox.Show(
                                    "Файл не є *.xlsm.",
                                    "Даний архів містить файл відмінний від *.xlsm.",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show(
                                    "Не вибраний файл.",
                                    "Ви не вибрали  файл.",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine(
                            "Робочий файл вже існує. Видобуття не відбулося.\nПродовження розпаковки."
                        );
                    }
                }
                // Викликаємо PowerShell скрипт для зміни макросів
                ExecutePowerShellScript(GetFirstFile(pathDestination), pathReport, pathBackups);

                // Копіюємо себе в директорію з робочим файлом
                SelfCopy(parentDirectory.FullName, thisDirectory, true);

                // Виводимо інформацію про успішне завершення програми
                MessageBox.Show(
                    "Підготовка файлу завершена.\n" +
                    "Тепер ви можете успішно з ним працювати.",
                    "Готово",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
        }

        protected static void CreateDirectories(
            string drive, 
            string nameAppDir,
            string report,
            string backup,
            string workFile,
            string starterShootCounter
            )
        // Створює директорії для звітів, бекапів та самого файлу
        {
            Console.WriteLine(
                "Перевірка наявності директорій."
            );
            string[] dirs = [
                Path.Combine(drive, nameAppDir, report),
            Path.Combine(drive, nameAppDir, backup),
            Path.Combine(drive, nameAppDir, workFile),
            Path.Combine(
            drive, nameAppDir, workFile, starterShootCounter
        )
            ];

            // Запускаємо цикл по кожній директорії та перевіряємо її наявність
            // Створюємо директорії, якщо їх немає
            foreach (string dir in dirs)
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                    Console.WriteLine("Створено шлях: {0}.", dir);
                }
                else
                {
                    Console.WriteLine("Шлях вже існує: {0}.", dir);
                }
            }
        }

        protected static string SelectDrive()
        // Повертає локальний диск для встановлення файлу
        {

            DriveInfo[] allDrive = DriveInfo.GetDrives();
            if (allDrive.Length > 1)
            {
                return allDrive[1].Name;
            }
            else
            {
                return allDrive[0].Name;
            }
        }

        private static string OpenExplorerAndGetFile(string typesFile)
        // Відкриває провідник для вибору файлу. Повертає шлях до вибраного файлу
        {
            Console.WriteLine(
                "Відкриття провідника для вибору файлу."
            );
            string selectedFilePath = "";

            // Ініціалізація діалогу вибору файлу
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                Console.WriteLine("Ініціалізація діалогу вибору файлу.");
                // Налаштування діалогу
                openFileDialog.InitialDirectory = Environment.GetFolderPath(
                    Environment.SpecialFolder.Desktop
                );
                openFileDialog.Title = "Вкажіть шлях до архіву з файлами обробки витрат ВП";

                // Встановка фільтрів для типів файлів
                openFileDialog.Filter = typesFile;

                // Показ діалогу та отримання результату. Якщо файл вибрано, зберігаємо шлях до нього
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Отримання шляху до вибраного файлу
                    selectedFilePath = openFileDialog.FileName;
                    Console.WriteLine(
                        $"Ви обрали файл: {selectedFilePath}", selectedFilePath
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Ви не обрали файл. Завершення роботи програми."
                    );
                }

            }
            return selectedFilePath;
        }

        private static void CallPowerShellScript(string scriptPath, string fileExcelPath)
        // Викликає PowerShell скрипт. Для роботи з COM об'єктами
        // Не використовується
        {
            // Перевірка, чи існує файл
            if (!File.Exists(fileExcelPath))
            {
                Console.WriteLine("Робочого файлу для обробки не знайдено. (CallPowerShellScript)");
                return;
            }
            if (!File.Exists(scriptPath))
            {
                Console.WriteLine("Скрипт не знайдено.");
                return;
            }
            else
            {
                // Створення нового процесу
                Process process = new Process();

                // Запуск скрипта
                process.StartInfo.FileName = "powershell.exe";

                // -ExecutionPolicy ByPass обходить політику виконання скриптів
                // -File вказує, що буде запускатися файл
                process.StartInfo.Arguments = $"-ExecutionPolicy ByPass -File \"{scriptPath}\" -PathFileExcel \"{fileExcelPath}\"";

                // Не показувати вікно консолі
                process.StartInfo.CreateNoWindow = true;

                // Не використовувати оболонку ОС
                process.StartInfo.UseShellExecute = false;

                try
                {
                    Console.WriteLine("Запуск скрипту зміни макросів.");

                    // Запуск процесу
                    process.Start();

                    // Очікування завершення процесу
                    process.WaitForExit();

                    Console.WriteLine("Скрипт завершив роботу.");
                    Console.WriteLine($"Код виходу: {process.ExitCode}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Сталася помилка: {ex.Message}");
                }

            }
        }

        private static string GetFirstFile(string pathToFile)
        // Повертає перший файл з директорії. Якщо файлів немає, повертає порожній рядок
        {
            // Повертає файл з директорії
            DirectoryInfo dirInfo = new DirectoryInfo(pathToFile);
            FileInfo[] files = dirInfo.GetFiles();
            if (files.Length > 0)
            {
                return files[0].FullName;
            }
            else
            {
                Console.WriteLine(
                    "Робочого файлу у директорії не знайдено. (GetFirstFile)"
                );
                return "";
            }
        }

        public static void ExecutePowerShellScript(
            string workFilePath, string pathReport, string pathBackups
            )
        // Виконує PowerShell скрипт для зміни макросів
        {
            // Завантажуємо скрипт з вбудованого ресурсу
            var assembly = Assembly.GetExecutingAssembly();
            var resourseName = "Starter.edit.ps1";
            string scriptContent = "";

            // Читаємо вміст вбудованого ресурсу за допомогою потоку
            using (Stream stream = assembly.GetManifestResourceStream(resourseName))
            {
                // Перевіряємо, чи вдалося знайти ресурс. Якщо потік null, ресурс не знайдено
                if (stream == null)
                {
                    Console.WriteLine("Не вдалося знайти вбудований ресурс.");
                    return;
                }
                using (StreamReader reader = new StreamReader(stream))
                {
                    scriptContent = reader.ReadToEnd();
                }
            }

            // Виконуємо скрипт
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddScript(scriptContent); // Додаємо скрипт
                ps.AddParameter("PathFileExcel", workFilePath); // Додаємо параметри 
                ps.AddParameter("PathToReport", pathReport); // Додаємо параметри
                ps.AddParameter("PathToBackup", pathBackups); // Додаємо параметри
                ps.AddParameter("ExecutionPolicy", "ByPass"); // Обхід політики виконання
                var result = ps.Invoke(); // Виконуємо скрипт

                // Якщо є помилки, виводимо їх
                if (ps.Streams.Error.Any())
                {
                    Console.WriteLine($"Сталася помилка під час виконання скрипта.");
                    foreach (var error in ps.Streams.Error)
                    // Виводимо помилки
                    {
                        Console.WriteLine(error.ToString());
                        MessageBox.Show(
                            error.ToString()+ "\n" + "Помилка при спробі зміни макросів.",
                            "Помилка виконання скрипта",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
            }
            Console.WriteLine("Макроси успішно змінено.");
        }

        public static void SelfCopy(string parentDirectory, string destinationDir, bool recursive)
        // Копіює себе в директорію з робочим файлом
        {
            var dir = new DirectoryInfo(parentDirectory); // Отримуємо інформацію про батьківську директорію

            // Перевіряємо, чи існує батьківська директорія
            if (!dir.Exists)
            {
                // Якщо батьківська директорія не існує, викидаємо виключення
                throw new DirectoryNotFoundException(
                    $"Батьківська директорія не існує або не знайдена: {dir.FullName}"
                );
            }

            // Якщо директорія призначення не існує, створюємо її
            if (!Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }

            foreach (FileInfo file in dir.GetFiles())
            // Копіюємо кожен файл з батьківської директорії в директорію призначення
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath, true);
            }

            if (recursive) // Якщо потрібно копіювати піддиректорії
            {
                // Рекурсивно копіюємо всі піддиректорії
                foreach (DirectoryInfo subDir in dir.GetDirectories())
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    SelfCopy(subDir.FullName, newDestinationDir, true); // Рекурсивний виклик
                }
            }
            Console.WriteLine($"Файли було скопійовано в {destinationDir}");
        }

        protected static bool IsExcelInstalled()
        // Перевіряє, чи встановлений Excel на комп'ютері
        {
            RegistryKey keyExcel = Registry.ClassesRoot.OpenSubKey(@"Excel.Application");
            if (keyExcel != null)
            {
                keyExcel.Close();
                return true; // Excel встановлений
            }
            else
            {
                return false; // Excel не встановлений
            }
        }
        
        protected static string CheckTypeFile(string typeFile)
        // Перевіряє розширення файлу та повертає , якщо користувач вибрав потрібний тип
        // Запускає цикл поки не буде вираний потрібний формат файлу
        {
            string archive = OpenExplorerAndGetFile(
                "Архіви zip (*.zip)|*.zip|All files (*.*)|*.*"
                ); // Відкриваємо архів
            string extention = Path.GetExtension(archive); // Отримуємо розширення

            // Перевіряємо чи не пустий рядок
            if (archive != "")
            {
                // Запускаємо цикл перевірки
                while (!extention.Equals("." + typeFile, StringComparison.OrdinalIgnoreCase))
                {

                    // Відкриваємо архів знову
                    archive = OpenExplorerAndGetFile(
                        "Архіви zip (*.zip)|*.zip|All files (*.*)|*.*"
                    );
                    extention = Path.GetExtension(archive); // Отримуємо розширення
                }
            } else
            {
                archive = "";
            }
                return archive;
        }

        protected static bool OnlyCheckTypeFile(string typeFile, string filePath)
        // Перевіряє тип файлу та повертає потрібний
        {
            string extention = Path.GetExtension(filePath);
            if (extention.Equals("." + typeFile, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            } else
            {
                return false;
            }
        }

        protected static List<ZipArchiveEntry> GetDataArchive(string pathArchiveZip)
        // Повертає вміст архіву
        {
            using(ZipArchive arch = ZipFile.OpenRead(pathArchiveZip)) // Створення об'єкта архіву з його переглядом
            {
                if (arch.Entries.Count > 0)
                {
                    return arch.Entries.ToList();
                }
                return [];
            }
        }
    }
}

