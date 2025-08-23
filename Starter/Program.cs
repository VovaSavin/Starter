// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using System.IO.Compression; // Для розпакування архівів
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics; // Для запуску провідника
using Excel = Microsoft.Office.Interop.Excel;
//using VBIDE = Microsoft.Vbe.Interop;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        UnpackageArchive();
    }

    protected static void UnpackageArchive()
        // Розпаковує архів з основним файлом та допоміжними файлами
    {
        CreateDirectories();
        string myPath = AppDomain.CurrentDomain.BaseDirectory;
        string zipName = "Обробка витрат ВП_08_1.20252.zip";
        DirectoryInfo parentDirectory = Directory.GetParent(Directory.GetParent(myPath).FullName);
        if (parentDirectory != null)
        {
            string toZip = Path.Combine(parentDirectory.FullName, zipName);
            if(File.Exists(toZip))
            {
                Console.WriteLine("Архів знайдено.");
                string pathDestination = Path.Combine(SelectDrive(), "Artilery", "WorkFile");
                ZipFile.ExtractToDirectory(toZip, pathDestination);
                Console.WriteLine(
                    $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                );
            } else
            {
                MessageBox.Show(
                    $"Архіву не знайдено за цією адресою: {toZip}.\nВидобуття не відбулося.\nВиберіть файл.", 
                    toZip
                );

                // Відкриваємо провідник для вибору файлу
                string openedFile = OpenExplorerAndGetFile(
                    "Архіви zip (*.zip)|*.zip|All files (*.*)|*.*"
                );
                if (openedFile != "")
                {
                    string pathDestination = Path.Combine(SelectDrive(), "Artilery", "WorkFile");
                    ZipFile.ExtractToDirectory(openedFile, pathDestination);
                    Console.WriteLine(
                        $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                    );
                } else
                {
                    Console.WriteLine("Вибір файлу не відбувся. Завершення роботи програми.");
                }
            }
        }
    }

    protected static void CreateDirectories()
    // Створює директорії для звітів, бекапів та самого файлу
    {
        string drive = SelectDrive(true);
        string[] dirs = [
            Path.Combine(drive, "Artilery", "Reports"),
            Path.Combine(drive, "Artilery", "Backups"),
            Path.Combine(drive, "Artilery", "WorkFile"),
        ];
        foreach(string dir in dirs)
        {
            if(!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
                Console.WriteLine("Створено шлях: {0}.", dir);
            }
        }
    }

    protected static string SelectDrive(bool info=false)
        // Повертає локальний диск для встановлення файлу
    {
        if(info)
        {
            Console.WriteLine("Визначення томів в системі.");
        }
        DriveInfo[] allDrive = DriveInfo.GetDrives();
        if(allDrive.Length > 1)
        {
            if(info)
            {
                Console.WriteLine("Дані будуть встановлені на диск {0}.", allDrive[1].Name);
            }
            return allDrive[1].Name;
        } else
        {
            if (info)
            {
                Console.WriteLine("Дані будуть встановлені на диск {0}.", allDrive[0].Name);
            }
            return allDrive[0].Name;
        }
    }

    private static string OpenExplorerAndGetFile(string typesFile)
    {
        Console.WriteLine("Відкриття провідника для вибору файлу.");
        string selectedFilePath = "";
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
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Отримання шляху до вибраного файлу
                selectedFilePath = openFileDialog.FileName;
                Console.WriteLine($"Ви обрали файл: {selectedFilePath}", selectedFilePath);
            }
            else
            {
                MessageBox.Show("Ви не обрали файл. Завершення роботи програми.");
            }

        }
        return selectedFilePath;
    }

    public static void IsVba()
    {
        try
        {
            Type vbaType = Type.GetTypeFromProgID("VBIDE.Application");
            if (vbaType != null)
            {
                object vbaApp = Activator.CreateInstance(vbaType);
                Console.WriteLine("VBA доступний.");
            }
            else
            {
                Console.WriteLine("VBA не встановлено.");
            }
        }
        catch
        {
            Console.WriteLine("VBA не встановлено або недоступно.");
        }
    }
}