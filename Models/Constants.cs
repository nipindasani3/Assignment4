using System;

namespace CSV.Models
{
    class Constants
    {

        public readonly Student Student = new Student { StudentId = "200447887", FirstName = "Nipin", LastName = "Dasani" };

        public class Locations
        {
            static Constants con = new Constants();
            //public readonly static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            public readonly static string ExePath = Environment.CurrentDirectory;
            public readonly static string ContentFolder = $"{ExePath}\\..\\..\\..\\Contents";
            public readonly static string DataFolder = $"{ContentFolder}\\Data";
            public readonly static string ImageFolder = $"{ContentFolder}\\Images";
            public readonly static string StudentFolder = $"{ContentFolder}\\StudentFolder";
            public readonly static string StudentDataFolder = $"{StudentFolder}\\Data";
            public readonly static string StudentImageFolder = $"{StudentFolder}\\Images";
            public const string InfoFile = "info.csv";
            public const string ImageFile = "myimage.jpg";
            public const string StudentCSVFile = "Students.csv";
            public const string StudentJSONFile = "Students.json";
            public const string StudentXMLFile = "Students.xml";
            public readonly static string InfoFilePath = $"{DataFolder}\\{InfoFile}";
            public readonly static string ImageFilePath = $"{ImageFolder}\\{ImageFile}";
            public readonly static string StudentCSVPath = $"{DataFolder}\\{StudentCSVFile}";
            public readonly static string StudentJSONPath = $"{DataFolder}\\{StudentJSONFile}";
            public readonly static string StudentXMLPath = $"{DataFolder}\\{StudentXMLFile}";
            //public readonly static string remoteUploadCsvFileDestination = "/" + con.Student.StudentId + " " + con.Student.FirstName
            //    + " " + con.Student.LastName + "/" + StudentCSVFile;
            //public readonly static string remoteUploadJsonFileDestination = "/" + con.Student.StudentId + " " + con.Student.FirstName
            //    + " " + con.Student.LastName + "/" + StudentJSONFile;
            //public readonly static string remoteUploadXmlFileDestination = "/" + con.Student.StudentId + " " + con.Student.FirstName
            //    + " " + con.Student.LastName + "/" + StudentXMLFile;
        }

        public class FTP
        {
            public const string Username = @"bdat100119f\bdat1001";
            public const string Password = "bdat1001";
            public const string BaseUrl = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914";
            public const int OperationPauseTime = 10000;
        }
    }
}
