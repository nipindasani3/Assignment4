using Assignment4.Models.Utilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Assignment4.Models
{
    public class Methods
    {
        //Download Info Data From FTP Server and sace to Local File
        public static void downloadFtpDataFiles()
        {
            Student student = new Student();
            String path = Constants.Locations.StudentDataFolder;
            Methods.directoryExit(path);
            List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            Console.WriteLine("========= Downloading Info File For Each Student =========\n");
            foreach (var Directory in directories)
            {
                student = new Student() { AbsoluteUrl = Constants.FTP.BaseUrl };
                student.FromDirectory(Directory);
                string infoPath = student.FullPathUrl + "/" + Constants.Locations.InfoFile;
                bool fileExist = FTP.FileExists(infoPath);
                //Console.WriteLine(student);
                if (fileExist == true)
                {
                    string csvPath = $@"{path}\{Directory}.csv";
                    FTP.DownloadFile(infoPath, csvPath);
                    byte[] bytes = FTP.DownloadFileBytes(infoPath);
                    string csvData = Encoding.Default.GetString(bytes);
                    string[] csvlines = csvData.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
                    if (csvlines.Length != 2)
                    {
                        Console.WriteLine("Error in CSV Format");
                    }
                    else
                    {
                        student.FromCSV(csvlines[1]);
                    }
                    Console.WriteLine("Found info File");
                }
                else
                {
                    Console.WriteLine("Could not find info file:");
                }
                Console.WriteLine("\t" + infoPath);
            }
        }

        //Download Image Data from FTP Server and save to Local File
        public static void downloadFtpImageFiles()
        {
            Student student = new Student();
            String path = Constants.Locations.StudentImageFolder;
            Methods.directoryExit(path);
            List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            Console.WriteLine("========= Downloading Image File For Each Student =========\n");
            foreach (var Directory in directories)
            {
                string imageFilePath = student.FullPathUrl + "/" + Constants.Locations.ImageFile;
                string imagePath = $@"{path}\{Directory}.jpeg";
                FTP.DownloadFile(imageFilePath, imagePath);
                bool fileExist1 = FTP.FileExists(imageFilePath);
                if (fileExist1 == true)
                {
                    Console.WriteLine("Found image File");
                }
                else
                {
                    Console.WriteLine("Could not find the file:");
                }
                Console.WriteLine("\t" + imageFilePath);
            }
        }

        //Read Data From info.Csv File
        //Retrun List of String of Student Data
        public static List<String> readCSV(String FilePath)
        {
            String fileContents;
            using (StreamReader stream = new StreamReader(FilePath))
            {
                fileContents = stream.ReadToEnd();
            }
            List<String> entries = fileContents.Split("\r\n", StringSplitOptions.RemoveEmptyEntries).ToList();
            return entries;
        }

        //Read Student Data from StudentCSV File
        //Return List of Student Object
        public static List<Student> _readCSV(String FilePath)
        {
            List<Student> studentlist = new List<Student>();
            List<String> entries = Methods.readCSV(FilePath);
            foreach (var x in entries)
            {
                Student tempstudent = new Student();
                tempstudent.FromCSV(x);
                studentlist.Add(tempstudent);
            }
            return studentlist;
        }

        //Read Data From Json File
        //Return List of Student Object
        public static List<Student> readJSON(String FilePath)
        {
            List<Student.Student_Data> studentData;
            using (StreamReader readFile = new StreamReader(FilePath))
            {
                string json = readFile.ReadToEnd();
                studentData = JsonConvert.DeserializeObject<List<Student.Student_Data>>(json);
            }
            List<Student> studentlist = setStudentObj(studentData);
            return studentlist;
        }

        //Read Data From Xml File
        //Return List of Student Object
        public static List<Student> readXML(String FilePath)
        {
            Student.StudentMapping studentMapping;
            var serializer = new XmlSerializer(typeof(Student.StudentMapping));
            string data = System.IO.File.ReadAllText(FilePath);
            using (var stream = new StringReader(data))
            using (var reader = XmlReader.Create(stream))
            {
                studentMapping = (Student.StudentMapping)serializer.Deserialize(reader);
            }
            List<Student> studentlist = setStudentObj(studentMapping.Student);
            return studentlist;
        }

        //Write CSV File
        public static void writeCSV()
        {
            List<Student> studentlist = new List<Student>();
            String error;
            (studentlist, error) = Methods.SetInfoData();
            try
            {
                using (StreamWriter fs = new StreamWriter(Constants.Locations.StudentCSVPath))
                {
                    foreach (var stu in studentlist)
                    {
                        fs.WriteLine(stu.ToCSV());
                    }
                }
                Console.WriteLine("\n============CSV File Created============\n");
                Console.WriteLine("========Error Found in Below Files========\n" + error);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        //Write JSON File
        public static void writeJSON()
        {
            List<Student> studentlist = new List<Student>();
            String error;
            List<Student.Student_Data> student_Data = new List<Student.Student_Data>();
            (studentlist, error) = SetInfoData();
            try
            {

                foreach (var c in studentlist)
                {
                    Student.Student_Data obj = new Student.Student_Data
                    {
                        _StudentID = $"{c.StudentId}",
                        _FirstName = $"{c.FirstName}",
                        _LastName = $"{c.LastName}",
                        _DateOfBirth = $"{c.DateOfBirthDT.ToShortDateString()}",
                        _ImageData = $"{c.ImageData}"
                    };
                    student_Data.Add(obj);
                }
                using (StreamWriter file = File.CreateText(Constants.Locations.StudentJSONPath))
                using (JsonTextWriter writer = new JsonTextWriter(file))
                {
                    Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
                    serializer.Serialize(file, student_Data);
                }
                Console.WriteLine("\n============JSON File Created============\n");
                Console.WriteLine("========Error Found in Below Files========\n" + error);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        //Write Xml File
        public static void writeXML()
        {
            List<Student> studentlist = new List<Student>();
            String error;
            (studentlist, error) = Methods.SetInfoData();
            try
            {
                XmlTextWriter writer = new XmlTextWriter(Constants.Locations.StudentXMLPath, System.Text.Encoding.UTF8);
                writer.WriteStartDocument(true);
                writer.Formatting = System.Xml.Formatting.Indented;
                writer.Indentation = 2;
                writer.WriteStartElement("Student_Info");
                foreach (var stu in studentlist)
                {
                    stu.ToXML(writer);
                }
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
                Console.WriteLine("\n============XML File Created============\n");
                Console.WriteLine("========Error Found in Below Files========\n" + error);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        //Set Values for Student Object from Student.Student_Data
        //Return List of Student Object
        public static List<Student> setStudentObj(List<Student.Student_Data> studentData)
        {
            List<Student> studentlist = new List<Student>();
            foreach (var stu in studentData)
            {
                Student student = new Student();
                String dataFromObj = stu.ToList();
                student.FromCSV(dataFromObj);
                studentlist.Add(student);
            }
            return studentlist;
        }

        //Check for Directory Existance
        public static void directoryExit(String path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }


        //Get All Student Csv Files From Student Folder and Load them into Student Object
        public static (List<Student>, string) SetInfoData()
        {
            string error = null;
            string path = Constants.Locations.StudentDataFolder;
            directoryExit(path);
            string[] getFileList = Directory.GetFiles(path);
            List<Student> studentlist = new List<Student>();
            foreach (string files in getFileList)
            {
                Student tempstudent = new Student();
                List<String> entries = new List<string>();
                entries = Methods.readCSV(files);
                try
                {
                    tempstudent.FromCSV(entries[1]);
                    studentlist.Add(tempstudent);
                }
                catch
                {
                    error = error + Path.GetFileName(files) + "\n";
                }
            }
            return (studentlist, error);
        }

        //Get Formatted file Like: CSV, JSON, XML, etc. and load data to Student Object
        public static List<Student> SetMyData(String FilePath)
        {
            var result = FilePath.Split('.').Last();
            List<Student> studentlist = new List<Student>();
            try
            {
                if (result == "csv")
                {
                    studentlist = _readCSV(FilePath);
                }
                else if (result == "json")
                {
                    studentlist = readJSON(FilePath);
                }
                else if (result == "xml")
                {
                    studentlist = readXML(FilePath);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            Console.WriteLine(studentlist.Count + " Student Found");
            return studentlist;
        }

 

        //Convert Image into base64 and add value to CSV file
        public static void SetImageData()
        {
            Image image = Image.FromFile(Constants.Locations.ImageFilePath);
            string base64Image = imaging.ImageToBase64(image, ImageFormat.Jpeg);
            var csv = File.ReadLines(Constants.Locations.InfoFilePath).Select((line, index) => index == 0 ? line + ",Image"
            : line + "," + base64Image).ToList();
            File.WriteAllLines(Constants.Locations.InfoFilePath, csv);
            Console.WriteLine("Image Data Entered Successfully");
        }

    }
}
