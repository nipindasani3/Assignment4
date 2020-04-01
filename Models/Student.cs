using Newtonsoft.Json.Linq;
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
    public class Student
    {
     
        public class StudentMapping
        {
            public StudentMapping()
            {
                Student = new List<Student_Data>();
            }
          
            public List<Student_Data> Student { get; set; }
        }

        public class Student_Data
        {
            [XmlElement("Student_ID")]
            public String _StudentID { get; set; }
            [XmlElement("First_Name")]
            public String _FirstName { get; set; }
            [XmlElement("Last_Name")]
            public String _LastName { get; set; }
            [XmlElement("Date_Of_Birth")]
            public String _DateOfBirth { get; set; }
            [XmlElement("Image_Data")]
            public String _ImageData { get; set; }

            public String ToList()
            {
                string result = $"{_StudentID}, {_FirstName }, {_LastName}, {_DateOfBirth}, {_ImageData} ";
                return result;
            }
        }

        public static string HeaderRow = $"{nameof(Student.StudentId)},{nameof(Student.FirstName)},{nameof(Student.LastName)},{nameof(Student.DateOfBirth)},{nameof(Student.ImageData)}";
        public String StudentId { get; set; }
        public String FirstName { get; set; }
        public String LastName { get; set; }

        private String _DateOfBirth;
        private Boolean _MyRecord { get; set; }
        //Get and set Age of student
        public int Age
        {
            get
            {
                return new DateTime(DateTime.Now.Subtract(DateOfBirthDT).Ticks).Year - 1; ;

            }
        }
        public String DateOfBirth
        {
            get { return _DateOfBirth; }
            set
            {
                _DateOfBirth = value;
                //Convert DateOfBirth to DateTime
                DateTime dtOut;
                DateTime.TryParse(_DateOfBirth, out dtOut);
                DateOfBirthDT = dtOut;
            }
        }
        public DateTime DateOfBirthDT { get; internal set; }
        public String ImageData { get; set; }
        public String AbsoluteUrl { get; set; }
        public String Directory { get; set; }

        public string FullPathUrl
        {
            get
            {
                return AbsoluteUrl + "/" + Directory;
            }
        }

        public List<string> Exceptions { get; set; } = new List<string>();

        public bool IsMe { get; set; }
        public void FromCSV(String csvdate)
        {
            string[] data = csvdate.Split(',', StringSplitOptions.None);
            try
            {
                StudentId = data[0];
                FirstName = data[1];
                LastName = data[2];
                DateOfBirth = data[3];
                ImageData = data[4];
            }
            catch (Exception e)
            {
                Exceptions.Add(e.Message);
            }
        }

        public void FromDirectory(string directory)
        {
            Directory = directory;
            if (String.IsNullOrEmpty(directory.Trim()))
            {
                return;
            }
            String[] data = directory.Trim().Split(' ', StringSplitOptions.None);
            StudentId = data[0];
            FirstName = data[1];
            LastName = data[2];
        }

        public string ToCSV()
        {
            string result = $"{StudentId}, {FirstName }, {LastName}, {DateOfBirthDT.ToShortDateString()}, {ImageData} ";
            return result;
        }

        //Create XML object of Student data
        public void ToXML(XmlTextWriter writer)
        {
            writer.WriteStartElement("Student");
            writer.WriteStartElement("Student_ID");
            writer.WriteString(StudentId);
            writer.WriteEndElement();
            writer.WriteStartElement("First_Name");
            writer.WriteString(FirstName);
            writer.WriteEndElement();
            writer.WriteStartElement("Last_Name");
            writer.WriteString(LastName);
            writer.WriteEndElement();
            writer.WriteStartElement("Date_Of_Birth");
            writer.WriteString(DateOfBirthDT.ToShortDateString());
            writer.WriteEndElement();
            writer.WriteStartElement("Image_Data");
            writer.WriteString(ImageData);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        public override string ToString()
        {
            string result = $"Student Id: {StudentId}, \nFirst Name: {FirstName }, \nLast Name: {LastName}, \nAge: {Age}";
            return result;
        }

    }
}


