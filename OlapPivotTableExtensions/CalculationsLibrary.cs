using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace OlapPivotTableExtensions
{
    public class CalculationsLibrary
    {
        public static string LibraryDirectory
        {
            get
            {
                return System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + "\\OlapPivotTableExtensions";
            }
        }

        public static string LibraryPath
        {
            get
            {
                return LibraryDirectory + "\\CalculationsLibrary.xml";
            }
        }

        /// <summary>
        /// Deserializes from the library XML
        /// </summary>
        public void Load()
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(CalculationsLibrary), null, new Type[] { typeof(Calculation) }, null, null);
                XmlTextReader reader = new XmlTextReader(new System.IO.StringReader(System.IO.File.ReadAllText(LibraryPath)));
                CalculationsLibrary library = (CalculationsLibrary)serializer.Deserialize(reader);
                this._Calculations = library.Calculations;
                reader.Close();
            }
            catch (Exception ex)
            {
                string s = ex.Message;
                s = s + "";
            }
        }

        /// <summary>
        /// Deserializes from the library XML
        /// </summary>
        public void Load(string Path)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CalculationsLibrary), null, new Type[] { typeof(Calculation) }, null, null);
            XmlTextReader reader = new XmlTextReader(new System.IO.StringReader(System.IO.File.ReadAllText(Path)));
            CalculationsLibrary library = (CalculationsLibrary)serializer.Deserialize(reader);
            this._Calculations = library.Calculations;
            reader.Close();
        }

        /// <summary>
        /// Serializes to the library XML
        /// </summary>
        public void Save()
        {
            if (!System.IO.Directory.Exists(LibraryDirectory))
            {
                System.IO.Directory.CreateDirectory(LibraryDirectory);
            }
            XmlSerializer serializer = new XmlSerializer(typeof(CalculationsLibrary), null, new Type[] { typeof(Calculation) }, null, null);
            XmlTextWriter writer = new XmlTextWriter(LibraryPath, Encoding.UTF8);
            writer.Formatting = Formatting.Indented;
            serializer.Serialize(writer, this);
            writer.Close();
        }

        /// <summary>
        /// Serializes to the library XML
        /// </summary>
        public void Save(string Path)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CalculationsLibrary), null, new Type[] { typeof(Calculation) }, null, null);
            XmlTextWriter writer = new XmlTextWriter(Path, Encoding.UTF8);
            writer.Formatting = Formatting.Indented;
            serializer.Serialize(writer, this);
            writer.Close();
        }

        private Calculation[] _Calculations;
        public Calculation[] Calculations
        {
            get { return _Calculations ?? new Calculation[] { }; }
            set {
                if (value != null)
                {
                    List<Calculation> list = new List<Calculation>(value);
                    list.Sort();
                    _Calculations = list.ToArray();
                }
            }
        }

        public void AddCalculation(string Name, string Formula)
        {
            List<Calculation> list;
            if (_Calculations != null)
                list = new List<Calculation>(_Calculations);
            else
                list = new List<Calculation>();
            foreach (Calculation c in list)
            {
                if (c.Name.Equals(Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    c.Name = Name;
                    c.Formula = Formula;
                    return;
                }
            }
            list.Add(new Calculation(Name, Formula));
            list.Sort();
            _Calculations = list.ToArray();
        }

        public Calculation GetCalculation(string Name)
        {
            if (_Calculations != null)
            {
                foreach (Calculation c in _Calculations)
                {
                    if (c.Name.Equals(Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        return c;
                    }
                }
            }
            return null;
        }

        public void DeleteCalculation(string Name)
        {
            List<Calculation> list;
            if (_Calculations != null)
                list = new List<Calculation>(_Calculations);
            else
                list = new List<Calculation>();
            for (int i = 0; i < list.Count; i++)
            {
                Calculation c = list[i];
                if (c.Name.Equals(Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    list.Remove(c);
                    _Calculations = list.ToArray();
                    return;
                }
            }
            throw new Exception(string.Format("Calculation {0} not found", Name));
        }

        public class Calculation : IComparable
        {
            public Calculation() { }
            public Calculation(string Name, string Formula)
            {
                _Name = Name;
                _Formula = Formula;
            }

            private string _Name;
            [XmlAttribute()]
            public string Name
            {
                get { return _Name; }
                set { _Name = value; }
            }

            private string _Formula;
            public string Formula
            {
                get { return _Formula; }
                set { _Formula = value; }
            }

            public int CompareTo(object obj)
            {
                if (obj == null || !(obj is Calculation)) throw new ArgumentException("Object was not a Calculation");
                Calculation c = (Calculation)obj;
                return (this._Name.CompareTo(c._Name));
            }
        }
    }

}
