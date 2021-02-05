using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Threading;

namespace Json_keyname_validator
{
    public partial class Form1 : Form
    {
        OpenFileDialog OpenFileDialog = new OpenFileDialog();
        dynamic fieldNames = new List<string>();
        dynamic validationExcelFilePath = "";
        string filename = "";
        string jsonString = "";
        string jsonFilePath = "";
        string excelFieldNamelength = "";
        string title = "Validation Console";


        public Form1()
        {
            InitializeComponent();
            label1.Text = "No file chosen"; 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog.ShowDialog();
                jsonFilePath = OpenFileDialog.FileName;
                if (jsonFilePath.Contains(".json"))
                {
                    filename = Path.GetFileName(jsonFilePath);
                    fieldNames = new List<string>();
                    
                    using (StreamReader r = new StreamReader(jsonFilePath))
                    {
                        jsonString = r.ReadToEnd();
                        bool jsonValidator = IsJson(jsonString);
                        if (jsonValidator == true)
                        {
                            label1.Text = "JSON File Name:" + ' ' + filename;
                            label1.Show();
                            fieldNames = GetFieldNames(jsonString);
                          
                            button2.Focus();
                        }
                        else
                        {
                            MessageBox.Show("JSON file contents invalid data", title);
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Please select a valid JSON file.", title);
                    label1.Text = "No File Chosen";
                    button1.Focus();
                }
               
            }
            catch (Exception ex)
            {

            }

        }
        private int FieldNameCount(List<string> fieldname)
        {
            int count = filename.Length;
            
            return count;
        }
        private static List<string> GetFieldNames(dynamic input)
        {
            List<string> fieldNames = new List<string>();

            try
            {
                // Deserialize the input json string to an object
                input = Newtonsoft.Json.JsonConvert.DeserializeObject(input);

                // Json Object could either contain an array or an object or just values
                // For the field names, navigate to the root or the first element
                input = input.Root ?? input.First ?? input;
                if (input != null)
                {

                    //// Get to the first element in the array
                    bool isArray = true;

                    
                    while (isArray)
                    {
                     
                        input =  input.First ?? input;

                        if (input.GetType() == typeof(Newtonsoft.Json.Linq.JObject) ||
                        input.GetType() == typeof(Newtonsoft.Json.Linq.JValue) || 
                        input == null)
                            isArray = false;
                    }

                    // check if the object is of type JObject. 
                    // If yes, read the properties of that JObject
                    if (input.GetType() == typeof(Newtonsoft.Json.Linq.JObject))
                    {
                        // Create JObject from object
                        Newtonsoft.Json.Linq.JObject inputJson =
                            Newtonsoft.Json.Linq.JObject.FromObject(input);

                        // Read Properties
                        var properties = inputJson.Properties();

                        // Loop through all the properties of that JObject
                        foreach (var property in properties)
                        {
                            // Check if there are any sub-fields (nested)
                            // i.e. the value of any field is another JObject or another JArray
                            if (property.Value.GetType() == typeof(Newtonsoft.Json.Linq.JObject) ||
                            property.Value.GetType() == typeof(Newtonsoft.Json.Linq.JArray))
                            {
                                // If yes, enter the recursive loop to extract sub-field names
                                var subFields = GetFieldNames(property.Value.ToString());

                                if (subFields != null && subFields.Count() > 0)
                                {
                                    // join sub-field names with field name 
                                    //(e.g. Field1.SubField1, Field1.SubField2, etc.)
                                    //fieldNames.AddRange(
                                    //    subFields
                                    //    .Select(n =>
                                    //    string.IsNullOrEmpty(n) ? property.Name :
                                    // string.Format("{0}.{1}", property.Name, n)));
                                    fieldNames.AddRange(
                                        subFields
                                        .Select(n =>
                                        string.IsNullOrEmpty(n) ? property.Name :
                                     string.Format("{0}", n)));
                                }
                            }
                            else
                            {
                                // If there are no sub-fields, the property name is the field name
                                fieldNames.Add(property.Name);
                            }
                        }
                    }
                    else
                        if (input.GetType() == typeof(Newtonsoft.Json.Linq.JValue))
                    {
                        // for direct values, there is no field name
                        fieldNames.Add(string.Empty);
                    }
                }
            }
            catch
            {
                throw;
            }

            return fieldNames;
        }

        public static bool IsJson(string input)
        {
            input = input.Trim();
            return input.StartsWith("{") && input.EndsWith("}")
                   || input.StartsWith("[") && input.EndsWith("]");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog.ShowDialog();
                validationExcelFilePath = OpenFileDialog.FileName;
                string extension = Path.GetExtension(validationExcelFilePath);
                if (extension == ".xlsx")
                {
                    string filename = Path.GetFileName(validationExcelFilePath);
                    label2.Text = "Excel File Name:" + ' ' + filename;
                    button3.Focus();
                }
                else
                {
                    MessageBox.Show("Please select the standard key names template(.xlsx).", title);
                    label2.Text = "";
                    button2.Focus();
                }
                
            }
            catch (Exception ex)
            {

            }

        }
        private string excelFieldNameCount(string fieldname)
        {
            string count = fieldname;

            return count;
        }
        private bool ValidateExcel(List<string> fieldNames, string validationExcelFilePath)
        {

            DateTime time = DateTime.Now.ToLocalTime();
                string localTime = time.ToString();
                string[] splittime = localTime.Split('/', ':');
                string appendDateTime = "";

                for (int i = 0; i <= splittime.Length - 1; i++)
                {
                    appendDateTime += splittime[i].Trim().ToString();
                }
                string getPath = Path.GetDirectoryName(jsonFilePath);
                string filename = Path.GetFileName(jsonFilePath);
                filename = filename.Substring(0, filename.Length - 5);
                filename = getPath + "\\" + appendDateTime + " " + filename + ".xlsx";
                File.Copy(validationExcelFilePath, filename, true);
                Application application = new Application();

                Workbook workbook = application.Workbooks.Open(filename,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);
                Worksheet sheet = workbook.Sheets["Keyword Validator"];
            var workSheets = new List<string>();
           

            try
                {
                    for (int row = 2; row <= 1000; row++)
                    {
                        var column1 = (string)((Range)sheet.Cells[row, 1]).Value;
                        var matches = false;
                        if (column1 != null)
                        {
                            matches = fieldNames.Any(x => column1.Trim().ToLower().IndexOf(x.Trim().ToLower()) > -1);
                        }

                        if (matches)
                        {
                            ((Range)sheet.Cells[row, 2]).Value = column1;
                            fieldNames.RemoveAll(x => ((string)x) == column1);
                        }
                    }
                    int n = 0;
               
                    for (int r = 2 ; r <= 100/*fieldNames.Count*/; r++)
                    {
                    
                        ((Range)sheet.Cells[r, 3]).Value = fieldNames[n];
                        n++;
                    
                    
                    }
                }
                catch (Exception ex)
                {

                }
                workbook.Save();
                workbook.Close(false, validationExcelFilePath, null);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                
            
            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            bool result1 = FileValidate();

            try
            {

                if (result1 == true)
                {
                    bool result = ValidateExcel(fieldNames, validationExcelFilePath);

                    if (result == true)
                    {
                        MessageBox.Show("Comparison Successfull.", "JSON Key Name Validator Success Console");
                        Thread.Sleep(2000);
                        label1.Text = "No File Chosen";
                        label2.Text = "";
                        validationExcelFilePath = "";
                        jsonFilePath = "";
                        excelFieldNamelength = "";
                    }
                }

            }
            catch (Exception ex)
            {

            }
        }

        private bool FileValidate()
        {
            bool result = true;
            bool valiate = false;

             excelFieldNamelength = excelFieldNameCount(validationExcelFilePath);
            while (result)
            {
                
                
                if (jsonFilePath.Contains(".json"))
                {
                    result = true;
                }
                else
               if (jsonFilePath.Contains(""))
                {
                    MessageBox.Show("Please select the JSON file.", title);
                    result = false;
                    break;
                }
                else
               if (FieldNameCount(fieldNames) == 0)
                {
                    MessageBox.Show("JSON file contents empty key name.", title);
                    result = false;
                    break;
                }
                if (excelFieldNamelength.Contains(".xlsx"))
                {
                    result = true;
                }
                else
                {
                    MessageBox.Show("Please select the standard key names template(.xlsx).", title);
                    result = false;
                    break;
                }
                if (result == true)
                {
                    result = false;
                    valiate = true;
                }
               
            }
           


            return valiate;
        }
    }
}

