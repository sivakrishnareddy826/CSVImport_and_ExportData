using CsvHelper;
using CsvHelper.Configuration;
using Dapper;
using OfficeOpenXml;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ExcelPractice
{
    public partial class Form1 : Form
    {
        private readonly IDbConnection dbConnection;
        public Form1()
        {
            InitializeComponent();
            // Replace "YourConnectionString" with your actual MySQL connection string
            var connectionFactory = new DbConnectionFactory("Server=localhost;Database=newdb;Uid=root;Pwd=root;");
            dbConnection = connectionFactory.CreateConnection();
            Load += EmployeeForm_Load;
            btnImport.Click += btnImport_Click; // Subscribe to the Import button click event
            btnExport.Click += btnExport_Click;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var newEmployee = new Employee
            {
                Name = txtName.Text,
                Role = txtRole.Text,
                Salary = Convert.ToDouble(txtSalary.Text),
                Gender = radioButton1.Checked ? "Male" : "Female",
                Dob = dateTimePickerDob.Value,
               // Status =textBox4.Text,
               
            };
            string sql = @"INSERT INTO Employee (Name, Role, Salary, Gender, Dob) 
                       VALUES (@Name, @Role, @Salary, @Gender, @Dob)";

            dbConnection.Execute(sql, newEmployee);

            MessageBox.Show("Employee added successfully.");
            ClearFields();
            LoadEmployees();
        }
        private void ClearFields()
        {
            //txtId.Clear();
            txtName.Clear();
            txtRole.Clear();
            txtSalary.Clear();
            dateTimePickerDob.Value = DateTime.Now;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }
        // Load event handler
        private void EmployeeForm_Load(object sender, EventArgs e)
        {
            LoadEmployees();
        }
        private void LoadEmployees()
        {
            string sql = "SELECT * FROM Employee";

            var employees = dbConnection.Query<Employee>(sql).ToList();

            /*        // Append '*' to the properties of the first object
                    if (employees.Count > 0)
                    {
                        Employee firstEmployee = employees[0];

                        // Assuming "Name" is a property of the Employee class

                        firstEmployee.Name = "*" + firstEmployee.Name;

                        firstEmployee.Gender = "*" + firstEmployee.Gender;
                        firstEmployee.Role = "*" + firstEmployee.Role;
                        // Repeat this for other properties you want to modify
                    }*/
            dataGridView1.DataSource = employees;
        }




        private void button2_Click(object sender, EventArgs e)
        {
            txtName.Clear();
            txtRole.Clear();
            txtSalary.Clear();
            dateTimePickerDob.Value = DateTime.Now;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV Files|*.csv|All Files|*.*";
                openFileDialog.Title = "Select CSV File to Import";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    try
                    {
                        ImportDataFromCsv<Employee>(filePath);
                        LoadEmployees(); // Reload the employees after importing
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error importing data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ImportEmployeesFromCsv(string filePath)
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    var employees = csv.GetRecords<Employee>().ToList();

                    string sql = @"INSERT INTO Employee (Id,Name, Role, Salary, Gender, Dob) 
                           VALUES (@Id,@Name, @Role, @Salary, @Gender, @Dob)";

                    dbConnection.Execute(sql, employees);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportDataFromCsv<T>(string filePath)
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    // Dynamically get the properties of the specified type T
                    var properties = typeof(T).GetProperties();

                    // Dynamically create the mapping using property names
                    var classMap = new DefaultClassMap<T>();

                    foreach (var property in properties)
                    {
                        classMap.Map(typeof(T), property);
                    }

                    csv.Context.RegisterClassMap(classMap);

                    // Read the records and map them to the specified type
                    var records = csv.GetRecords<T>().ToList();

                    // Build the parameterized SQL insert statement
                    var sql = @"INSERT INTO Employee (Id,Name, Role, Salary, Gender, Dob) 
                           VALUES (@Id,@Name, @Role, @Salary, @Gender, @Dob)";

                    // Execute the dynamic query
                    dbConnection.Execute(sql, records);
                    //dbConnection.Execute(sql, employees);
                    MessageBox.Show($"Imported {records.Count} records successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV Files|*.csv|All Files|*.*";
                saveFileDialog.Title = "Save CSV File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    try
                    {
                        ExportDataToCsv(filePath);
                        MessageBox.Show("Export successful.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportDataToCsv(string filePath)
        {
            try
            {
                // Get the List<Employee> from the DataGridView's DataSource
                var records = (List<Employee>)dataGridView1.DataSource;

                // Export the List<Employee> to CSV
                using (var writer = new StreamWriter(filePath))
                using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    csv.WriteRecords(records);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
