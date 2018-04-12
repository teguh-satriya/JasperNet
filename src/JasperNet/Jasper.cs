using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace JasperNet
{
    public class JasperDbConnection : ConfigurationSection
    {
        [ConfigurationProperty("Driver", DefaultValue = "generic", IsRequired = true)]
        public string Driver
        {
            get
            {
                return (string)this["Driver"];
            }
            set
            {
                this["Driver"] = value;
            }
        }

        [ConfigurationProperty("JdbcDriver")]
        public string JdbcDriver
        {
            get
            {
                return (string)this["JdbcDriver"];
            }
            set
            {
                this["JdbcDriver"] = value;
            }
        }

        [ConfigurationProperty("JdbcUrl")]
        public string JdbcUrl
        {
            get
            {
                return (string)this["JdbcUrl"];
            }
            set
            {
                this["JdbcUrl"] = value;
            }
        }

        [ConfigurationProperty("Database")]
        public string Database
        {
            get
            {
                return (string)this["Database"];
            }
            set
            {
                this["Database"] = value;
            }
        }

        [ConfigurationProperty("Port")]
        public string Port
        {
            get
            {
                return (string)this["Port"];
            }
            set
            {
                this["Port"] = value;
            }
        }

        [ConfigurationProperty("Username", IsRequired = true)]
        public string Username
        {
            get
            {
                return (string)this["Username"];
            }
            set
            {
                this["Username"] = value;
            }
        }

        [ConfigurationProperty("Password")]
        public string Password
        {
            get
            {
                return (string)this["Password"];
            }
            set
            {
                this["Password"] = value;
            }
        }
    }

    public class JasperOptions
    {
        public string Input { get; set; }
        public string Output { get; set; }
        public string RenderPath { get; set; }
        public string[] Format { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
        public JasperDbConnection DbConnection { get; set; }

        public JasperOptions() {
            Format = new string[] { };
            Parameters = new Dictionary<string, string>();
            DbConnection = new JasperDbConnection();
        }
    }

    public class Jasper
    {
        protected string appPath;
        protected string jasperStarterPath;
        protected string jasperFilePath;
        protected string reportPath;
        protected string reportName;
        protected string command;

        protected string[] format = new string[] {
            "pdf", "rtf", "xls", "xlsx", "docx", "odt", "ods", "pptx", "csv", "html", "xhtml", "xml", "jrprint"
        };

        protected string[] supportedJava = new string[] {"1.7","1.8"};

        public Jasper()
        {
            this.appPath = HttpContext.Current.Server.MapPath("~");

            RegistryKey localMachineRegistry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                                                        Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);
            RegistryKey jreKey = localMachineRegistry.OpenSubKey("SOFTWARE\\JavaSoft\\Java Runtime Environment");

            if (jreKey == null)
            {
                throw new Exception("JasperNet require Java Version 7 or above");
            }

            string jreVer = jreKey.GetValue("CurrentVersion").ToString();

            if (!this.supportedJava.Contains(jreVer))
            {
                throw new Exception("JasperNet require JRE/JDK Version 7 or above");
            }

            //Set file path untuk jasperstarter
            if (string.IsNullOrEmpty(this.jasperStarterPath))
            {
                this.jasperStarterPath = this.appPath + "JasperStarter\\bin\\jasperstarter";
            }
            else
            {
                this.jasperStarterPath = this.appPath + this.jasperStarterPath;
            }
            
            if (!File.Exists(this.jasperStarterPath))
            {
                throw new FileNotFoundException("JasperStarter file missing!");
            }

            //Set folder path untuk file jasper
            this.jasperFilePath = this.appPath + "Reports\\";
            if (!Directory.Exists(this.jasperFilePath))
            {
                Directory.CreateDirectory(this.jasperFilePath);
            }
        }

        public void Process(JasperOptions config)
        {
            string renderCommand = this.jasperStarterPath + " pr ";
            string jasperFile = this.jasperFilePath + config.Input;
            string renderPath = this.appPath + config.RenderPath;
            string reportFile;
            string renderFile;

            if (!File.Exists(jasperFile))
            {
                throw new FileNotFoundException("Jasper file missing!");
            }

            if (string.IsNullOrEmpty(config.Output))
            {
                throw new Exception("Please specify output");
            }

            if (config.Format.Length == 0)
            {
                throw new Exception("Please specify format file");
            }

            foreach (var ext in config.Format)
            {
                if (!this.format.Contains(ext))
                {
                    throw new Exception("Format Render not supported");
                }
            }

            if (!Directory.Exists(renderPath))
            {
                Directory.CreateDirectory(renderPath);
            }
            
            reportFile = renderPath + config.Output;
            
            renderFile = String.Join(" ", config.Format);

            this.command = renderCommand + jasperFile + " -o " + reportFile + " -f " + renderFile;
            
            if (config.Parameters.Count > 0)
            {
                this.command = this.command + " -P";
                foreach (KeyValuePair<string, string> param in config.Parameters)
                {
                    string value;
                    if (param.Value.Contains(" "))
                    {
                        value = "\"" + param.Value + "\"";
                    }
                    else
                    {
                        value = param.Value;
                    }

                    this.command = this.command + " " + param.Key + "=" + value;
                }
            }

            if (!string.IsNullOrEmpty(config.DbConnection.Driver) && !string.IsNullOrEmpty(config.DbConnection.Username))
            {
                this.command = this.command + " -t " + config.DbConnection.Driver;
                this.command = this.command + " -u " + config.DbConnection.Username;

                if (!string.IsNullOrEmpty(config.DbConnection.Password))
                {
                    this.command = this.command + " -p " + config.DbConnection.Password;
                }
                if (!string.IsNullOrEmpty(config.DbConnection.Database))
                {
                    this.command = this.command + " -n " + config.DbConnection.Database;
                }
                if (!string.IsNullOrEmpty(config.DbConnection.Port))
                {
                    this.command = this.command + " --db-port " + config.DbConnection.Port;
                }
                if (!string.IsNullOrEmpty(config.DbConnection.JdbcDriver))
                {
                    this.command = this.command + " --db-driver " + config.DbConnection.JdbcDriver;
                }
                if (!string.IsNullOrEmpty(config.DbConnection.JdbcUrl))
                {
                    this.command = this.command + " --db-url " + config.DbConnection.JdbcUrl;
                }
            }

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.FileName = "CMD.exe";
            startInfo.Arguments = "/C " + this.command;
            process.StartInfo = startInfo;
            process.StartInfo.UseShellExecute = false;
            process.Start();
            process.WaitForExit();

            if (config.Format.Any())
            {
                string format = config.Format.First();
                this.reportPath = reportFile + "." + format;
                this.reportName = config.Output + "." + format;
            }
        }

        public FileContentResult Download(bool delete = true)
        {
            if (!File.Exists(this.reportPath))
            {
                throw new FileNotFoundException("Rendering report failed");
            }

            var fileBytes = File.ReadAllBytes(this.reportPath);

            if (delete)
            {
                File.Delete(this.reportPath);
            }

            var response = new FileContentResult(fileBytes, "application/octet-stream")
            {
                FileDownloadName = this.reportName
            };

            return response;
        }


    }
}
