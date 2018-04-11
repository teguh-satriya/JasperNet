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
        public List<string> Format { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
        public JasperDbConnection DbConnection { get; set; }
    }

    public class Jasper
    {
        protected string appPath;
        protected string jasperStarterPath;
        protected string jasperFilePath;
        protected string reportPath;
        protected string reportName;

        protected List<string> format = new List<string>(new string[] {
            "pdf", "rtf", "xls", "xlsx", "docx", "odt", "ods", "pptx", "csv", "html", "xhtml", "xml", "jrprint"
        });

        public Jasper()
        {
            this.appPath = HttpContext.Current.Server.MapPath("~");

            //Set file path untuk jasperstarter
            this.jasperStarterPath = this.appPath + "\\JasperStarter\\bin\\jasperstarter";
            if (!File.Exists(this.jasperStarterPath))
            {
                throw new FileNotFoundException("JasperStarter file missing!");
            }

            //Set folder path untk file jasper
            this.jasperFilePath = this.appPath + "\\Reports\\";
            if (!Directory.Exists(this.jasperFilePath))
            {
                Directory.CreateDirectory(this.jasperFilePath);
            }
        }

        public void Compile(JasperOptions config)
        {
            string renderCommand = this.jasperStarterPath + " process ";
            string jasperFile = this.jasperFilePath + config.Input;
            string renderPath = this.appPath + config.RenderPath;
            string reportFile;
            string renderFile;
            string command;

            if (!File.Exists(jasperFile))
            {
                throw new FileNotFoundException("Jasper file missing!");
            }

            if (!Directory.Exists(renderPath))
            {
                Directory.CreateDirectory(renderPath);
            }

            if (config.Output == null)
            {
                throw new Exception("Please specify output");
            }

            reportFile = renderPath + config.Output;

            foreach (var ext in config.Format)
            {
                if (!this.format.Contains(ext))
                {
                    throw new Exception("Format Render not supported");
                }
            }
            renderFile = String.Join(" ", config.Format);

            command = renderCommand + jasperFile + " -o " + reportFile + " -f " + renderFile;

            if (config.Parameters != null)
            {
                command = command + " -P";
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

                    command = command + " " + param.Key + "=" + value;
                }
            }

            if (config.DbConnection.Driver != null && config.DbConnection.Username != null)
            {
                command = command + " -t " + config.DbConnection.Driver;
                command = command + " -u " + config.DbConnection.Username;

                if (config.DbConnection.Password != null)
                {
                    command = command + " -p " + config.DbConnection.Password;
                }
                if (config.DbConnection.Database != null)
                {
                    command = command + " -n " + config.DbConnection.Database;
                }
                if (config.DbConnection.JdbcDriver != null)
                {
                    command = command + " --db-driver " + config.DbConnection.JdbcDriver;
                }
                if (config.DbConnection.JdbcUrl != null)
                {
                    command = command + " --db-url " + config.DbConnection.JdbcUrl;
                }
            }

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.FileName = "CMD.exe";
            startInfo.Arguments = "/C " + command;
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

        public FileContentResult Render(bool delete = true)
        {
            if (!File.Exists(this.reportPath))
            {
                throw new FileNotFoundException("Report file not found");
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
