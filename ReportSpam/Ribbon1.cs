using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ReportSpam
{
    public partial class ReportSpam
    {
        private string path = @"c:\ReportSpamOfficeAddIn\";
        private string fileName = "settings.ini";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //if the settings file does not exist, create it
            if (!File.Exists(path+fileName))
            {
                CreateFile();
            }
        }

        private void CreateFile()
        {
            try
            {
                Directory.CreateDirectory(path);
                File.WriteAllText(path + fileName, "email@address.com");
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show(@"Error creating/writing to email settings file "+path+fileName);
            }
        }

        private string ReadFile()
        {
            try
            {
                return File.ReadAllText(path + fileName);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Unable to read settings file "+path+fileName);
                return "";
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application application = new Outlook.Application();
            Outlook.NameSpace ns = application.GetNamespace("MAPI");

            try
            {
                //get selected mail item
                Object selectedObject = application.ActiveExplorer().Selection[1];
                Outlook.MailItem selectedMail = (Outlook.MailItem)selectedObject;

                //create message
                Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                newMail.Recipients.Add(ReadFile());
                newMail.Subject = "SPAM";
                newMail.Attachments.Add(selectedMail, Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem);

                newMail.Send();
                selectedMail.Delete();

                System.Windows.Forms.MessageBox.Show("Spam notification has been sent.");
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("You must select a message to report.");
            }
        }
    }
}
