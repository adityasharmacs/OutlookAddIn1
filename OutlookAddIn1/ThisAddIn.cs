﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.SharePoint.Client;
using mapi = Microsoft.Office.Interop.Outlook.MAPIFolder;
using System.Text.RegularExpressions;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;

            //for (int i = 0; i < inspectors.Count; i++)
            //{
            //    Outlook.Inspector inspec = inspectors[i];

            //    object current = inspec.CurrentItem;
                                
            //    //String inspectorPresent = inspectors[i]
            //}

            //inspectors.

            inspectors.NewInspector += Inspectors_NewInspector;
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;

            Outlook.Links links = mailItem.Links;
            
            Outlook.Recipients recipients = mailItem.Recipients;
            
            String body = mailItem.HTMLBody;

            //var matches = Regex.Matches(body, @"<a\shref=""(?<url>.*?)"">(?<text>.*?)</a>");
            //foreach (Match match in matches)
            //{
            //    Console.WriteLine("URL: " + match.Groups["url"].Value + " -- Text = " + match.Groups["text"].Value);
            //    System.Windows.Forms.MessageBox.Show("URL: " + match.Groups["url"].Value + " -- Text = " + match.Groups["text"].Value);
            //}

            //((([A-Za-z]{3,9}:(?:\/\/)?)(?:[\-;:&=\+\$,\w]+@)?[A-Za-z0-9\.\-]+|(?:www\.|[\-;:&=\+\$,\w]+@)[A-Za-z0-9\.\-]+)((?:\/[\+~%\/\.[?sharepoint?]\-_]*)?\??(?:[\-\+=&;%@\.\w_]*)#?(?:[\.\!\/\\\w]*))?)

            var matches = Regex.Matches(body, @"<a\shref=""((([A-Za-z]{3,9}:(?:\/\/)?)(?:[\-;:&=\+\$,\w]+@)?[A-Za-z0-9\.\-]+|(?:www\.|[\-;:&=\+\$,\w]+@)[A-Za-z0-9\.\-]+)((?:\/[\+~%\/\.[?sharepoint.com?]\-_]*)?\??(?:[\-\+=&;%@\.\w_]*)#?(?:[\.\!\/\\\w]*))?)</a>");
            foreach (Match match in matches)
            {
                System.Windows.Forms.MessageBox.Show(match.Value);
            }



            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
