using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Timers;
using System.Runtime.InteropServices;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using FactureAutoImport.Apihelper;
using System.Configuration;

namespace FactureAutoImport
{
    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public int dwServiceType;
        public ServiceState dwCurrentState;
        public int dwControlsAccepted;
        public int dwWin32ExitCode;
        public int dwServiceSpecificExitCode;
        public int dwCheckPoint;
        public int dwWaitHint;
    };
    public partial class AutoImport : ServiceBase
    {
        private int eventId = 1;
        public AutoImport()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("AutoImportSource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "AutoImportSource", "AutoImportLog");
            }
            eventLog1.Source = "AutoImportSource";
            eventLog1.Log = "AutoImportLog";
        }

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry("In OnStart.");
            try
            {
                GetAllEmails(Convert.ToString(ConfigurationManager.AppSettings["HostAddress"]));
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.InnerException.Message, EventLogEntryType.Error, eventId++);
            }

            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 10000; // 10 seconds
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        
            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop.");
        }
        protected override void OnContinue()
        {
            eventLog1.WriteEntry("In OnContinue.");
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            // TODO: Insert monitoring activities here.
            eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);
            System.Threading.Thread.Sleep(10000);


            try
            {
                 GetAllEmails(Convert.ToString(ConfigurationManager.AppSettings["HostAddress"]));
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.InnerException.Message, EventLogEntryType.Error, eventId++);
            }
            
        }

        public  void GetAllEmails(string HostEmailAddress)
        {
            try
            {
                GmailService GmailService = GmailAPIHelper.GetService();
                List<Gmail> EmailList = new List<Gmail>();
                UsersResource.MessagesResource.ListRequest ListRequest = GmailService.Users.Messages.List(HostEmailAddress);
                ListRequest.LabelIds = "Label_8051553239077085895";
                ListRequest.IncludeSpamTrash = false;
                ListRequest.Q = "is:unread"; //only from unread emails

                //GET ALL EMAILS
                ListMessagesResponse ListResponse = ListRequest.Execute();

                if (ListResponse != null && ListResponse.Messages != null)
                {
                    //loop throught each email
                    foreach (Message Msg in ListResponse.Messages)
                    {
                        //Mark message as read
                        GmailAPIHelper.MsgMarkAsRead(HostEmailAddress, Msg.Id);

                        UsersResource.MessagesResource.GetRequest Message = GmailService.Users.Messages.Get(HostEmailAddress, Msg.Id);
                       // Console.WriteLine("\n-----------------NEW MAIL----------------------");
                        eventLog1.WriteEntry("NEW MAIL", EventLogEntryType.Information, eventId++);
                       

                        //MAKE ANOTHER REQUEST FOR THAT EMAIL ID...
                        Message MsgContent = Message.Execute();

                        if (MsgContent != null)
                        {
                            string FromAddress = string.Empty;
                            string Date = string.Empty;
                            string Subject = string.Empty;
                            string MailBody = string.Empty;
                            string ReadableText = string.Empty;

                            //LOOP THROUGH THE HEADERS AND GET THE FIELDS WE NEED (SUBJECT, MAIL)
                            foreach (var MessageParts in MsgContent.Payload.Headers)
                            {
                                if (MessageParts.Name == "From")
                                {
                                    FromAddress = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Date")
                                {
                                    Date = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Subject")
                                {
                                    Subject = MessageParts.Value;
                                }
                            }
                            //READ MAIL BODY
                            Console.WriteLine("STEP-2: Read Mail Body");
                            List<string> FileName = GmailAPIHelper.GetAttachments(HostEmailAddress, Msg.Id, Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]));
                            if (FileName != null)
                            {
                                eventLog1.WriteEntry("mail de : " + FromAddress, EventLogEntryType.Information, eventId++);
                            }
                          
                        }
                    }
                }
              
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.Message, EventLogEntryType.Error, eventId++);
            }
        }

        internal void TestStartupAndStop(string[] args)
        {
            this.OnStart(args);
            Console.ReadLine();
            this.OnStop();
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);
    }
}
