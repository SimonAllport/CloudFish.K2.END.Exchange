using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace CloudFish.K2.END.Exchange
{
    public class Email
    {
        private static ExchangeService ConnectToExchange()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.Credentials = new WebCredentials("Username", "Password");

            service.Url = new Uri("Web service");
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "Email Address");

            return service;
        }



        public static string SendEmail(string Subject,string Body, string To, string Cc,string Bcc,int importance, string sn,string Folio, int? ProcessInstanceId, string ProcessTypeId, string BusinessKey)
        {
            string result = string.Empty;
            ExchangeService service = ConnectToExchange();
            try
            {
                if (To != null || To.Length != 0)
                {
                    EmailMessage email = new EmailMessage(service);
                    email.Subject = Subject;
                    email.Body = new MessageBody(BodyType.HTML, Body);

                    Guid SN_PropertySetId = Guid.NewGuid();
                    ExtendedPropertyDefinition SN_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(SN_PropertySetId, "SN", MapiPropertyType.String);
                    email.SetExtendedProperty(SN_ExtendedPropertyDefinition, (!String.IsNullOrEmpty(sn) ? sn : "0_0"));

                    Guid    Folio_PropertySetId = Guid.NewGuid();
                    ExtendedPropertyDefinition Folio_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(Folio_PropertySetId, "Folio", MapiPropertyType.String);
                    email.SetExtendedProperty(Folio_ExtendedPropertyDefinition, (!String.IsNullOrEmpty(Folio) ? Folio : "Email Message"));

                    Guid ProcessInstanceId_PropertySetId = Guid.NewGuid();
                    ExtendedPropertyDefinition ProcessInstanceId_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(ProcessInstanceId_PropertySetId, "ProcessInstanceId", MapiPropertyType.String);
                    email.SetExtendedProperty(ProcessInstanceId_ExtendedPropertyDefinition, (ProcessInstanceId > 0 | ProcessInstanceId != null ? ProcessInstanceId : 0));

                    Guid BusinessKey_PropertySetId = Guid.NewGuid();
                    ExtendedPropertyDefinition BusinessKey_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(BusinessKey_PropertySetId, "BusinessKey", MapiPropertyType.String);
                    email.SetExtendedProperty(BusinessKey_ExtendedPropertyDefinition, (!String.IsNullOrEmpty(BusinessKey) ? BusinessKey : "0"));

                    Guid ProcessTypeId_PropertySetId = Guid.NewGuid();
                    ExtendedPropertyDefinition ProcessTypeId_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(ProcessTypeId_PropertySetId, "ProcessTypeId", MapiPropertyType.String);
                    email.SetExtendedProperty(ProcessTypeId_ExtendedPropertyDefinition, (!String.IsNullOrEmpty(ProcessTypeId) ? ProcessTypeId : "00000000-0000-0000-0000-000000000000"));

                    Guid MessageId_PropertySetId = Guid.NewGuid();
                    string MessageId = Guid.NewGuid().ToString();
                    ExtendedPropertyDefinition MessageId_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(MessageId_PropertySetId, "ProcessTypeId", MapiPropertyType.String);
                    email.SetExtendedProperty(MessageId_ExtendedPropertyDefinition, MessageId);


                    if (To.Contains(";"))
                    {
                        String[] to = To.Split(';');
                        foreach (var address in to)
                        {
                            email.ToRecipients.Add(address);
                        }
                    }
                    else
                    {
                        email.ToRecipients.Add(To);
                    }


                    if (!string.IsNullOrEmpty(Cc))
                    {
                        if (Cc.Contains(";"))
                        {
                            String[] to = Cc.Split(';');
                            foreach( var address in to)
                            {
                                email.CcRecipients.Add(address);
                            }
                        }
                        else
                        {
                            email.CcRecipients.Add(Cc);

                        }
                    }

                    if (!string.IsNullOrEmpty(Bcc))
                    {
                        if (Bcc.Contains(";"))
                        {
                            String[] to = Bcc.Split(';');
                            foreach (var address in to)
                            {
                                email.BccRecipients.Add(address);
                            }
                        }
                        else
                        {
                            email.BccRecipients.Add(Cc);

                        }
                    }

                    if (importance > 0)
                    {
                        email.Importance = (importance == 1 ? Microsoft.Exchange.WebServices.Data.Importance.Normal : Importance.High);
                    }

                    email.SendAndSaveCopy();

                    result = email.Id.ToString();
                }
            }
            catch(Exception ex)
            {
                result = "Error: " + ex.Message.ToString(); 
            }
            finally
            {

            }
            return result;
        }


        public static List<EmailBox> GetMailBox(string MailBoxType,int PageSize)
        {
            ItemView view = new ItemView(PageSize);
            List<EmailBox> list = new List<EmailBox>();

            Guid SN_PropertySetId = Guid.NewGuid();
            ExtendedPropertyDefinition SN_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(SN_PropertySetId, "SN", MapiPropertyType.String);
            
            Guid Folio_PropertySetId = Guid.NewGuid();
            ExtendedPropertyDefinition Folio_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(Folio_PropertySetId, "Folio", MapiPropertyType.String);
         
            Guid ProcessInstanceId_PropertySetId = Guid.NewGuid();
            ExtendedPropertyDefinition ProcessInstanceId_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(ProcessInstanceId_PropertySetId, "ProcessInstanceId", MapiPropertyType.String);
           
            Guid BusinessKey_PropertySetId = Guid.NewGuid();
            ExtendedPropertyDefinition BusinessKey_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(BusinessKey_PropertySetId, "BusinessKey", MapiPropertyType.String);
           
            Guid ProcessTypeId_PropertySetId = Guid.NewGuid();
            ExtendedPropertyDefinition ProcessTypeId_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(ProcessTypeId_PropertySetId, "ProcessTypeId", MapiPropertyType.String);
            
            Guid MessageId_PropertySetId = Guid.NewGuid();
            string MessageId = Guid.NewGuid().ToString();
            ExtendedPropertyDefinition MessageId_ExtendedPropertyDefinition = new ExtendedPropertyDefinition(MessageId_PropertySetId, "ProcessTypeId", MapiPropertyType.String);

            ExchangeService service = ConnectToExchange();
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, SN_ExtendedPropertyDefinition, Folio_ExtendedPropertyDefinition, ProcessInstanceId_ExtendedPropertyDefinition, BusinessKey_ExtendedPropertyDefinition, ProcessTypeId_ExtendedPropertyDefinition, MessageId_ExtendedPropertyDefinition);

            FindItemsResults<Item> findResults = service.FindItems((MailBoxType == "Sent" ? WellKnownFolderName.SentItems : WellKnownFolderName.Inbox), view);
            foreach(Item email in findResults.Items)
            {
                Item mail = Item.Bind(service, email.Id);
                list.Add(new EmailBox
                {
                    MailBoxType = MailBoxType,
                    Subject = mail.Subject,
                    Body = mail.Body,
                    Importance = mail.Importance.ToString(),
                Id = mail.Id.ToString(),
                Categories = mail.Categories.ToString(),
                DateTimeCreated = mail.DateTimeCreated,
                DateTimeReceived = mail.DateTimeReceived,
                DateTimeSent = mail.DateTimeSent,
                Cc = mail.DisplayCc,
                To = mail.DisplayTo,
                SN = (email.ExtendedProperties.Count > 0 ? email.ExtendedProperties[0].Value.ToString():string.Empty),
                Folio = (email.ExtendedProperties.Count > 0 ? email.ExtendedProperties[1].Value.ToString(): string.Empty),
                ProcessInstanceId = (email.ExtendedProperties.Count > 0 ? email.ExtendedProperties[2].Value.ToString(): string.Empty),
                BusinessKey = (email.ExtendedProperties.Count > 0 ? email.ExtendedProperties[3].Value.ToString(): string.Empty),
                ProcessTypeId = (email.ExtendedProperties.Count > 0 ? email.ExtendedProperties[4].Value.ToString(): string.Empty),
                MessageId = (email.ExtendedProperties.Count > 0 ? email.ExtendedProperties[5].Value.ToString(): string.Empty)
                
                    });

            }
            return list;




        }

        public static EmailBox GetEmail(string Id)
        {
            EmailBox email = new EmailBox();
            ExchangeService service = ConnectToExchange();

            try
            {
                Item mail = Item.Bind(service, (ItemId)Id);
                {

                  
                    email.Subject = mail.Subject;
                    email.Body = mail.Body;
                    email.Importance = mail.Importance.ToString();
                    email.Id = mail.Id.ToString();
                    email.Categories = mail.Categories.ToString() ;
                    email.DateTimeCreated = mail.DateTimeCreated;
                    email.DateTimeReceived = mail.DateTimeReceived;
                    email.DateTimeSent = mail.DateTimeSent;
                    email.Cc = mail.DisplayCc;
                    email.To = mail.DisplayTo;
                    email.SN = (mail.ExtendedProperties.Count > 0 ? mail.ExtendedProperties[0].Value.ToString(): string.Empty);
                    email.Folio = (mail.ExtendedProperties.Count > 0 ? mail.ExtendedProperties[1].Value.ToString(): string.Empty);
                    email.ProcessInstanceId = (mail.ExtendedProperties.Count > 0 ? mail.ExtendedProperties[2].Value.ToString(): string.Empty);
                    email.BusinessKey = (mail.ExtendedProperties.Count > 0 ? mail.ExtendedProperties[3].Value.ToString(): string.Empty);
                    email.ProcessTypeId = (mail.ExtendedProperties.Count > 0 ? mail.ExtendedProperties[4].Value.ToString(): string.Empty);
                    email.MessageId = (mail.ExtendedProperties.Count > 0 ? mail.ExtendedProperties[5].Value.ToString(): string.Empty);


                }


            }
            catch(Exception ex)
            { }
            finally
            {

            }
            return email;
        }

        public static Boolean DeleteEmail(string Id, string DeleteType)
        {
            Boolean Result = false;
            ExchangeService service = ConnectToExchange();
            try
            {
                Item mail = Item.Bind(service, (ItemId)Id);
                mail.Delete(DeleteMode.MoveToDeletedItems);
                Result = true;
            }
            catch(Exception ex)
            { }
            finally
            {

            }
            return Result;
        }
    }

    public class EmailBox
    {
        public string Body { get; set; }
        public string Subject { get; set; }
        public string Importance { get; set; }
        public string ProcessTypeId { get; set; }
        public string Id { get; set; }
        public string Categories { get; set; }
        public DateTime DateTimeCreated { get; set; }
        public DateTime DateTimeReceived { get; set; }
        public DateTime DateTimeSent { get; set; }
        public string Cc { get; set; }
        public string SN { get; set; }
        public string To { get; set; }
        public string Folio { get; set; }
        public string ProcessInstanceId { get; set; }
        public string BusinessKey { get; set; }
        public string MessageId { get; set; }
        public string MailBoxType { get; set; }
    }
}
