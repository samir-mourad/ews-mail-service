using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MailService
{
    public class MailMessage
    {
        private Regex _pattern;

        public string UniqueId { get; set; }
        public string FromName { get; set; }
        public string FromMail { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string ConversationId { get; set; }
        public DateTime SentDate { get; set; }
        public DateTime ReceivedDate { get; set; }
        public bool HasAttachments { get; set; }

        public IEnumerable<AttachmentMessage> Attachments { get; set; }

        public string Filename => $"{ReceivedDate:yyyyMMddHHmmss}_{_pattern.Replace(Subject.Trim(), string.Empty)}";

        public MailMessage()
        {
            _pattern = new Regex(@"/| |\|-|:");
        }

    }

    public class AttachmentMessage
    {
        public string ContentId { get; set; }
        public string ContentType { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }

        public AttachmentMessage(Attachment attachment)
        {
            ContentId = attachment.ContentId;
            ContentType = attachment.ContentType;
            Id = attachment.Id;
            Name = attachment.Name;
        }
    }
}
