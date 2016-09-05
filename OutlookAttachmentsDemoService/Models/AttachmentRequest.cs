using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAttachmentsDemoService.Models
{
    public class AttachmentRequest
    {
        public string attachmentToken { get; set; }
        public string ewsUrl { get; set; }
        public string service { get; set; }
        public AttachmentDetail[] attachments { get; set; }
    }
}