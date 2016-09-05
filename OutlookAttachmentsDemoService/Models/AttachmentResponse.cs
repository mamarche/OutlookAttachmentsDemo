using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAttachmentsDemoService.Models
{
    public class AttachmentResponse
    {
        public string[] attachmentNames { get; set; }
        public int attachmentsProcessed { get; set; }
    }
}