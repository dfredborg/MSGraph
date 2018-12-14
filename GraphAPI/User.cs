using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace GraphAPI
{
    public class User : TableEntity
    {
        public string name { get; set; }
        public string accessToken { get; set; }
        public string refreshToken { get; set; }
    }
}
