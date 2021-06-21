using System.Net;

namespace MailService
{
    public class MailCredential
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public string Domain { get; set; }
        public string Server { get; set; }

        public MailCredential(string username, string password, string domain, string server)
        {
            Username = username;
            Password = password;
            Domain = domain;
            Server = server;
        }

        public NetworkCredential GetNetworkCredential()
        {
            return new NetworkCredential(Username, Password, Domain);
        }
    }
}
