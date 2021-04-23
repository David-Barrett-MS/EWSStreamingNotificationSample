/*
 * By David Barrett, Microsoft Ltd. 2020. Use at your own risk.  No warranties are given.
 * 
 * DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 * */

using System;
using System.Net;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using EWS = Microsoft.Exchange.WebServices.Data;
using Autodiscover = Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Identity.Client;
using System.Windows.Forms;

namespace EWSStreamingNotificationSample.Auth
{
    public enum AuthType
    {
        None,
        Default,
        Basic,
        Certificate,
        OAuth
    }

    public class CredentialHandler
    {
        String _password = String.Empty;
        X509Certificate2 _certificate = null;
        AuthType _authType = AuthType.Basic;
        private AuthenticationResult _lastOAuthResult = null;
        private OAuthHelper _oAuthHelper = new OAuthHelper();
        private ClassLogger _logger = null;

        public CredentialHandler(AuthType authType, ClassLogger Logger=null)
        {
            _authType = authType;
            _logger = Logger;
        }

        public string Username { get; set; } = String.Empty;

        public string Domain { get; set; } = String.Empty;

        public string TenantId { get; set; } = String.Empty;

        public string ApplicationId { get; set; } = String.Empty;

        public string ClientSecret { get; set; } = String.Empty;

        public string Password
        {
            set { _password = value; }
        }

        public string OAuthToken
        {
            get {
                if (_lastOAuthResult == null)
                    return null;
                return _lastOAuthResult.AccessToken;
            }
        }

        public bool OAuthTokenExpired()
        {
            if (_lastOAuthResult != null && _lastOAuthResult.ExpiresOn > DateTime.Now)
                return false;
            return true;
        }

        public X509Certificate2 Certificate
        {
            get { return _certificate; }
            set { _certificate = value; }
        }


        public void GetAppTokenWithSecret()
        {
            try
            {
                _lastOAuthResult = Task.Run(async () => await OAuthHelper.GetApplicationToken(ApplicationId, TenantId, ClientSecret)).Result;
                _logger?.Log($"Token obtained, expires {_lastOAuthResult.ExpiresOn}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Unable to acquire token", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GetAppTokenWithCertificate()
        {
            try
            {
                _lastOAuthResult = Task.Run(async () => await OAuthHelper.GetApplicationToken(ApplicationId, TenantId, Certificate)).Result;
                _logger?.Log($"Token obtained, expires {_lastOAuthResult.ExpiresOn}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Unable to acquire token", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AcquireToken()
        {
            _logger?.Log("Requesting new access token");
            if (_certificate != null)
                GetAppTokenWithCertificate();
            else
                GetAppTokenWithSecret();
        }

        private bool HaveValidCredentials()
        {
            switch (_authType)
            {
                case AuthType.Default:
                    return true;

                case AuthType.Basic:
                    if (!String.IsNullOrEmpty(Username)) return true;
                    return false;

                case AuthType.Certificate:
                    return (_certificate != null);

                case AuthType.OAuth:
                    // Check if we already have valid access token
                    if (_lastOAuthResult != null && _lastOAuthResult.ExpiresOn > DateTime.Now) return true;

                    // We don't have an access token, so check the application configuration
                    if (String.IsNullOrEmpty(ApplicationId)) return false;
                    if (String.IsNullOrEmpty(TenantId)) return false;
                    if ((_certificate == null) && String.IsNullOrEmpty(ClientSecret)) return false;

                    // Obtain a token
                    AcquireToken();
                    if (_lastOAuthResult != null && _lastOAuthResult.ExpiresOn > DateTime.Now) return true;
                    return false;

                default: return false;
            }
        }

        public void LogCredentials(ClassLogger Logger)
        {
            StringBuilder sCredentialInfo = new StringBuilder();
            switch (_authType)
            {
                case AuthType.Default:
                    sCredentialInfo.AppendLine("Using default credentials");
                    sCredentialInfo.Append("Username: ");
                    sCredentialInfo.AppendLine(Environment.UserName);
                    sCredentialInfo.Append("Domain: ");
                    sCredentialInfo.AppendLine(Environment.UserDomainName);
                    break;

                case AuthType.Basic:
                    sCredentialInfo.AppendLine("Using specific credentials");
                    sCredentialInfo.Append("Username: ");
                    sCredentialInfo.AppendLine(Username);
                    sCredentialInfo.Append("Domain: ");
                    sCredentialInfo.AppendLine(Domain);
                    break;

                case AuthType.Certificate:
                    sCredentialInfo.AppendLine("Using certificate");
                    if (_certificate != null)
                    {
                        sCredentialInfo.Append("Subject: ");
                        sCredentialInfo.AppendLine(_certificate.Subject);
                    }
                    else
                        sCredentialInfo.AppendLine("NO CERTIFICATE SPECIFIED");
                    break;

                case AuthType.OAuth:
                    sCredentialInfo.AppendLine("Using OAuth");
                    if (_lastOAuthResult != null)
                        sCredentialInfo.AppendLine($"Current access token expiry: {_lastOAuthResult.ExpiresOn}");
                    break;

            }

            Logger.Log(sCredentialInfo.ToString(), "Request Credentials");
        }

        public bool ApplyCredentialsToHttpWebRequest(HttpWebRequest Request)
        {
            if (!HaveValidCredentials())
                return false;

            switch (_authType)
            {
                case AuthType.Default:
                    Request.UseDefaultCredentials = true;
                    return true;

                case AuthType.Basic:
                    Request.Credentials = new NetworkCredential(Username, _password);
                    return true;

                case AuthType.Certificate:
                    Request.ClientCertificates = new X509CertificateCollection();
                    Request.ClientCertificates.Add(_certificate);
                    return true;

                case AuthType.OAuth:
                    Request.Headers["Authorization"] = $"Bearer {_lastOAuthResult.AccessToken}";
                    return true;
            }

            return false;
        }

        public bool ApplyCredentialsToExchangeService(EWS.ExchangeService Exchange)
        {
            if (!HaveValidCredentials())
                return false;

            switch (_authType)
            {
                case AuthType.Default:
                    Exchange.UseDefaultCredentials = true;
                    return true;

                case AuthType.Basic:
                    Exchange.Credentials = new NetworkCredential(Username, _password);
                    return true;

                case AuthType.Certificate:
                    //Request.ClientCertificates = new X509CertificateCollection();
                    //Request.ClientCertificates.Add(_certificate);
                    return false;

                case AuthType.OAuth:
                    Exchange.Credentials = new EWS.OAuthCredentials(_lastOAuthResult.AccessToken);
                    return true;
            }

            return false;
        }

        public bool ApplyCredentialsToAutodiscoverService(Autodiscover.AutodiscoverService Autodiscover)
        {
            if (!HaveValidCredentials())
                return false;

            switch (_authType)
            {
                case AuthType.Default:
                    Autodiscover.UseDefaultCredentials = true;
                    return true;

                case AuthType.Basic:
                    Autodiscover.Credentials = new NetworkCredential(Username, _password);
                    return true;

                case AuthType.Certificate:
                    //Request.ClientCertificates = new X509CertificateCollection();
                    //Request.ClientCertificates.Add(_certificate);
                    return false;

                case AuthType.OAuth:
                    Autodiscover.Credentials = new EWS.OAuthCredentials(_lastOAuthResult.AccessToken);
                    return true;
            }

            return false;
        }
    }
}
