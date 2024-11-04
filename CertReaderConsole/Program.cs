using System.Security.Cryptography.X509Certificates;

namespace CertReaderConsole
{
    public class Program
    {
        public static X509Certificate2 GetLatestCertificate()
        {
            using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);

                // Fetch all certificates and order them by the issuance date (NotBefore)
                var latestCertificate = store.Certificates
                    .OfType<X509Certificate2>()
                    .OrderByDescending(cert => cert.NotBefore)
                    .FirstOrDefault();

                store.Close();

                return latestCertificate;
            }
        }

        static void Main(string[] args)
        {
            var latestCertificate = GetLatestCertificate();
            if (latestCertificate != null)
            {
                Console.WriteLine("Latest Certificate:");
                Console.WriteLine("Subject: " + latestCertificate.Subject);
                Console.WriteLine("Friendly Name: " + latestCertificate.FriendlyName);
                Console.WriteLine("Issuer: " + latestCertificate.Issuer);
                Console.WriteLine("Issued On: " + latestCertificate.NotBefore);
                Console.WriteLine("Expires On: " + latestCertificate.NotAfter);
            }
            else
            {
                Console.WriteLine("No certificate found!");
            }
        }
    }
}
