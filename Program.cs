using System;
using System.Security.Cryptography;
using System.Windows.Forms;
using ExemploAssinadorXML.Services;

namespace ExemploAssinadorXML
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Registrar algoritmo de assinatura
            CryptoConfig.AddAlgorithm(typeof(RSAPKCS1SHA256SignatureDescription), 
                @"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256");
            
            Application.Run(new MainForm());
        }
    }
}
