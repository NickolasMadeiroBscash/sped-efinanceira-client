using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class ConfiguracaoPersistenciaService
    {
        private const string CONFIG_FILE = "efinanceira_config.xml";

        public void SalvarConfiguracao(EfinanceiraConfig config, DadosAbertura dadosAbertura, DadosFechamento dadosFechamento)
        {
            try
            {
                var configCompleta = new ConfiguracaoCompleta
                {
                    Config = config,
                    DadosAbertura = dadosAbertura,
                    DadosFechamento = dadosFechamento,
                    DataSalvamento = DateTime.Now
                };

                string caminhoArquivo = ObterCaminhoArquivo();
                XmlSerializer serializer = new XmlSerializer(typeof(ConfiguracaoCompleta));
                
                using (FileStream stream = new FileStream(caminhoArquivo, FileMode.Create))
                {
                    serializer.Serialize(stream, configCompleta);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar configuração: {ex.Message}", ex);
            }
        }

        public ConfiguracaoCompleta CarregarConfiguracao()
        {
            try
            {
                string caminhoArquivo = ObterCaminhoArquivo();
                
                if (!File.Exists(caminhoArquivo))
                {
                    return null;
                }

                XmlSerializer serializer = new XmlSerializer(typeof(ConfiguracaoCompleta));
                using (FileStream stream = new FileStream(caminhoArquivo, FileMode.Open))
                {
                    return (ConfiguracaoCompleta)serializer.Deserialize(stream);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao carregar configuração: {ex.Message}", ex);
            }
        }

        private string ObterCaminhoArquivo()
        {
            string diretorio = Path.Combine(Application.StartupPath, "config");
            if (!Directory.Exists(diretorio))
            {
                Directory.CreateDirectory(diretorio);
            }
            return Path.Combine(diretorio, CONFIG_FILE);
        }
    }

    [Serializable]
    [XmlRoot("ConfiguracaoCompleta")]
    public class ConfiguracaoCompleta
    {
        public EfinanceiraConfig Config { get; set; }
        public DadosAbertura DadosAbertura { get; set; }
        public DadosFechamento DadosFechamento { get; set; }
        public DateTime DataSalvamento { get; set; }
    }
}
