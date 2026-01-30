using System;

namespace ExemploAssinadorXML.Services
{
    public static class EfinanceiraPeriodoUtil
    {
        /// <summary>
        /// Calcula as datas de início e fim do período semestral baseado no formato YYYYMM
        /// Onde MM pode ser:
        /// - 01 = Primeiro semestre (Janeiro a Junho)
        /// - 02 = Segundo semestre (Julho a Dezembro)
        /// - 06 = Primeiro semestre (Janeiro a Junho) - formato legado
        /// - 12 = Segundo semestre (Julho a Dezembro) - formato legado
        /// </summary>
        /// <param name="periodo">Período no formato YYYYMM (ex: 202301 para Jan-Jun/2023, 202302 para Jul-Dez/2023)</param>
        /// <returns>Tupla com (DataInicio, DataFim) no formato AAAA-MM-DD</returns>
        public static (string DataInicio, string DataFim) CalcularPeriodoSemestral(string periodo)
        {
            if (string.IsNullOrWhiteSpace(periodo) || periodo.Length != 6)
            {
                throw new ArgumentException("Período deve estar no formato YYYYMM (ex: 202301 para Jan-Jun ou 202302 para Jul-Dez)");
            }

            int ano = int.Parse(periodo.Substring(0, 4));
            int mes = int.Parse(periodo.Substring(4, 2));

            string dataInicio;
            string dataFim;

            // Novo formato: 01 = Primeiro semestre (Jan-Jun), 02 = Segundo semestre (Jul-Dez)
            if (mes == 1 || mes == 6)
            {
                // Primeiro semestre: Janeiro a Junho
                dataInicio = $"{ano}-01-01";
                dataFim = $"{ano}-06-30";
            }
            else if (mes == 2 || mes == 12)
            {
                // Segundo semestre: Julho a Dezembro
                dataInicio = $"{ano}-07-01";
                dataFim = $"{ano}-12-31";
            }
            else
            {
                throw new ArgumentException($"Mês inválido no período. Deve ser 01 ou 06 (Jan-Jun) ou 02 ou 12 (Jul-Dez). Recebido: {mes:00}");
            }

            return (dataInicio, dataFim);
        }

        /// <summary>
        /// Calcula o período semestral baseado na data atual
        /// Fevereiro processa Jul-Dez do ano anterior (período 02)
        /// Agosto processa Jan-Jun do ano atual (período 01)
        /// </summary>
        public static string CalcularPeriodoAtual()
        {
            DateTime hoje = DateTime.Now;
            int mesAtual = hoje.Month;
            int anoAtual = hoje.Year;

            if (mesAtual == 2)
            {
                // Fevereiro: processa Jul-Dez do ano anterior (segundo semestre)
                int anoAnterior = anoAtual - 1;
                return $"{anoAnterior}02";
            }
            else if (mesAtual == 8)
            {
                // Agosto: processa Jan-Jun do ano atual (primeiro semestre)
                return $"{anoAtual}01";
            }
            else
            {
                // Para outros meses, retorna o período do mês anterior
                DateTime mesAnterior = hoje.AddMonths(-1);
                int mes = mesAnterior.Month;
                int ano = mesAnterior.Year;

                if (mes <= 6)
                {
                    return $"{ano}01"; // Primeiro semestre
                }
                else
                {
                    return $"{ano}02"; // Segundo semestre
                }
            }
        }

        /// <summary>
        /// Valida se o período está no formato correto
        /// Aceita: 01 ou 06 (primeiro semestre) e 02 ou 12 (segundo semestre)
        /// </summary>
        public static bool ValidarPeriodo(string periodo)
        {
            if (string.IsNullOrWhiteSpace(periodo) || periodo.Length != 6)
            {
                return false;
            }

            if (!int.TryParse(periodo, out int periodoInt))
            {
                return false;
            }

            int mes = periodoInt % 100;
            // Aceita: 01 ou 06 (Jan-Jun) e 02 ou 12 (Jul-Dez)
            return (mes == 1 || mes == 6) || (mes == 2 || mes == 12);
        }
    }
}
