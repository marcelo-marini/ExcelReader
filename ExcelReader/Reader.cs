using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelReader {

    public class Reader {

        public DataTable ObterAniversariantes() {

            DateTime data = DateTime.Now.Date; // Aqui estou configurando sempre a data atual, porém ela pode vir como input de uma pesquisa
            DataTable resultado = LerArquivoExcel(); // Aqui pego o arquivo inteiro

            DataTable aniversariantes = resultado.Clone();

            foreach (DataRow item in resultado.Rows) {

                DateTime.TryParse(item["Data"].ToString(), out DateTime dataArquivo);

                // Aqui verifico se a data do arquivo, é igual a data que foi "escolhida" como input
                // Caso for, adiciono um novo registro na tabela de retorno

                if (dataArquivo.ToString("dd/MM") == data.ToString("dd/MM")) {

                    // Os nomes das colunas vão variar de acordo com o seu arquivo
                    var row = aniversariantes.NewRow();
                    row["Nome"] = item["Nome"];
                    row["Data"] = item["Data"];
                    aniversariantes.Rows.Add(row);
                }
            }

            return aniversariantes;
        }

        private DataTable LerArquivoExcel() {

            DataSet result = null;

            // diretório do arquivo
            string path = "C:\\Users\\Marcelo\\Desktop\\Teste.xlsx";


            var conf = new ExcelDataSetConfiguration {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration {
                    UseHeaderRow = true // Indica que a primeira têm os nomes das colunas
                }
            };

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //open file and returns as Stream
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read)) {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream)) {

                    result = reader.AsDataSet(conf); // Adiciono o arquivo num dataset
                }
            }

            // Estou retornando apenas a primeira tabela
            // OBS: se o arquivo excel tiver N planilhas, o DataSet terá N tabelas também
            return result.Tables[0];
        }
    }
}
