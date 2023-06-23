namespace SGID.Models.Email
{
    public static class EmailTemplate
    {
        public static string LerArquivoHtml<T>(string filePath, T model)
        {
            var sr = new StreamReader(filePath);

            var html = "";

            while (!sr.EndOfStream)
            {
                var conteudo = sr.ReadLine();

                foreach (var property in model.GetType().GetProperties())
                    if (conteudo.Contains($"[{property.Name}]"))
                        if (property.GetValue(model) != null)
                            conteudo = conteudo.Replace($"[{property.Name}]", property.GetValue(model).ToString());

                html += conteudo;
            }

            sr.Close();

            return html;
        }
    }
}
