using SGID.Data.ViewModel;
using System.Text;

namespace SGID.Data.Models
{
    public static class Logger
    {
        public static void Log(Exception exception, ApplicationDbContext dblog, string page,string user)
        {
            // Create an instance of StringBuilder. This class is in System.Text namespace
            StringBuilder sbExceptionMessage = new ();
            var mensagem = "";

            sbExceptionMessage.Append("Exception Type" + Environment.NewLine);
            // Get the exception type
            sbExceptionMessage.Append(exception.GetType().Name);
            // Environment.NewLine writes new line character - \n
            sbExceptionMessage.Append(Environment.NewLine + Environment.NewLine);
            sbExceptionMessage.Append("Message" + Environment.NewLine);
            // Get the exception message
            sbExceptionMessage.Append(exception.Message + Environment.NewLine + Environment.NewLine);
            sbExceptionMessage.Append("Stack Trace" + Environment.NewLine);
            // Get the exception stack trace
            sbExceptionMessage.Append(exception.StackTrace + Environment.NewLine + Environment.NewLine);

            // Retrieve inner exception if any
            Exception innerException = exception.InnerException;
            // If inner exception exists
            while (innerException != null)
            {
                sbExceptionMessage.Append("Exception Type" + Environment.NewLine);
                sbExceptionMessage.Append(innerException.GetType().Name);
                sbExceptionMessage.Append(Environment.NewLine + Environment.NewLine);
                sbExceptionMessage.Append("Message" + Environment.NewLine);
                sbExceptionMessage.Append(innerException.Message + Environment.NewLine + Environment.NewLine);

                sbExceptionMessage.Append("Stack Trace" + Environment.NewLine);
                sbExceptionMessage.Append(innerException.StackTrace + Environment.NewLine + Environment.NewLine);

                // Retrieve inner exception if any
                innerException = innerException.InnerException;
            }

            try
            {

                mensagem = sbExceptionMessage.ToString();
                Log erro = new()
                {
                    App = "SGID",
                    Page = page,
                    DataError = DateTime.Now,
                    Description = mensagem,
                    User = user
                };
                dblog.Logs.Add(erro);
                dblog.SaveChanges();
            }
            catch
            {
  
            }
        }
    }
}
