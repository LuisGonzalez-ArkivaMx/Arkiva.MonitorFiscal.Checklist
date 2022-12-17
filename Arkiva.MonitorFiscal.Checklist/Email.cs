using MFiles.VAF.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Arkiva.MonitorFiscal.Checklist
{
    public class Email
    {
        public bool Enviar(AlternateView mensaje, List<string> emails, string emailService, string hostService, int portService, string usernameService, string passwordService)
        {
            bool bResult = false;

            MailMessage oEmail = new MailMessage();
            string sDe = string.Format("{0} <{1}>", "Administrador M-Files", emailService);
            oEmail.From = new MailAddress(sDe, sDe.Trim().Substring(0, sDe.Trim().IndexOf("<")));

            foreach (var email in emails)
            {
                oEmail.To.Add(new MailAddress(email)); // Se agregan los emails destinatarios
            }

            oEmail.Subject = "Checklist Documental";
            oEmail.AlternateViews.Add(mensaje);
            oEmail.IsBodyHtml = true;
            oEmail.Priority = MailPriority.Normal;

            // Configuracion Smtp
            SmtpClient smtp = new SmtpClient();
            smtp.Host = hostService; // smtp.live.com // smtp.gmail.com
            smtp.Port = portService;
            smtp.EnableSsl = true;    
            
            smtp.Credentials = new NetworkCredential(usernameService, passwordService);
            
            try
            {
                // Enviar Email
                smtp.Send(oEmail);
                oEmail.Dispose();

                bResult = true;
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error en el envio del correo electronico. " + ex);
                oEmail.Dispose();
            }

            return bResult;
        }
    }
}
