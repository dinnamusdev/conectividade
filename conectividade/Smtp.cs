using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.Net.Sockets;

namespace MailDinnamuS
{
    public class Smtp
    {
        private String SMTP_Serv="";
        private int SMTP_Port=0;
        public bool Conectado = false;
        private SmtpClient smtp = null;
        public String MSG = "";
        public bool EnviarEmail(MailMessage m) {

            try
            {
                smtp.Send(m);
                    
                return true;
            }
            catch (Exception ex)
            {


                MSG = ex.Message;
                return false;
            }
        }
        public void Fechar() { 
                
            smtp.Dispose();
            Conectado = false;
        
        }


        public  bool TestConnection(string smtpServerAddress, int port)
        {

            
            try
            {
                IPHostEntry hostEntry = Dns.GetHostEntry(smtpServerAddress);
                IPEndPoint endPoint = new IPEndPoint(hostEntry.AddressList[0], port);
                using (Socket tcpSocket = new Socket(endPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp))
                {

                    //tcpSocket.
                    //try to connect and test the rsponse for code 220 = success
                    tcpSocket.Connect(endPoint);
                    if (!CheckResponse(tcpSocket, 220))
                    {
                        return false;
                    }

                    // send HELO and test the response for code 250 = proper response
                    SendData(tcpSocket, string.Format("HELO {0}\r\n", Dns.GetHostName()));
                    if (!CheckResponse(tcpSocket, 250))
                    {
                        return false;
                    }

                    // if we got here it's that we can connect to the smtp server
                    return true;
                }
            }
            catch (Exception ex)
            {
                 MSG = ex.Message;
                return false;
            }
           
        }

        private  void SendData(Socket socket, string data)
        {
            try
            {
            byte[] dataArray = Encoding.ASCII.GetBytes(data);
            socket.Send(dataArray, 0, dataArray.Length, SocketFlags.None);
            }
            catch (Exception ex)
            {
                
                MSG = ex.Message;
            }
         
        }

        private  bool CheckResponse(Socket socket, int expectedCode)
        {
            while (socket.Available == 0)
            {
                System.Threading.Thread.Sleep(100);
            }
            byte[] responseArray = new byte[1024];
            socket.Receive(responseArray, 0, socket.Available, SocketFlags.None);
            string responseData = Encoding.ASCII.GetString(responseArray);
            int responseCode = Convert.ToInt32(responseData.Substring(0, 3));
            if (responseCode == expectedCode)
            {
                return true;
            }
            return false;
        }
        public bool TestarConexao() {
            if (this.SMTP_Serv.Trim().Length > 0)
            {
                return TestConnection(this.SMTP_Serv, this.SMTP_Port);
            }
            else {
                return false;
            }
        }
        public bool Configurar(String SMTP_Serv, int SMTP_Port, bool SMTP_Autenticar, String SMTP_UserName, String SMTP_Password, bool SMTP_SSL) {
            try
            {
                Conectado = TestConnection(SMTP_Serv, SMTP_Port);
                if (Conectado)
                {
                    this.SMTP_Serv = SMTP_Serv;
                    this.SMTP_Port=SMTP_Port;
                    smtp = new SmtpClient(SMTP_Serv, SMTP_Port);

                    
                    
                    smtp.Credentials = new NetworkCredential(SMTP_UserName, SMTP_Password);
                    //smtp.
                    smtp.EnableSsl = SMTP_SSL;
                    //smtp.DeliveryMethod = SmtpDeliveryMethod.
                    //smtp.Timeout = 100000;

                    smtp.UseDefaultCredentials = false;
                    
                    /* 
                    MailMessage mail = new MailMessage(SMTP_UserName,  "a@b.xx","Integrador 1.0 DinnamuS", "Teste de Configuração");
                                
                    smtp.Send(mail);*/
                }

                return Conectado;
            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                return false;
            }
        
        }
    }
}
