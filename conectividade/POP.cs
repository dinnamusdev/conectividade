using System;
using System.Net.Mail;
using OpenPop.Pop3;
using OpenPop.Mime;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace MailDinnamuS
{
    /*
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("EmailPOP")]
    [ComVisible(true)]*/
    public class Pop  
    {



        public bool MensagensExcluidas = false;
        private Pop3Client pop3 = new Pop3Client();
        private String MSG = "";
        public bool EstaConectado() {
            bool bRet = false;
            try
            {
                bRet=pop3.Connected; 
                if (bRet)
                {
                    if (!pop3.LastServerResponse.Contains("OK")) {
                        bRet = false;
                    }
                }
            }
            catch (Exception ex)
            {

                bRet = false;
                Dispose(); 
                MSG = ex.Message; 
            } 
            
           
            return bRet;
        }
        public bool Conectado =false;
        /*
        public EmailPOP(String ServidorPOP, int Porta , bool SSL , String Usuario, String Senha){
        
            Conectado = Conectar(ServidorPOP,Porta,SSL,Usuario,Senha);
	                       
        }*/
        
        public String MensagemErro() {
            return (MSG != "" ? MSG : "");
        }
        public String UltimaMensagemServidor() {
            String Msg = "";
            try
            {
                Msg = (pop3.LastServerResponse == null ? "" : pop3.LastServerResponse);
            }
            catch (Exception ex)
            {

                MSG = ex.Message; ;
                Dispose(); 
            }
           

            return Msg;
        }
        public bool Disconectar()
        {
            try
            {

                pop3.Disconnect();
                Conectado = false;
                return true;
            }
            catch (Exception ex)
            {
                
                 MSG = ex.Message;
                 return false;
            }
            
        }

        public bool ResetarConexao()
        {
            try
            {

                pop3.Reset();

                return true;
            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                Dispose(); 
                return false;
            }

        }

        public bool NOOPConexao()
        {
            try
            {

                pop3.NoOperation();

                return true;
            }
            catch (Exception ex)
            {

                MSG = ex.Message;

                return false;
            }

        }
        public void Dispose() {
            try
            {
                
                pop3.Dispose();
            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                try
                {
                    pop3 = new Pop3Client();
                }
                catch (Exception ex2)
                {
                    
                     MSG = ex2.Message;
                }
            }
           
        }

        public bool Conectar(String ServidorPOP, int Porta , bool SSL , String Usuario, String Senha) {
            try
            {
                Conectado = false;                
                
                pop3.Connect(ServidorPOP, Porta, SSL);
                
                pop3.Authenticate(Usuario, Senha);
                
                
                Conectado = true;
                return true;
                
            }
            catch (Exception ex)
            {
                MSG = ex.Message;
                Dispose(); 
                return false;
            }
        
        }
        public bool MensagemDeletar(int nID)
        {
            try
            {

                pop3.DeleteMessage(nID);

                return true;
            }
            catch (Exception ex)
            {
                
                  MSG = ex.Message;
                  Dispose(); 
                  return false;
            }
        
        }
        public bool MensagemDeletarTodas()
        {
            try
            {

                pop3.DeleteAllMessages();

                return true;
            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                Dispose(); 
                return false;
            }

        }
        public String MensagemInformacao(int nIDMensagem)
        {
            return MensagemInformacao(nIDMensagem, "|");
        }
        public String MensagemCorpo(int nIDMensagem) {
            String cRetorno = "";
            try
            {

                Message m = pop3.GetMessage(nIDMensagem);

                cRetorno= m.ToMailMessage().Body;

            }
            catch (Exception ex)
            {
                
                MSG = ex.Message;
                Dispose(); 
                
            }
            return cRetorno;
        }

        public String MensagemInformacao(int nIDMensagem, String Separador)
        {
            String cRetorno = "";
            try
            {

                Message m = pop3.GetMessage(nIDMensagem);

                String cMSG = m.ToMailMessage().Body;

                cRetorno = m.Headers.From.Address.ToUpper() + Separador +
                          m.Headers.To[0].Address.ToUpper() + Separador +
                          m.Headers.Subject.ToString();




            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                Dispose(); 

            }
            return cRetorno;
        }

        public String MensagemAnexos(int nIDMensagem)
        {
            
            String Anexos = "";

            try
            {

                Message m = pop3.GetMessage(nIDMensagem);
                
                for (int i = 0; i < m.MessagePart.MessageParts.Count; i++)
			    {
                    MessagePart mp = m.MessagePart.MessageParts[i];
                   
                    if (mp.IsAttachment)
                    {
                        Anexos = Anexos + (Anexos.Length>0 ? "|" : "") +(mp.FileName);
                    }
			    }
                
               
            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                Dispose(); 

            }
            return Anexos;
        }
        public bool MensagemAnexosSalvar(int nIDMensagem,String cAnexo, String cLocalGravacao )
        {
            
            List<String> ls = new List<string>();
            
            try
            {

                Message m = pop3.GetMessage(nIDMensagem);

                for (int i = 0; i < m.MessagePart.MessageParts.Count; i++)
                {
                    MessagePart mp = m.MessagePart.MessageParts[i];

                    if (mp.IsAttachment)
                    {
                        if (mp.FileName.ToUpper()==cAnexo.ToUpper())
                        {
                            mp.Save(new System.IO.FileInfo(cLocalGravacao + "\\" + cAnexo));
                            break;
                        }
                        
                    }
                }


            }
            catch (Exception ex)
            {

                MSG = ex.Message;
                Dispose(); 
                return false;

            }
            return true;
        }

        public int MensagensTotal()
        {
            try
            {
                
                return pop3.GetMessageCount();
            }
            catch (Exception ex)
            {
                
                MSG = ex.Message;
                Dispose(); 
                return -1;
            }
        
        }
        
    }
}
