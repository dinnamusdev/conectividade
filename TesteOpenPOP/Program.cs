using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MailDinnamuS;
 


namespace TesteOpenPOP
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            EmailPOP _EmailPOP = new EmailPOP();
            
            _EmailPOP.Conectar("pop.mail.yahoo.com.br", 995, true, "fernando_zoe@yahoo.com.br", "dinnamus180");

            int nTotalMsg = _EmailPOP.MensagensTotal();
            String cAnexos="";
            for (int i = 1; i <= nTotalMsg; i++)
            {
                
                System.Console.Out.Write(_EmailPOP.MensagemCorpo(i) +"\n");

                cAnexos = _EmailPOP.MensagemAnexos(i);

                if (cAnexos.Length>0)
                {
                    
                    System.Console.Out.Write(cAnexos);
                    String[] cLista = cAnexos.Split('|');
                    for (int j = 0; j < cLista.Length; j++)
                    {
                        _EmailPOP.MensagemAnexosSalvar(i, cLista[j], "c:\\");                            
                    }

                }

            }
             * */



            Smtp _smtp = new Smtp();

            _smtp.Configurar("smtp.mail.yahoo.com.br", 25, true, "fernando_zoe@yahoo.com.br", "dinnamus180", false);
            
            System.Console.ReadKey();
        }
    }
}
