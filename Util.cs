using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegratorSales
{
    class Util
    {


        public static string truncatetexto(string texto, int tamanho)
        {
            string textotruncado = texto;
            if (texto.Length > tamanho)
            {
                textotruncado = texto.Substring(0, tamanho);
            }

            return textotruncado;
        }

    }
}
