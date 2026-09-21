using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APOffice
{
    class clsCtrlCodici
    {
        public string FindMod10Digit(string input)
        {
            if (string.IsNullOrEmpty(input)) return "0";

            try
            {
                int num1 = 3;
                int num2 = 0;
                for (int length = input.Length; length > 0; --length)
                {
                    char c = input[length - 1];
                    if (char.IsDigit(c))
                    {
                        num2 += (c - '0') * num1;
                        num1 = 4 - num1;
                    }
                    else
                    {
                        // Se contiene caratteri non numerici (es: codici speciali o lettere), restituisce l'ultimo carattere senza crashare
                        return input.Substring(input.Length - 1, 1);
                    }
                }
                int num3 = num2 % 10;
                if (num3 == 0)
                    return "0";
                else
                    return (10 - num3).ToString();
            }
            catch
            {
                if (input.Length > 0)
                    return input.Substring(input.Length - 1, 1);
                return "0";
            }
        }

        public bool ControllaPartitaIva(string PartitaIva)
        {
            int[] ListaPari = { 0, 2, 4, 6, 8, 1, 3, 5, 7, 9 };

            // normalizziamo la cifra
            if (PartitaIva.Length < 11)
                PartitaIva = PartitaIva.PadLeft(11, '0');

            // lunghezza errata non fa neanche il controllo
            if (PartitaIva.Length != 11)
                return false;

            int Somma = 0;

            for (int k = 0; k < 11; k++)
            {
                string s = PartitaIva.Substring(k, 1);

                // otteniamo contemporaneamente 
                // il valore, la posizione e testiamo se ci sono
                // caratteri non numerici

                int i = "0123456789".IndexOf(s);
                if (i == -1)
                    return false;

                int x = int.Parse(s);
                if (k % 2 == 1) // Pari perchè iniziamo da zero
                    x = ListaPari[i];

                Somma += x;
            }

            return ((Somma % 10 == 0) && (Somma != 0));
        }


        public bool VerificaCodiceFiscale(string Codice)
        {
            // Dato un Codice Fiscale verifica il codice di controllo
            // Input: il Codice Fiscale da verificare, 16 caratteri
            // Output: true/false
            // 2010.12.05
            Codice = Codice.ToUpper();
            if (Codice.Length != 16)
                return false; // errore
            else
            {
                if (Codice.Substring(15, 1) ==
                    CalcolaCodiceControllo(Codice.Substring(0, 15)))
                    return true;
                else
                    return false;
            }
        }
        private static string CalcolaCodiceControllo(string Codice)
        {
            // Calcola il codice di controllo del Codice Fiscale
            // Input: i primi 15 caratteri del Codice Fiscale
            // Output: il codice di controllo
            // 2010.12.05
            int Contatore = 0;
            Codice = Codice.ToUpper();
            if (Codice.Length != 15)
                return "0"; // zero: errore
            else
            {
                for (int i = 0; i < Codice.Length; i++)
                {
                    Contatore += ValoreDelCarattere(Codice.Substring(i, 1), i);
                }
                Contatore %= 26; // si considera il resto
                return Convert.ToChar(Contatore + 65).ToString();
            }
        }
        private static int ValoreDelCarattere(string Carattere, int Posizione)
        {
            int Valore = 0;
            switch (Carattere)
            {
                case "A":
                case "0":
                    if ((Posizione % 2) == 0)
                        Valore = 1;
                    else
                        Valore = 0;
                    break;
                case "B":
                case "1":
                    if ((Posizione % 2) == 0)
                        Valore = 0;
                    else
                        Valore = 1;
                    break;
                case "C":
                case "2":
                    if ((Posizione % 2) == 0)
                        Valore = 5;
                    else
                        Valore = 2;
                    break;
                case "D":
                case "3":
                    if ((Posizione % 2) == 0)
                        Valore = 7;
                    else
                        Valore = 3;
                    break;
                case "E":
                case "4":
                    if ((Posizione % 2) == 0)
                        Valore = 9;
                    else
                        Valore = 4;
                    break;
                case "F":
                case "5":
                    if ((Posizione % 2) == 0)
                        Valore = 13;
                    else
                        Valore = 5;
                    break;
                case "G":
                case "6":
                    if ((Posizione % 2) == 0)
                        Valore = 15;
                    else
                        Valore = 6;
                    break;
                case "H":
                case "7":
                    if ((Posizione % 2) == 0)
                        Valore = 17;
                    else
                        Valore = 7;
                    break;
                case "I":
                case "8":
                    if ((Posizione % 2) == 0)
                        Valore = 19;
                    else
                        Valore = 8;
                    break;
                case "J":
                case "9":
                    if ((Posizione % 2) == 0)
                        Valore = 21;
                    else
                        Valore = 9;
                    break;
                case "K":
                    if ((Posizione % 2) == 0)
                        Valore = 2;
                    else
                        Valore = 10;
                    break;
                case "L":
                    if ((Posizione % 2) == 0)
                        Valore = 4;
                    else
                        Valore = 11;
                    break;
                case "M":
                    if ((Posizione % 2) == 0)
                        Valore = 18;
                    else
                        Valore = 12;
                    break;
                case "N":
                    if ((Posizione % 2) == 0)
                        Valore = 20;
                    else
                        Valore = 13;
                    break;
                case "O":
                    if ((Posizione % 2) == 0)
                        Valore = 11;
                    else
                        Valore = 14;
                    break;
                case "P":
                    if ((Posizione % 2) == 0)
                        Valore = 3;
                    else
                        Valore = 15;
                    break;
                case "Q":
                    if ((Posizione % 2) == 0)
                        Valore = 6;
                    else
                        Valore = 16;
                    break;
                case "R":
                    if ((Posizione % 2) == 0)
                        Valore = 8;
                    else
                        Valore = 17;
                    break;
                case "S":
                    if ((Posizione % 2) == 0)
                        Valore = 12;
                    else
                        Valore = 18;
                    break;
                case "T":
                    if ((Posizione % 2) == 0)
                        Valore = 14;
                    else
                        Valore = 19;
                    break;
                case "U":
                    if ((Posizione % 2) == 0)
                        Valore = 16;
                    else
                        Valore = 20;
                    break;
                case "V":
                    if ((Posizione % 2) == 0)
                        Valore = 10;
                    else
                        Valore = 21;
                    break;
                case "W":
                    if ((Posizione % 2) == 0)
                        Valore = 22;
                    else
                        Valore = 22;
                    break;
                case "X":
                    if ((Posizione % 2) == 0)
                        Valore = 25;
                    else
                        Valore = 23;
                    break;
                case "Y":
                    if ((Posizione % 2) == 0)
                        Valore = 24;
                    else
                        Valore = 24;
                    break;
                case "Z":
                    if ((Posizione % 2) == 0)
                        Valore = 23;
                    else
                        Valore = 25;
                    break;
                default:
                    Valore = 0;
                    break;
            }
            return Valore;
        }
    }

}
