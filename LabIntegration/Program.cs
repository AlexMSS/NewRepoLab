using LabComm;
using System;

namespace LabIntegration
{
    class Program
    {
        static void Main(string[] args)
        {
            Processing.Initiation();
            Console.WriteLine("Тест 1");
            //if (!Files.GoFurther()) Console.WriteLine("Проблема с лицензией!");

            //if (Processing.Ini.LoadDictionaries) Processing.DLoadDictionaries.Invoke();
            Invitro.LoadDictionaries();
            if (Processing.Upload && Processing.Ini.DriverCode != "INV") Processing.OrderTimer.Enabled = true;
            if (Processing.Upload && Processing.Ini.DriverCode == "INV") 
            { 
                Processing.bcChangeTimer.Enabled = true; Console.WriteLine("Запущена проверка заказов.");
                //Processing.delChangeTimer.Enabled = true; Console.WriteLine("Запущена проверка на удаление.");
            }
            Processing.ResultTimer.Enabled = true;
 
            Processing.WaitQ();
        }
    }
}
