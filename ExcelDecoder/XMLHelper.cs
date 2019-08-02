// --------------------------------------------------------------------------------------------------------------------
// Copyright (C) 2019 Alexandre Almeida
//
// This program is free software: you can redistribute it and/or modify
// it under the +terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see http://www.gnu.org/licenses/. 
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.IO;
using System.Windows.Controls;
using System.Xml.Linq;

namespace ExcelDecoder
{
    class XMLHelper
    {
        // Função responsavel por quebrar a proteção de folhas individuais
        public static void XMLManipulation()
        {
            // Definimos a diretoria onde as folhas do excel em formato .xml estão armazenadas de forma temporaria
            string workSheetsPath = System.IO.Path.GetTempPath() + "ExcelDecoder" + "/" + "xl" + "/" + "worksheets" + "/";
            

            // Criamos uma Array de Strings com todos os ficheiros .xml encontrados na diretoria
            string[] array2 = Directory.GetFiles(workSheetsPath, "*.xml");

            // Criamos um loop que vai processar de forma sequencial todos os arquivos armazenados na nossa Array
            foreach (string name in array2)
            {
                
                try
                {
                    // Carrega para a memoria os ficheiros .xml
                    var xdoc = XDocument.Load(name);

                    // Define o Namespace dos ficheiros
                    XNamespace xns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                    // Eliminamos o elemento sheetProtection assim quebrando a proteção da página
                    xdoc.Element(xns + "worksheet").Element(xns + "sheetProtection").RemoveAll();

                    // Substituimos a versão original dos ficheiros .xml por versões desbloqueadas
                    xdoc.Save(name);

                }
                
                // Ignorem este catch apenas serve para debug, se o ficheiro .xml não conter o elemento sheetProtection imprimimos uma mensagem na consola
                catch (Exception ex)
                {
                    Console.WriteLine("Erro XML");
                }

            }
           

        }
            
           
            }
           
        }
    

