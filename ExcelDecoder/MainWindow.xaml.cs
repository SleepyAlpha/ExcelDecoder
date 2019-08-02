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
using System.IO.Compression;
using System.Windows;


namespace ExcelDecoder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        // Definimos as Strings necessárias
        String originalFilePath;
        String xlPath = "xl";      

        // Detetamos e criamos uma String com a diretoria da pasta temporaria do Windows e dentro dessa diretoria criamos uma pasta temporaria com o nome ExcelDecoder
        String tempPath = System.IO.Path.GetTempPath() + "ExcelDecoder";

        public MainWindow()
        {
            InitializeComponent();

            // Alteramos a propriedade das caixas de texto para apenas leitura
            
            textBox1.IsReadOnly = true;
        }

        
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            // Cria a janela de dialogo
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();


            // Filtro para apenas mostrar ficheiros compativeis com o patcher
            dlg.DefaultExt = ".xlsm";
            dlg.Filter = "Ficheiros Excel (*.xlsm)|*.xlsm";


            // Mostra a janela 
            Nullable<bool> result = dlg.ShowDialog();


            // Obtem o ficheiro selecionado
            if (result == true)
            {            
                originalFilePath = dlg.FileName;               
            }
            textBox.Text += "Ficheiro " + originalFilePath + " Selecionado" + Environment.NewLine;


            // Elimina a pasta de trabalho se ela já existir, caso o programa já tiver sido executado com algum erro
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {

            // Como o excel usa ficheiros no formato zip, podemos simplesmente extrair o seu conteudo com a função ZipFile para a pasta de trabalho
            textBox.Text += "A Descompactar" + originalFilePath + Environment.NewLine;
            ZipFile.ExtractToDirectory(originalFilePath, tempPath);
            

            // Define a diretoria completa para encontrar-mos o ficheiro binario a modificar
            String fullXlPath = tempPath + "/" + xlPath + "/";

            // Usando a função PatchFile modificamos o ficheiro vbaProject.bin para desativar a proteção das macros
            textBox.Text += "Patch Hexadecimal Aplicado" + Environment.NewLine;
            Patcher.PatchFile(fullXlPath  + "vbaProject.bin", fullXlPath + "vbaProject.bin");


            // Usando a função XMLManipulation removemos a proteção de paginas individuais
            textBox.Text += "Folhas Desprotegidas" + Environment.NewLine;
            XMLHelper.XMLManipulation();


            // Usando outra vez a função ZipFile criamos um novo ficheiro compativel proveniente da pasta de trabalho
            textBox.Text += "A Recriar Arquivo" + Environment.NewLine;
            ZipFile.CreateFromDirectory(tempPath, originalFilePath + "Unlocked.xlsm");
            

            // Para não gastar espaço desnecessariamente ao utilizador apagamos a pasta de trabalho no ultimo momento de execução do programa
            Directory.Delete(tempPath, true);
            textBox.Text += "Ficheiro Descodificado" + Environment.NewLine;

        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // Lógica que gere o que acontece quando o utilizador interage com a TextBox, como o utilizador não a vai utilizar podemos deixar esta função em branco
        }

        private void TextBox1_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // Lógica que gere o que acontece quando o utilizador interage com a TextBox1, como o utilizador não a vai utilizar podemos deixar esta função em branco
        }
    }
}
