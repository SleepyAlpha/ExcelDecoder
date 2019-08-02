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

using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelDecoder
{
    class Patcher
    {
        // Array de String com os valores hexadecimais a procurar, neste caso os valores equivalentes há String DPB encontrados no ficheiro original
        public static readonly byte[] PatchFind = {  0x44, 0x50, 0x42 };

        // Array de String com os valores hexadecimais a substituir, neste caso os valores equivalentes há String DPx
        public static readonly byte[] PatchReplace = {  0x44, 0x50, 0x78 };

        // Lógica necessária
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool DetectPatch(byte[] sequence, int position)
        {
            if (position + PatchFind.Length > sequence.Length) return false;
            for (int p = 0; p < PatchFind.Length; p++)
            {
                if (PatchFind[p] != sequence[position + p]) return false;
            }
            return true;
        }

        // Função responsável por alterar os valores hexadecimais necessários para quebrar a proteção das macros
        public static void PatchFile(string originalFile, string patchedFile)
        {
            
            // Lê todos os Bytes do ficheiro a aplicar o patch.
            byte[] fileContent = File.ReadAllBytes(originalFile);

            // Deteta e aplica o patch.
            for (int p = 0; p < fileContent.Length; p++)
            {
                if (!DetectPatch(fileContent, p)) continue;

                for (int w = 0; w < PatchFind.Length; w++)
                {
                    fileContent[p + w] = PatchReplace[w];
                }
            }

            // Depois de processado substituimos o ficheiro original pela versão modificada.
            File.WriteAllBytes(patchedFile, fileContent);
        }

    }

}
