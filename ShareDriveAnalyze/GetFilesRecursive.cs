using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace alphaDrive
{
    class GetFilesRecursive
    {

        public static object[] GetFiles(string b)
        {

            object[] returnResults = new object[2];

            // 1.
            // Store results in the file results list.
            List<string> result = new List<string>();

            List<string> dirs = new List<string>();

            // 2.
            // Store a stack of our directories.
            Stack<string> stack = new Stack<string>();

            // 3.
            // Add initial directory.
            stack.Push(b);

            // 4.
            // Continue while there are directories to process
            while (stack.Count > 0)
            {
                // A.
                // Get top directory
                string dir = stack.Pop();

                try
                {
                    // B
                    // Add all files at this directory to the result List.
                    result.AddRange(Directory.GetFiles(dir, "*.*"));

                    // C
                    // Add all directories at this directory.
                    foreach (string dn in Directory.GetDirectories(dir))
                    {
                        dirs.AddRange(Directory.GetDirectories(dir));
                        stack.Push(dn);
                    }
                }
                catch
                {
                    // D
                    // Could not open the directory
                }
            }

            returnResults[0] = result;
            returnResults[1] = dirs;

            return returnResults;
        }
    }
}
