using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RTBox59

{
    //KOAD: Interface necessary to get IntelliSense working in AWCHW
    [Guid("2BAEB75A-EF79-4C79-A5E8-6F90371C1192")]
    [ComVisible(true)]                                      //KOAD: It's better to set it also here even if ComVisibility is set in AssemblyInfo
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]       //KOAD: VB6 probably uses Dispatch interface, dual is set just in case
    public interface IRTBox59
    {
        int GetFileLength();
        string ReadLine(int line);
        bool ReplaceValue(string newValue);
        void SaveFile(string path);
        bool LoadRTF(string path);
        void DeleteLine( int line);
        void DeleteBlock(int firstLine, int linesAmount);
        void Paginate( int line);
        int GetXPosition();
        int GetYPosition();
        int GetEOFChar();
        string GetSelectedText();
        string GetSelectedRTF();
        bool SelectText(int x,int len);
        bool IsValid(string lineToBeSearched, string searchedValue);
        int CountPlaceHolders(string line, string placeHolder);
    }
    [Guid("282B13B7-BF66-429E-8CFB-46A37B0B1069")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class RTBox59 : IRTBox59
    {

        RichTextBox RTxtBox; //RichTextBox Control to manipulate template
        int yPos; //Y position of cursor
        int xPos; // X position of cursor
        
        //KOAD: Default constructor - must be present to make .DLL interoperable with VB6
        public RTBox59()
        {
        }

        public bool LoadRTF(string path)
        {
            try
            {
                RTxtBox = new RichTextBox();
                RTxtBox.LoadFile(path);
                RTxtBox.WordWrap = false;
                yPos = 0;
                xPos = RTxtBox.GetFirstCharIndexFromLine(yPos);
                
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public int GetFileLength()
        {
            return RTxtBox.Lines.Length;
        }

        public int GetYPosition()
        {
            return  yPos;
        }

        public int GetXPosition()
        {
            return xPos;
        }

        public int GetEOFChar()
        {
            return 164;
        }

        //readline method
        public string ReadLine(int line)
        {
            int s1 = RTxtBox.GetFirstCharIndexFromLine(line);
            //int s2 = line < RTxtBox.Lines.Length - 1 ?
            //          RTxtBox.GetFirstCharIndexFromLine(line + 1) - 1 :
            //          RTxtBox.Text.Length;

            //RTxtBox.Select(s1, s2 - s1);
            yPos = line;
            xPos = s1;
            //return RTxtBox.SelectedText;
            return RTxtBox.Lines[line];

        }
        
        public bool SelectText(int x, int len)
        {

            try
            {
                RTxtBox.Select(x, len);
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }

        public string GetSelectedText()
        {
            return RTxtBox.SelectedText;
        }

        public string GetSelectedRTF()
        {
            return RTxtBox.SelectedRtf;
        }
        //replace value method
        public bool ReplaceValue( string newValue)
        {

            //Regex regex = new Regex(placeHolder);
            //string originalLine = ReadLine(lineNumber);

            //    string newLine = originalLine;
            //    newLine = regex.Replace(originalLine, newValue);

            //    changeLine(RTxtBox, lineNumber, newLine);
            //RTFlines = RTxtBox.Lines;
            
            
            RTxtBox.SelectedText = newValue;
            //xPos = xPos + newValue.Length;
            return true;
        }

        public void SaveFile(string path)
        {

            //RTxtBox.Lines = RTFlines;
            RTxtBox.SaveFile(path);
        }

        
        //delete line method
        public void DeleteLine( int line)
        {
            //int s1 = RTxtBox.GetFirstCharIndexFromLine(line);
            //int s2 = line < RTxtBox.Lines.Length - 1 ?
            //          RTxtBox.GetFirstCharIndexFromLine(line + 1) - 1 :
            //          RTxtBox.Text.Length;

            //RTxtBox.Select(s1, s2 - s1);
            //RTxtBox.SelectedText = String.Empty;

            int s1 = RTxtBox.GetFirstCharIndexFromLine(line);   //Grab first char in current line
            //int s2 = line < RTxtBox.Lines.Length - 1 ?          //Grab last char in current line
            //          RTxtBox.GetFirstCharIndexFromLine(line + 1) - 1 :
            //          RTxtBox.Text.Length;
            //if (line == 0)
            //{
            //    s1 = RTxtBox.GetFirstCharIndexFromLine(line);
            //}
            //else
            //{
            //    s1 = RTxtBox.GetFirstCharIndexFromLine(line - 1) + RTxtBox.Lines[line - 1].Length;
            //}
            //int s2 = line < RTxtBox.Lines.Length - 1 ?
            //          RTxtBox.GetFirstCharIndexFromLine(line + 1) - 1 :
            //          RTxtBox.Text.Length;



            //RTxtBox.Select(s1, s2 - s1);
            //Regex.Replace(RTxtBox.SelectedRtf, @"\\par", "");
            //RTxtBox.SelectedRtf = "";
            //RTxtBox.SelectedText = String.Empty;
            RTxtBox.Select(s1, RTxtBox.Lines[line].Length + 1 ); //KOAD: Grab whole line plus 2 chars overhead to check for new line char
            if (Regex.IsMatch(RTxtBox.SelectedText, @"\n") && RTxtBox.Lines[line].Length > 0)     
            {
                RTxtBox.SelectedText = String.Empty;
            }
           
        }

        public void DeleteBlock(int firstLine, int linesAmount)
        {
            int lastLine = firstLine + linesAmount-1; // - 1 because it includes also firstLine
            int s1 = RTxtBox.GetFirstCharIndexFromLine(firstLine - 1) + RTxtBox.Lines[firstLine - 1].Length;
            int s2 = lastLine < RTxtBox.Lines.Length - 1 ?
                      RTxtBox.GetFirstCharIndexFromLine(lastLine + 1) - 1 :
                      RTxtBox.Text.Length;

            //RTxtBox.Select(s1, s2 - s1);
            //Regex.Replace(RTxtBox.SelectedRtf, @"\\par", "");
            //RTxtBox.SelectedRtf = "";
            //RTxtBox.SelectedText = String.Empty;
            for (int i = 0;i<linesAmount;i++)
            {
                DeleteLine(firstLine);
            }
        }

        //paginate method
        public void Paginate(int line)//TODO: remove RTB as parameter and think of way to mark pagebreaks sss
        {
            int s1 = RTxtBox.GetFirstCharIndexFromLine(line);
            int s2 = line < RTxtBox.Lines.Length - 1 ?
                      RTxtBox.GetFirstCharIndexFromLine(line + 1) - 1 :
                      RTxtBox.Text.Length;

            RTxtBox.Select(s1, s2 - s1);
            RTxtBox.SelectedRtf = @"{\rtf1 \par \page}";
        }




        //========================
        //KOAD: DEPRECATED METHODS
        //========================
        public int CountPlaceHolders(string line, string placeHolder)
        {
            Regex pattern = new Regex(placeHolder);
            return pattern.Matches(line).Count;

        }

        void changeLine(RichTextBox RTB, int line, string text) //KOAD: Probably replaced by ReplaceValue and can be deleted
        {
            int s1 = RTB.GetFirstCharIndexFromLine(line);
            int s2 = line < RTB.Lines.Length - 1 ?
                      RTB.GetFirstCharIndexFromLine(line + 1) - 1 :
                      RTB.Text.Length;

            RTB.Select(s1, s2 - s1);
            RTB.SelectedText = text;
        }

        public bool IsValid(string lineToBeSearched, string searchedValue) //KOAD: Probably not needed and can be deleted
        {
            Regex pattern = new Regex(searchedValue);

            return pattern.IsMatch(lineToBeSearched);
        }
    }
}
