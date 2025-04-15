using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ScintillaNET;

namespace ExcelEssentials.Scripts
{
    public static class UtilsScintilla
    {
        public static void Comment(Scintilla editor, string commentSymbol = "--")
        {
            int startLine = editor.LineFromPosition(editor.SelectionStart);
            int endLine = editor.LineFromPosition(editor.SelectionEnd);
            editor.SetSelection(editor.Lines[startLine].Position, editor.Lines[endLine].EndPosition - Environment.NewLine.Length);

            string selectedText = editor.SelectedText;

            if (!string.IsNullOrEmpty(selectedText))
            {
                List<string> lines = selectedText.Split(new[] { Environment.NewLine }, StringSplitOptions.None).ToList();

                if (lines.Where(p => p.Trim() != Environment.NewLine && p.Trim() != string.Empty).All(p => p.TrimStart().StartsWith(commentSymbol)))
                {
                    for (int i = 0; i < lines.Count; i++)
                        if (lines[i].Trim() != Environment.NewLine && lines[i].Trim() != string.Empty)
                            lines[i] = lines[i].Split(new[] { commentSymbol }, 2, StringSplitOptions.None)[1];
                }
                else
                {
                    for (int i = 0; i < lines.Count; i++)
                        if (lines[i].Trim() != Environment.NewLine && lines[i].Trim() != string.Empty)
                            lines[i] = commentSymbol + lines[i];
                }

                editor.ReplaceSelection(string.Join(Environment.NewLine, lines));
            }
            editor.SetSelection(editor.Lines[startLine].Position, editor.Lines[endLine].EndPosition - (editor.Lines[endLine].Text.EndsWith(Environment.NewLine) ? Environment.NewLine.Length : 0));
        }

        public static void IndentAfterReturn(Scintilla editor)
        {
            int currentPos = editor.CurrentPosition;
            int currentLine = editor.LineFromPosition(currentPos);
            editor.Lines[currentLine].Indentation = editor.Lines[currentLine - 1].Indentation;
            editor.GotoPosition(editor.Lines[currentLine].EndPosition - (editor.Lines[currentLine].Text.EndsWith(Environment.NewLine) ? Environment.NewLine.Length : 0));
        }

        public static void ReformatTextToSql(Scintilla editor, string text = null)
        {
            if (text == null)
                text = editor.SelectedText;
            if (string.IsNullOrWhiteSpace(text))
                return;

            List<char> DelimiterChars = new List<char> { ' ', @"'"[0], '(', ')', ',', '.', '\t', '\n', '\r', ';', '|' };
            text = $"({string.Join(", ", text.Split(DelimiterChars.ToArray(), StringSplitOptions.RemoveEmptyEntries).Select(p => $"\'{p.Trim()}\'").ToArray())})";
            editor.ReplaceSelection(text);
        }

        public static void SelectBlock(Scintilla editor, string blockStartIdentifier = "-----", string blockEndIdentifier = "-----") // Selects block that is serrounded by at least 5 '-'
        {
            int currentLine = editor.LineFromPosition(editor.CurrentPosition);

            // Find the start line
            int startLine = currentLine;
            if (string.IsNullOrWhiteSpace(blockStartIdentifier))
                startLine = 0;
            else
                while (startLine > 0)
                {
                    string lineText = editor.Lines[startLine].Text.Trim();
                    if (lineText.StartsWith(blockStartIdentifier))
                    {
                        break;
                    }
                    startLine--;
                }

            // Find the end line
            int endLine = currentLine;
            if (string.IsNullOrWhiteSpace(blockEndIdentifier))
                endLine = editor.Lines.Count - 1;
            else
                while (endLine < editor.Lines.Count - 1)
                {
                    string lineText = editor.Lines[endLine].Text.Trim();
                    if (lineText.StartsWith(blockEndIdentifier))
                    {
                        break;
                    }
                    endLine++;
                }

            // Select the lines
            int startPosition = editor.Lines[startLine].Position;
            int endPosition = editor.Lines[endLine].EndPosition;

            editor.SetSelection(startPosition, endPosition);
        }

        public static void WrapIntoSqlBlock(Scintilla editor)
        {
            int indentation = editor.Lines[editor.CurrentLine].Indentation;
            int selStartPos = editor.SelectionStart;
            int linesCount = editor.LineFromPosition(editor.SelectionEnd) - editor.LineFromPosition(selStartPos) + 1;
            string text = $"(\n{editor.SelectedText}\n)";
            editor.ReplaceSelection(text);
            editor.Update();
            editor.SetSelection(editor.Lines[editor.LineFromPosition(selStartPos) + 1].Position, editor.Lines[editor.LineFromPosition(selStartPos) - 1 + linesCount].EndPosition);
            editor.Update();

            editor.Lines[editor.LineFromPosition(selStartPos)].Indentation = indentation;
            editor.Lines[editor.LineFromPosition(selStartPos) + linesCount + 1].Indentation = indentation;

            for (int i = editor.LineFromPosition(editor.SelectionStart); i <= editor.LineFromPosition(editor.SelectionEnd); i++)
                editor.Lines[i].Indentation = Math.Max(indentation, editor.Lines[i].Indentation) + 4;
        }

        public static void MoveLineUp(Scintilla editor)
        {
            int startLine = editor.LineFromPosition(editor.SelectionStart);
            int endLine = editor.LineFromPosition(editor.SelectionEnd);
            editor.SetSelection(editor.Lines[startLine].Position, editor.Lines[endLine].EndPosition - Environment.NewLine.Length);

            string selectedText = editor.SelectedText;

            if (startLine > 0)
            {
                editor.DeleteRange(editor.Lines[startLine].Position, selectedText.Length + Environment.NewLine.Length);
                editor.InsertText(editor.Lines[startLine - 1].Position, selectedText + Environment.NewLine);

                editor.SetSelection(editor.Lines[startLine - 1].Position, editor.Lines[endLine - 1].EndPosition - (editor.Lines[endLine].Text.EndsWith(Environment.NewLine) ? Environment.NewLine.Length : 0));
            }
        }

        public static void MoveLineDown(Scintilla editor)
        {
            int startLine = editor.LineFromPosition(editor.SelectionStart);
            int endLine = editor.LineFromPosition(editor.SelectionEnd);
            editor.SetSelection(editor.Lines[startLine].Position, editor.Lines[endLine].EndPosition - Environment.NewLine.Length);

            string selectedText = editor.SelectedText;

            if (endLine < editor.Lines.Count - 1)
            {
                editor.DeleteRange(editor.Lines[startLine].Position, selectedText.Length + Environment.NewLine.Length);
                editor.InsertText(editor.Lines[endLine - (endLine - startLine)].EndPosition, selectedText + Environment.NewLine);

                editor.SetSelection(editor.Lines[startLine + 1].Position, editor.Lines[endLine + 1].EndPosition - (editor.Lines[endLine].Text.EndsWith(Environment.NewLine) ? Environment.NewLine.Length : 0));
            }
        }

        public static void SetupVbaEditor(Scintilla editor)
        {
            editor.LexerName = "vb";
            editor.ReadOnly = true;

            editor.StyleClearAll();
            editor.CaretLineBackColor = Color.FromArgb(0, 255, 255, 255);
            editor.Styles[Style.Default].BackColor = Color.FromArgb(30, 30, 30);
            editor.Styles[Style.Default].Font = "Consolas";
            editor.Styles[Style.Default].Size = 9;
            editor.Styles[Style.Default].Bold = true;
            editor.Margins[0].Width = 25;
            editor.Margins[1].Width = 8;

            // Set VBA syntax highlighting styles similar to Visual Studio Dark Theme
            editor.Styles[Style.Vb.Default].ForeColor = Color.FromArgb(240, 240, 240); // Almost white
            editor.Styles[Style.Vb.Comment].ForeColor = Color.FromArgb(100, 100, 100); // Gray
            editor.Styles[Style.Vb.Number].ForeColor = Color.FromArgb(214, 157, 133); // Orange
            editor.Styles[Style.Vb.Keyword].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Vb.Keyword2].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Vb.Keyword3].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Vb.Keyword4].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Vb.DocKeyword].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Vb.String].ForeColor = Color.FromArgb(181, 220, 168); // Light green
            editor.Styles[Style.Vb.Label].ForeColor = Color.FromArgb(181, 220, 168); // Light green
            editor.Styles[Style.Vb.Operator].ForeColor = Color.FromArgb(240, 240, 240); // Almost white
            editor.Styles[Style.Vb.Identifier].ForeColor = Color.FromArgb(240, 240, 240); // Almost white
            editor.Styles[Style.LineNumber].ForeColor = Color.FromArgb(100, 100, 100); // Gray

            // Set VBA code's background color
            editor.Styles[Style.Vb.Default].BackColor = Color.FromArgb(30, 30, 30);
            editor.Styles[Style.Vb.Comment].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Vb.Number].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Vb.Keyword].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Vb.String].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Vb.Operator].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Vb.Identifier].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.LineNumber].BackColor = Color.FromArgb(15, 15, 15);

            /*editor.ClearCmdKey(Keys.Control | Keys.Oem2);
            editor.ClearCmdKey(Keys.Control | Keys.Divide);
            editor.ClearCmdKey(Keys.Shift | Keys.Control | Keys.Divide);
            editor.ClearCmdKey(Keys.Shift | Keys.Control | Keys.Oem2);*/
        }

        private static void InsertMatchingBracket(Scintilla editor, char addedChar, char closingChar)
        {
            int currentPos = editor.CurrentPosition;
            bool isNotBracket = new char[] { '"', '\'' }.Contains(closingChar);

            if (!isNotBracket)
            {
                editor.InsertText(currentPos, addedChar.ToString());
                if (editor.BraceMatch(currentPos) != Scintilla.InvalidPosition)
                {
                    editor.DeleteRange(currentPos - 1, 1);
                    editor.GotoPosition(currentPos);
                    return;
                }
                editor.DeleteRange(currentPos - 1, 1);
            }

            editor.InsertText(currentPos, closingChar.ToString());
            editor.GotoPosition(currentPos);
        }

        private static void SkipClosingBracket(Scintilla editor, char openingChar, char closingChar)
        {
            int currentPos = editor.CurrentPosition;
            int nextChar = editor.GetCharAt(currentPos - 1);
            int previousChar = editor.GetCharAt(currentPos - 2);

            if (nextChar == closingChar && previousChar == openingChar && editor.BraceMatch(currentPos - 2) != Scintilla.InvalidPosition)
            {
                editor.DeleteRange(currentPos - 1, 1);
                editor.GotoPosition(currentPos);
            }
        }

        private static void HighlightMatchingBrackets(Scintilla editor, int indicatorIndexForBrackets)
        {
            if (editor == null || !editor.Indicators.Select(p => p.Index).Contains(indicatorIndexForBrackets))
                return;

            int position = editor.CurrentPosition;

            // Check for an opening or closing bracket at the current position
            if (editor.GetCharAt(position) == '(')
            {
                int matchPos = editor.BraceMatch(position);
                if (matchPos != Scintilla.InvalidPosition)
                {
                    // Highlight the matching brackets
                    editor.IndicatorCurrent = indicatorIndexForBrackets;
                    editor.IndicatorFillRange(position, 1); // Highlight the opening bracket
                    editor.IndicatorFillRange(matchPos, 1); // Highlight the matching closing bracket
                }
            }
            else if (editor.GetCharAt(position) == ')')
            {
                int matchPos = editor.BraceMatch(position);
                if (matchPos != Scintilla.InvalidPosition)
                {
                    // Highlight the matching brackets
                    editor.IndicatorCurrent = indicatorIndexForBrackets;
                    editor.IndicatorFillRange(position, 1); // Highlight the closing bracket
                    editor.IndicatorFillRange(matchPos, 1); // Highlight the matching opening bracket
                }
            }
        }

        private static void HighlightVariables(Scintilla editor, int indicatorIndex)
        {
            if (editor == null || !editor.Indicators.Select(p => p.Index).Contains(indicatorIndex))
                return;

            // Define the regex pattern to match words that start with ":::" and are not connected to other characters
            string pattern = @"(?<!\S):::\w+";
            var matches = Regex.Matches(editor.Text, pattern);
            foreach (Match match in matches)
            {
                // Apply the indicator to the matched range
                foreach (var ind in editor.Indicators.Select(p => p.Index).Where(p => p != indicatorIndex))
                {
                    editor.IndicatorCurrent = ind;
                    editor.IndicatorClearRange(match.Index, match.Length);
                }
                editor.IndicatorCurrent = indicatorIndex;
                editor.IndicatorFillRange(match.Index, match.Length);
            }
        }

        private static void HighlightCustomWords(Scintilla editor, int indicatorIndex, string search, string regexPatternL = @"(?i)", string regexPatternR = "")
        {
            if (editor == null || search == null || !editor.Indicators.Select(p => p.Index).Contains(indicatorIndex))
                return;

            // Define the regex pattern to match words that start with ":::" and are not connected to other characters
            string pattern = regexPatternL + Regex.Escape(search) + regexPatternR;
            var matches = Regex.Matches(editor.Text, pattern);
            foreach (Match match in matches)
            {
                // Apply the indicator to the matched range except selection
                if (editor.SelectionStart == match.Index)
                    continue;
                editor.IndicatorCurrent = indicatorIndex;
                editor.IndicatorFillRange(match.Index, match.Length);
            }
        }

        private static void Editor_DragOver(object sender, DragEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.Move;
        }

        private static void Editor_DragEnter(object sender, DragEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.Move;
        }

        private static void Editor_DragDrop(object sender, DragEventArgs e)
        {
            Scintilla editor = sender as Scintilla;
            Point point = editor.PointToClient(new Point(e.X, e.Y));
            int insertPos = editor.CharPositionFromPoint(point.X, point.Y);
            string selectedText = editor.SelectedText;
            int startSelection = editor.SelectionStart;
            int endSelection = editor.SelectionEnd;

            // Insert the selected text at the new position
            editor.InsertText(insertPos, selectedText);

            if (Control.ModifierKeys == Keys.Control)
                return;

            // If the new position is before the original position adjust the original position
            if (insertPos < startSelection)
            {
                startSelection += selectedText.Length;
                endSelection += selectedText.Length;
            }

            // Remove the selected text from the original position
            editor.DeleteRange(startSelection, endSelection - startSelection);
        }
    }
}

