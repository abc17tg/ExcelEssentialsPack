using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ScintillaNET;

namespace ExcelAddInByMarcinOlszewski.Scripts
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
            editor.Lexer = Lexer.Vb;
            editor.ReadOnly = true;

            editor.StyleClearAll();
            editor.CaretLineVisible = false;
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


        public static void SetupSqlEditor(Scintilla editor)
        {
            editor.DragEnter += Editor_DragEnter;
            editor.DragDrop += Editor_DragDrop;
            editor.DragOver += Editor_DragOver;
            editor.Lexer = Lexer.Sql;

            editor.StyleClearAll();
            editor.CaretLineVisible = false;
            editor.Styles[Style.Default].BackColor = Color.FromArgb(30, 30, 30);
            editor.Styles[Style.Default].Font = "Consolas";
            editor.Styles[Style.Default].Size = 10;
            editor.Styles[Style.Default].Bold = true;
            editor.Margins[0].Width = 25;
            editor.Margins[1].Width = 8;

            // Set SQL syntax highlighting styles similar to Visual Studio Dark Theme
            editor.Styles[Style.Sql.Default].ForeColor = Color.FromArgb(240, 240, 240); // Almost white
            editor.Styles[Style.Sql.Comment].ForeColor = Color.FromArgb(100, 100, 100); // Gray
            editor.Styles[Style.Sql.CommentLine].ForeColor = Color.FromArgb(100, 100, 100); // Gray
            editor.Styles[Style.Sql.CommentDoc].ForeColor = Color.FromArgb(100, 100, 100); // Gray
            editor.Styles[Style.Sql.Number].ForeColor = Color.FromArgb(214, 157, 133); // Orange
            editor.Styles[Style.Sql.Word].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Sql.Word2].ForeColor = Color.FromArgb(86, 156, 214); // Blue
            editor.Styles[Style.Sql.String].ForeColor = Color.FromArgb(181, 220, 168); // Light green
            editor.Styles[Style.Sql.Character].ForeColor = Color.FromArgb(181, 220, 168); // Light green
            editor.Styles[Style.Sql.Operator].ForeColor = Color.FromArgb(240, 240, 240); // Almost white
            editor.Styles[Style.Sql.Identifier].ForeColor = Color.FromArgb(240, 240, 240); // Almost white
            editor.Styles[Style.LineNumber].ForeColor = Color.FromArgb(100, 100, 100); // Gray

            editor.Styles[Style.Sql.Default].BackColor = Color.FromArgb(30, 30, 30);
            editor.Styles[Style.Sql.Comment].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.CommentLine].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.CommentDoc].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.Number].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.Word].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.Word2].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.String].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.Character].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.Operator].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.Sql.Identifier].BackColor = editor.Styles[Style.Default].BackColor;
            editor.Styles[Style.LineNumber].BackColor = Color.FromArgb(15, 15, 15); ;

            editor.Styles[Style.Sql.Word].Case = StyleCase.Upper;
            editor.Styles[Style.Sql.Word2].Case = StyleCase.Upper;

            editor.Styles[Style.Sql.Default].Bold = true;
            editor.Styles[Style.Sql.Comment].Bold = true;
            editor.Styles[Style.Sql.CommentLine].Bold = true;
            editor.Styles[Style.Sql.CommentDoc].Bold = true;
            editor.Styles[Style.Sql.Number].Bold = true;
            editor.Styles[Style.Sql.Word].Bold = true;
            editor.Styles[Style.Sql.Word2].Bold = true;
            editor.Styles[Style.Sql.String].Bold = true;
            editor.Styles[Style.Sql.Character].Bold = true;
            editor.Styles[Style.Sql.Operator].Bold = true;
            editor.Styles[Style.Sql.Identifier].Bold = true;
            // Set SQL keywords

            editor.SetKeywords(0, SqlKeywords);

            editor.ClearCmdKey(Keys.Control | Keys.Oem2);
            editor.ClearCmdKey(Keys.Control | Keys.Divide);
            editor.ClearCmdKey(Keys.Shift | Keys.Control | Keys.Divide);
            editor.ClearCmdKey(Keys.Shift | Keys.Control | Keys.Oem2);
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




        public readonly static string SqlKeywords = "absolute action add admin after aggregate alias all allocate alter and any are array as asc assertion at authorization before begin binary bit blob boolean both breadth by call cascade cascaded case cast catalog char character check class clob close collate collation column commit completion connect connection constraint constraints constructor continue corresponding count create cross cube current current_date current_path current_role current_time current_timestamp current_user cursor cycle data date day deallocate dec decimal declare default deferrable deferred delete depth deref desc describe descriptor destroy destructor deterministic diagnostics dictionary disconnect distinct domain double drop dynamic each else end end-exec equals escape every except exception exec execute external false fetch first float for foreign found free from full function general get global go goto grant group grouping having host hour identity ignore immediate in indicator initcap initialize initially inner inout input insert int integer intersect interval into is isolation iterate join key language large last lateral leading left length less level like limit local localtime localtimestamp locator map match minute modifies modify module month names national natural nchar nclob new next no none not null numeric object of off old on only open operation option or order ordinality out outer output over pad parameter parameters partial partition path postfix precision prefix preorder prepare preserve primary prior privileges procedure public read reads real recursive ref references referencing regexp_like regexp_replace regexp_substr relative restrict result return returns revoke right role rollback rollup round row_number rows savepoint schema scope scroll search second section select sequence session session_user set sets size smallint some space specific specifictype sql sqlexception sqlstate sqlwarning start state statement static structure sum system_user table temporary terminate than then time timestamp timezone_hour timezone_minute to to_char to_date to_number top trailing transaction translation treat trigger trim true under union unique unknown unnest update usage user using value values varchar variable varying view when whenever where with without work write year zone";
    }



}

