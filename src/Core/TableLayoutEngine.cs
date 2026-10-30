using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace XlsPrinter.Core
{
    public class GridCellData
    {
        public string Text { get; set; }
        public bool IsBold { get; set; }
        public bool IsRtl { get; set; }
        public float X { get; set; }
        public float Y { get; set; }

        public GridCellData()
        {
            Text = string.Empty;
        }
    }

    public class GridRowData
    {
        public int PageNumber { get; set; }
        public int SourceRowIndex { get; set; }
        public float Y { get; set; }
        public bool Selected { get; set; }
        public GridCellData[] Cells { get; set; }

        public GridRowData(int columnCount)
        {
            Selected = true;
            Cells = new GridCellData[columnCount];
            for (int i = 0; i < columnCount; i++)
                Cells[i] = new GridCellData();
        }
    }

    public class ColumnSpan
    {
        public float MinX { get; set; }
        public float MaxX { get; set; }
        public float RightEdge { get; set; }
        public int ElementCount { get; set; }

        public float CenterX
        {
            get { return (MinX + MaxX) / 2f; }
        }

        public ColumnSpan(float minX, float maxX)
        {
            MinX = minX;
            MaxX = maxX;
            RightEdge = maxX;
            ElementCount = 1;
        }

        public void Expand(float x, float right)
        {
            if (x < MinX) MinX = x;
            if (right > MaxX) MaxX = right;
            RightEdge = right;
            ElementCount++;
        }
    }

    public class ReconstructedTable
    {
        public List<float> ColumnCentersX { get; set; }
        public List<string> ColumnHeaders { get; set; }
        public List<bool> SelectedColumns { get; set; }
        public List<GridRowData> Rows { get; set; }
        public bool IsPredominantlyRtl { get; set; }

        public ReconstructedTable()
        {
            ColumnCentersX = new List<float>();
            ColumnHeaders = new List<string>();
            SelectedColumns = new List<bool>();
            Rows = new List<GridRowData>();
        }
    }

    /// <summary>
    /// Reconstructs structured rows and columns from 2D-positioned GDI text elements.
    /// Supports filtering by page region (selection rectangle), automatic Y-row clustering,
    /// X-column alignment, and Right-to-Left (Arabic) column ordering.
    /// </summary>
    public class TableLayoutEngine
    {
        public float RowVerticalToleranceMm { get; set; }
        public float ColumnHorizontalToleranceMm { get; set; }
        public bool ReverseColumnsForRtl { get; set; }
        public bool EliminateRepeatedHeadersAndFooters { get; set; }
        public bool ReverseVisualArabic { get; set; }
        public bool NormalizePresentationForms { get; set; }

        public TableLayoutEngine()
        {
            RowVerticalToleranceMm = 2.2f;       // Elements within 2.2mm vertically belong to the same row
            ColumnHorizontalToleranceMm = 8.0f;  // Elements within 8mm horizontally snap to the same column
            ReverseColumnsForRtl = false;
            EliminateRepeatedHeadersAndFooters = true;
            ReverseVisualArabic = true;
            NormalizePresentationForms = true;
        }

        /// <summary>
        /// Builds a structured table from all selected elements across pages (or within an optional crop rectangle).
        /// </summary>
        public ReconstructedTable BuildTable(PrintJobDocument doc, int pageFilter, RectangleF? regionFilterMm)
        {
            List<ExtractedTextElement> activeElements = new List<ExtractedTextElement>();
            int rtlCount = 0;

            for (int p = 0; p < doc.Pages.Count; p++)
            {
                ExtractedPage page = doc.Pages[p];
                if (pageFilter > 0 && page.PageNumber != pageFilter)
                    continue;

                for (int i = 0; i < page.Elements.Count; i++)
                {
                    ExtractedTextElement el = page.Elements[i];
                    if (!el.Selected) continue;

                    if (regionFilterMm.HasValue)
                    {
                        RectangleF r = regionFilterMm.Value;
                        if (r.Width > 1f && r.Height > 1f)
                        {
                            RectangleF elRect = new RectangleF(el.X, el.Y, Math.Max(1f, el.Width), Math.Max(1f, el.Height));
                            if (!r.IntersectsWith(elRect))
                                continue;
                        }
                    }

                    activeElements.Add(el);
                    if (el.IsRtl) rtlCount++;
                }
            }

            ReconstructedTable table = new ReconstructedTable();
            if (activeElements.Count == 0)
                return table;

            table.IsPredominantlyRtl = (rtlCount * 2 >= activeElements.Count);

            // Step 1: Group elements by (PageNumber, Y-row cluster)
            List<List<ExtractedTextElement>> rawRows = ClusterIntoRows(activeElements);

            // Step 2: Discover natural X column spans across all rows
            List<ColumnSpan> colSpans = DiscoverColumnSpans(rawRows);
            if (colSpans.Count == 0)
                colSpans.Add(new ColumnSpan(0f, 100f));

            float[] boundaries = new float[Math.Max(0, colSpans.Count - 1)];
            for (int b = 0; b < boundaries.Length; b++)
            {
                boundaries[b] = (colSpans[b].MaxX + colSpans[b + 1].MinX) / 2f;
            }

            // Step 3: Populate GridRowData for each clustered row
            List<GridRowData> prelimRows = new List<GridRowData>();
            bool[] colHasData = new bool[colSpans.Count];

            for (int r = 0; r < rawRows.Count; r++)
            {
                List<ExtractedTextElement> rowElems = rawRows[r];
                GridRowData gridRow = new GridRowData(colSpans.Count);
                gridRow.PageNumber = rowElems[0].PageNumber;
                gridRow.SourceRowIndex = r + 1;
                gridRow.Y = rowElems[0].Y;

                rowElems.Sort(delegate(ExtractedTextElement a, ExtractedTextElement b)
                {
                    return a.X.CompareTo(b.X);
                });

                for (int i = 0; i < rowElems.Count; i++)
                {
                    ExtractedTextElement el = rowElems[i];
                    int colIdx = FindColumnForElement(boundaries, colSpans, el);
                    if (ReverseColumnsForRtl)
                    {
                        colIdx = (colSpans.Count - 1) - colIdx;
                    }
                    colHasData[colIdx] = true;

                    GridCellData cell = gridRow.Cells[colIdx];
                    if (string.IsNullOrEmpty(cell.Text))
                    {
                        cell.Text = el.Text;
                        cell.IsBold = el.IsBold;
                        cell.IsRtl = el.IsRtl;
                        cell.X = el.X;
                        cell.Y = el.Y;
                    }
                    else
                    {
                        if (cell.IsRtl || el.IsRtl || ArabicTextHelper.ContainsArabic(cell.Text) || ArabicTextHelper.ContainsArabic(el.Text))
                        {
                            // In an RTL cell, the element with larger X (further right) was read first in logical order
                            if (el.X > cell.X)
                            {
                                cell.Text = el.Text + " " + cell.Text;
                            }
                            else
                            {
                                cell.Text = cell.Text + " " + el.Text;
                            }
                        }
                        else
                        {
                            cell.Text = cell.Text + " " + el.Text;
                        }
                        cell.IsBold = cell.IsBold || el.IsBold;
                        cell.IsRtl = cell.IsRtl || el.IsRtl;
                    }
                }

                prelimRows.Add(gridRow);
            }

            // Step 3b: Merge wrapped line continuations and complementary summary sub-lines
            List<GridRowData> mergedRows = new List<GridRowData>();
            for (int r = 0; r < prelimRows.Count; r++)
            {
                GridRowData curr = prelimRows[r];
                if (mergedRows.Count == 0)
                {
                    mergedRows.Add(curr);
                    continue;
                }

                GridRowData prev = mergedRows[mergedRows.Count - 1];
                float yDist = curr.Y - prev.Y;

                if (curr.PageNumber == prev.PageNumber && yDist >= 0.8f && yDist <= 4.8f)
                {
                    List<int> occCurr = new List<int>();
                    List<int> occPrev = new List<int>();
                    for (int c = 0; c < colSpans.Count; c++)
                    {
                        if (!string.IsNullOrEmpty(curr.Cells[c].Text)) occCurr.Add(c);
                        if (!string.IsNullOrEmpty(prev.Cells[c].Text)) occPrev.Add(c);
                    }

                    // Condition 1: Single cell wrapped line (continuation in a text cell like description)
                    if (occCurr.Count == 1 && occPrev.Contains(occCurr[0]))
                    {
                        int targetCol = occCurr[0];
                        string prevTxt = prev.Cells[targetCol].Text;
                        string currTxt = curr.Cells[targetCol].Text;

                        prev.Cells[targetCol].Text = prevTxt + " " + currTxt;
                        continue;
                    }

                    // Condition 2: Complementary sub-lines (disjoint columns within summary rows)
                    bool hasOverlap = false;
                    for (int i = 0; i < occCurr.Count; i++)
                    {
                        if (occPrev.Contains(occCurr[i]))
                        {
                            hasOverlap = true;
                            break;
                        }
                    }

                    if (!hasOverlap && yDist <= 4.2f)
                    {
                        for (int i = 0; i < occCurr.Count; i++)
                        {
                            int c = occCurr[i];
                            prev.Cells[c] = curr.Cells[c];
                        }
                        continue;
                    }
                }

                mergedRows.Add(curr);
            }
            prelimRows = mergedRows;

            // If a summary label (e.g. الاجمالي, المتوسط للفترة, متوسط سعر م3) landed in the Unit column adjacent to numbers,
            // normalize it to the Description column (the standard text column)
            for (int r = 0; r < prelimRows.Count; r++)
            {
                GridRowData row = prelimRows[r];
                if (colSpans.Count >= 5)
                {
                    int unitCol = ReverseColumnsForRtl ? 2 : colSpans.Count - 3;
                    int descCol = ReverseColumnsForRtl ? 1 : colSpans.Count - 2;
                    if (unitCol >= 0 && unitCol < colSpans.Count && descCol >= 0 && descCol < colSpans.Count)
                    {
                        string unitTxt = (row.Cells[unitCol].Text ?? string.Empty).Trim();
                        string descTxt = (row.Cells[descCol].Text ?? string.Empty).Trim();
                        if (string.IsNullOrEmpty(descTxt) && !string.IsNullOrEmpty(unitTxt))
                        {
                            if (unitTxt.Contains("الاجمال") || unitTxt.Contains("المتوسط") || unitTxt.Contains("متوسط") ||
                                unitTxt.Contains("المجموع") || unitTxt.ToLowerInvariant().Contains("total") || unitTxt.ToLowerInvariant().Contains("average"))
                            {
                                row.Cells[descCol] = row.Cells[unitCol];
                                row.Cells[unitCol] = new GridCellData();
                            }
                        }
                    }
                }
            }

            // Recompute colHasData from final prelimRows
            for (int c = 0; c < colSpans.Count; c++) colHasData[c] = false;
            for (int r = 0; r < prelimRows.Count; r++)
            {
                for (int c = 0; c < colSpans.Count; c++)
                {
                    if (!string.IsNullOrEmpty(prelimRows[r].Cells[c].Text))
                        colHasData[c] = true;
                }
            }

            // Remove any empty columns that received no cells after nearest-column snapping
            List<int> nonEmptyColIndices = new List<int>();
            for (int c = 0; c < colSpans.Count; c++)
            {
                if (colHasData[c])
                    nonEmptyColIndices.Add(c);
            }
            if (nonEmptyColIndices.Count == 0)
                nonEmptyColIndices.Add(0);

            for (int i = 0; i < nonEmptyColIndices.Count; i++)
            {
                int srcCol = nonEmptyColIndices[i];
                int spanIdx = ReverseColumnsForRtl ? (colSpans.Count - 1 - srcCol) : srcCol;
                table.ColumnCentersX.Add(colSpans[spanIdx].CenterX);
                table.ColumnHeaders.Add(GetExcelColumnName(i));
                table.SelectedColumns.Add(true);
            }
            if (EliminateRepeatedHeadersAndFooters)
            {
                prelimRows = RemoveRepeatedHeadersAndFooters(prelimRows);
            }

            for (int r = 0; r < prelimRows.Count; r++)
            {
                GridRowData srcRow = prelimRows[r];
                GridRowData finalRow = new GridRowData(nonEmptyColIndices.Count);
                finalRow.PageNumber = srcRow.PageNumber;
                finalRow.SourceRowIndex = srcRow.SourceRowIndex;
                finalRow.Y = srcRow.Y;
                finalRow.Selected = srcRow.Selected;
                for (int i = 0; i < nonEmptyColIndices.Count; i++)
                {
                    GridCellData cell = srcRow.Cells[nonEmptyColIndices[i]];
                    if (cell != null && !string.IsNullOrEmpty(cell.Text) && ArabicTextHelper.ContainsArabic(cell.Text))
                    {
                        cell.Text = ArabicTextHelper.NormalizeArabicText(cell.Text, NormalizePresentationForms, ReverseVisualArabic);
                    }
                    finalRow.Cells[i] = cell;
                }
                table.Rows.Add(finalRow);
            }

            return table;
        }

        private List<GridRowData> RemoveRepeatedHeadersAndFooters(List<GridRowData> rows)
        {
            if (rows == null || rows.Count == 0) return rows;

            int maxPage = 1;
            for (int r = 0; r < rows.Count; r++)
            {
                if (rows[r].PageNumber > maxPage) maxPage = rows[r].PageNumber;
            }

            if (maxPage <= 1) return rows;

            // 1. Identify the header Y limit and signatures from Page 1
            // Header block includes all rows up to and including the table column headers
            HashSet<string> page1HeaderSignatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> headerCellTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            float headerYLimit = 0f;

            for (int r = 0; r < rows.Count; r++)
            {
                GridRowData row = rows[r];
                if (row.PageNumber != 1) break;

                if (IsColumnHeaderRow(row))
                {
                    if (row.Y > headerYLimit)
                        headerYLimit = row.Y;
                }
            }

            // Fallback: if no explicit column header matched, use first data row Y
            if (headerYLimit <= 0f)
            {
                for (int r = 0; r < rows.Count; r++)
                {
                    GridRowData row = rows[r];
                    if (row.PageNumber != 1) break;
                    if (IsLikelyDataRow(row))
                    {
                        headerYLimit = row.Y - 1.0f;
                        break;
                    }
                }
            }
            if (headerYLimit <= 0f) headerYLimit = 48.0f;
            headerYLimit += 3.0f; // tolerance margin

            // Collect all header row signatures from Page 1 within headerYLimit
            for (int r = 0; r < rows.Count; r++)
            {
                GridRowData row = rows[r];
                if (row.PageNumber != 1) break;

                if (row.Y <= headerYLimit)
                {
                    string sig = GetRowSignature(row);
                    if (!string.IsNullOrEmpty(sig))
                        page1HeaderSignatures.Add(sig);

                    for (int c = 0; c < row.Cells.Length; c++)
                    {
                        string t = (row.Cells[c].Text ?? string.Empty).Trim();
                        if (t.Length >= 2)
                            headerCellTexts.Add(t);
                    }
                }
            }

            // 2. Filter rows across all pages
            List<GridRowData> filtered = new List<GridRowData>();
            for (int r = 0; r < rows.Count; r++)
            {
                GridRowData row = rows[r];

                // Remove page footer (page numbers like "صفحة 1 من 11", "Page 1 of 11")
                if (IsPageFooterRow(row, row.PageNumber, maxPage))
                {
                    continue;
                }

                // For subsequent pages (Page > 1), remove repeated headers within the header region
                if (row.PageNumber > 1 && row.Y <= headerYLimit)
                {
                    string sig = GetRowSignature(row);
                    if (page1HeaderSignatures.Contains(sig))
                    {
                        continue;
                    }

                    if (IsColumnHeaderRow(row))
                    {
                        continue;
                    }

                    if (IsAllHeaderCells(row, headerCellTexts))
                    {
                        continue;
                    }
                }

                filtered.Add(row);
            }

            return filtered;
        }

        private static bool IsColumnHeaderRow(GridRowData row)
        {
            int matches = 0;
            for (int c = 0; c < row.Cells.Length; c++)
            {
                string txt = (row.Cells[c].Text ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(txt)) continue;

                if (txt.Contains("التاريخ") || txt.Contains("فاتورة") || txt.Contains("بيان") ||
                    txt.Contains("وحدة") || txt.Contains("كمية") || txt.Contains("فئة") ||
                    txt.Contains("جملة") || txt.Contains("المدة") || txt.Contains("جنية") ||
                    txt.Contains("رصيد") || txt.Contains("مدين") || txt.Contains("دائن") ||
                    txt.ToLowerInvariant().Contains("date") || txt.ToLowerInvariant().Contains("invoice") ||
                    txt.ToLowerInvariant().Contains("description") || txt.ToLowerInvariant().Contains("qty") ||
                    txt.ToLowerInvariant().Contains("price") || txt.ToLowerInvariant().Contains("total"))
                {
                    matches++;
                }
            }
            return matches >= 2;
        }

        private static string GetRowSignature(GridRowData row)
        {
            StringBuilder sb = new StringBuilder();
            for (int c = 0; c < row.Cells.Length; c++)
            {
                string txt = (row.Cells[c].Text ?? string.Empty).Trim();
                if (!string.IsNullOrEmpty(txt))
                {
                    if (sb.Length > 0) sb.Append(" | ");
                    sb.Append(txt);
                }
            }
            return sb.ToString();
        }

        private static bool IsLikelyDataRow(GridRowData row)
        {
            int numericCount = 0;
            bool hasDate = false;

            for (int c = 0; c < row.Cells.Length; c++)
            {
                string txt = (row.Cells[c].Text ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(txt)) continue;

                if ((txt.Contains("/") || txt.Contains("-") || txt.Contains(".")) &&
                    txt.Length >= 6 && txt.Length <= 10)
                {
                    int digitCount = 0;
                    for (int i = 0; i < txt.Length; i++)
                    {
                        if (char.IsDigit(txt[i])) digitCount++;
                    }
                    if (digitCount >= 4) hasDate = true;
                }

                double num;
                if (double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out num))
                {
                    numericCount++;
                }
            }

            return hasDate || numericCount >= 2;
        }

        private static bool IsPageFooterRow(GridRowData row, int pageNum, int maxPage)
        {
            string sig = GetRowSignature(row);
            if (string.IsNullOrEmpty(sig)) return false;

            string lower = sig.ToLowerInvariant();
            if (lower.Contains("صفحة") || lower.Contains("صفحه") ||
                lower.Contains("page ") || lower.Contains("of " + maxPage))
            {
                return true;
            }

            if (pageNum < maxPage)
            {
                if (lower.Contains("ما قبله") || lower.Contains("المنقول") ||
                    lower.Contains("carried forward") || lower.Contains("brought forward"))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsAllHeaderCells(GridRowData row, HashSet<string> headerCellTexts)
        {
            int nonBlank = 0;
            int headerMatches = 0;

            for (int c = 0; c < row.Cells.Length; c++)
            {
                string txt = (row.Cells[c].Text ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(txt)) continue;
                nonBlank++;

                if (headerCellTexts.Contains(txt) ||
                    txt.Contains("التاريخ") || txt.Contains("فاتورة") || txt.Contains("بيان") ||
                    txt.Contains("وحدة") || txt.Contains("كمية") || txt.Contains("فئة") ||
                    txt.Contains("جملة") || txt.Contains("المدة") || txt.Contains("جنية") ||
                    txt.Contains("المخزن") || txt.Contains("الشركة") || txt.Contains("سنة"))
                {
                    headerMatches++;
                }
            }

            return nonBlank > 0 && headerMatches * 2 >= nonBlank;
        }

        private List<ColumnSpan> DiscoverColumnSpans(List<List<ExtractedTextElement>> rows)
        {
            List<List<ExtractedTextElement>> tableRows = new List<List<ExtractedTextElement>>();
            int maxCount = 0;
            List<ExtractedTextElement> bestRow = null;

            for (int r = 0; r < rows.Count; r++)
            {
                List<ExtractedTextElement> row = rows[r];
                if (row.Count >= 2)
                {
                    if (row.Count > maxCount)
                    {
                        maxCount = row.Count;
                        bestRow = row;
                    }
                }
            }

            for (int r = 0; r < rows.Count; r++)
            {
                List<ExtractedTextElement> row = rows[r];
                if (row.Count >= Math.Max(3, maxCount - 1))
                {
                    tableRows.Add(row);
                }
            }

            List<ColumnSpan> spans = new List<ColumnSpan>();
            if (bestRow == null || bestRow.Count == 0)
            {
                spans.Add(new ColumnSpan(0f, 100f));
                return spans;
            }

            List<ExtractedTextElement> seed = new List<ExtractedTextElement>(bestRow);
            seed.Sort(delegate(ExtractedTextElement a, ExtractedTextElement b) { return a.X.CompareTo(b.X); });

            for (int i = 0; i < seed.Count; i++)
            {
                spans.Add(new ColumnSpan(seed[i].X, seed[i].Right));
            }

            // For all other table rows, match elements monotonically to spans without merging across columns
            for (int r = 0; r < tableRows.Count; r++)
            {
                List<ExtractedTextElement> row = tableRows[r];
                if (row == bestRow) continue;

                List<ExtractedTextElement> sortedRow = new List<ExtractedTextElement>(row);
                sortedRow.Sort(delegate(ExtractedTextElement a, ExtractedTextElement b) { return a.X.CompareTo(b.X); });

                int lastAssignedSpan = -1;
                for (int i = 0; i < sortedRow.Count; i++)
                {
                    ExtractedTextElement el = sortedRow[i];
                    int bestIdx = -1;
                    float bestScore = float.MaxValue;

                    for (int s = lastAssignedSpan + 1; s < spans.Count; s++)
                    {
                        ColumnSpan sp = spans[s];
                        float overlap = Math.Min(el.Right, sp.MaxX) - Math.Max(el.X, sp.MinX);
                        float dist = 0f;
                        if (el.Right < sp.MinX) dist = sp.MinX - el.Right;
                        else if (el.X > sp.MaxX) dist = el.X - sp.MaxX;

                        float rightDiff = Math.Abs(el.Right - sp.RightEdge);
                        float leftDiff = Math.Abs(el.X - sp.MinX);
                        float centerDiff = Math.Abs(el.CenterX - sp.CenterX);

                        float score;
                        if (overlap > 0f)
                            score = -overlap;
                        else
                            score = dist * 2f + Math.Min(centerDiff, Math.Min(rightDiff, leftDiff));

                        if (score < bestScore)
                        {
                            bestScore = score;
                            bestIdx = s;
                        }
                    }

                    if (bestIdx >= 0 && (bestScore < 0f || bestScore < 15f))
                    {
                        bool canExpandLeft = (bestIdx == 0 || el.X > spans[bestIdx - 1].MaxX + 0.5f);
                        bool canExpandRight = (bestIdx == spans.Count - 1 || el.Right < spans[bestIdx + 1].MinX - 0.5f);
                        if (bestIdx == spans.Count - 2 && (i == sortedRow.Count - 1 || el.Right < sortedRow[i + 1].X - 0.5f))
                        {
                            canExpandRight = true;
                        }

                        if (canExpandLeft && el.X < spans[bestIdx].MinX)
                            spans[bestIdx].MinX = el.X;
                        if (canExpandRight && el.Right > spans[bestIdx].MaxX)
                            spans[bestIdx].MaxX = el.Right;
                        spans[bestIdx].RightEdge = el.Right;
                        spans[bestIdx].ElementCount++;
                        lastAssignedSpan = bestIdx;
                    }
                }
            }

            // Ensure spans are sorted by MinX and strictly bounded
            spans.Sort(delegate(ColumnSpan a, ColumnSpan b) { return a.MinX.CompareTo(b.MinX); });
            for (int i = 0; i < spans.Count - 1; i++)
            {
                if (spans[i].MaxX >= spans[i + 1].MinX)
                {
                    float mid = (spans[i].MaxX + spans[i + 1].MinX) / 2f;
                    spans[i].MaxX = mid - 0.1f;
                    spans[i + 1].MinX = mid + 0.1f;
                }
            }

            return spans;
        }

        private static int FindColumnForElement(float[] boundaries, List<ColumnSpan> spans, ExtractedTextElement el)
        {
            int bestSpan = -1;
            float maxOverlap = 0f;
            for (int i = 0; i < spans.Count; i++)
            {
                float overlap = Math.Min(el.Right, spans[i].MaxX) - Math.Max(el.X, spans[i].MinX);
                if (overlap > maxOverlap)
                {
                    maxOverlap = overlap;
                    bestSpan = i;
                }
            }
            if (bestSpan >= 0 && maxOverlap > 1.0f)
                return bestSpan;

            float cx = el.CenterX;
            for (int i = 0; i < boundaries.Length; i++)
            {
                if (cx < boundaries[i])
                    return i;
            }
            return spans.Count - 1;
        }

        private List<List<ExtractedTextElement>> ClusterIntoRows(List<ExtractedTextElement> elements)
        {
            // Sort by PageNumber, then vertical position Y
            List<ExtractedTextElement> sorted = new List<ExtractedTextElement>(elements);
            sorted.Sort(delegate(ExtractedTextElement a, ExtractedTextElement b)
            {
                if (a.PageNumber != b.PageNumber)
                    return a.PageNumber.CompareTo(b.PageNumber);
                return a.CenterY.CompareTo(b.CenterY);
            });

            List<List<ExtractedTextElement>> rows = new List<List<ExtractedTextElement>>();
            List<ExtractedTextElement> currentRow = new List<ExtractedTextElement>();
            float currentRowCenterY = 0f;
            int currentPage = -1;

            for (int i = 0; i < sorted.Count; i++)
            {
                ExtractedTextElement el = sorted[i];
                if (currentRow.Count == 0)
                {
                    currentRow.Add(el);
                    currentRowCenterY = el.CenterY;
                    currentPage = el.PageNumber;
                }
                else
                {
                    float dynamicTol = Math.Max(RowVerticalToleranceMm, Math.Min(el.Height, currentRow[0].Height) * 0.55f);
                    if (el.PageNumber == currentPage && Math.Abs(el.CenterY - currentRowCenterY) <= dynamicTol)
                    {
                        currentRow.Add(el);
                        // Update running average of row CenterY
                        float sum = 0f;
                        for (int k = 0; k < currentRow.Count; k++) sum += currentRow[k].CenterY;
                        currentRowCenterY = sum / currentRow.Count;
                    }
                    else
                    {
                        rows.Add(currentRow);
                        currentRow = new List<ExtractedTextElement>();
                        currentRow.Add(el);
                        currentRowCenterY = el.CenterY;
                        currentPage = el.PageNumber;
                    }
                }
            }

            if (currentRow.Count > 0)
                rows.Add(currentRow);

            return rows;
        }

        public static string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
    }
}
