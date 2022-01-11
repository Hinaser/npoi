/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.UserModel.Helpers
{
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System.Collections.Generic;
    using System.Linq;

    public abstract class CellShifter
    {
        protected ISheet sheet;

        public CellShifter(ISheet sh)
        {
            sheet = sh;
        }

        public void ShiftUpMergedRegionsOnRowRemoval(int startRow, int startCol, int endRow, int endCol)
        {
            int nRowUp = endRow - startRow + 1;
            ISet<int> removedIndices = new HashSet<int>();
            List<CellRangeAddress> newMergedRegions = new List<CellRangeAddress>();
            int size = sheet.NumMergedRegions;
            for (int i = 0; i < size; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);

                if(merged.FirstRow > startRow)
                {
                    removedIndices.Add(i);
                    var newMergedRegion = new CellRangeAddress(
                        merged.FirstRow - nRowUp,
                        merged.LastRow - nRowUp,
                        merged.FirstColumn,
                        merged.LastColumn
                    );
                    newMergedRegions.Add(newMergedRegion);
                }
            }

            if(removedIndices.Count == 0)
            {
                return;
            }

            sheet.RemoveMergedRegions(removedIndices.ToList());
            foreach(var cra in newMergedRegions)
            {
                sheet.AddMergedRegion(cra);
            }
        }

        /**
         *  Remove merged regions where it overlaps with a removing cell range.
         *  
         * @param startRow the row to start removing
         * @param startCol the col to start removing
         * @param endRow   the row to end removing
         * @param endCol   the col to end removing
         */
        public void RemoveMergedRegions(int startRow, int startCol, int endRow, int endCol)
        {
            ISet<int> removedIndices = new HashSet<int>();
            int size = sheet.NumMergedRegions;
            for (int i = 0; i < size; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);

                bool notOverlapped = (merged.FirstRow > endRow || merged.FirstColumn > endCol)
                    || (merged.LastRow < startRow || merged.LastColumn < startCol);
                // remove merged region that overlaps Shifting
                if (!notOverlapped)
                {
                    removedIndices.Add(i);
                }
            }

            if (!(removedIndices.Count == 0)/*.IsEmpty()*/)
            {
                sheet.RemoveMergedRegions(removedIndices.ToList());
            }
        }

        /**
         * Updated named ranges
         */
        public abstract void UpdateNamedRanges(FormulaShifter Shifter);

        /**
         * Update formulas.
         */
        public abstract void UpdateFormulas(FormulaShifter Shifter);

        /**
         * Update the formulas in specified row using the formula Shifting policy specified by Shifter
         *
         * @param row the row to update the formulas on
         * @param Shifter the formula Shifting policy
         */

        public abstract void UpdateRowFormulas(IRow row, FormulaShifter Shifter);

        public abstract void UpdateConditionalFormatting(FormulaShifter Shifter);

        /**
         * Shift the Hyperlink anchors (not the hyperlink text, even if the hyperlink
         * is of type LINK_DOCUMENT and refers to a cell that was Shifted). Hyperlinks
         * do not track the content they point to.
         *
         * @param Shifter the formula Shifting policy
         */
        public abstract void UpdateHyperlinks(FormulaShifter Shifter);
    }
}