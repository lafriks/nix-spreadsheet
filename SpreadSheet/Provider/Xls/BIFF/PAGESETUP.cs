/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2012, Lauris Bukšis-Haberkorns <lauris@nix.lv>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */

using System;
using System.Collections.Generic;
using System.Text;
using Nix.CompoundFile;
using Nix.SpreadSheet.PageSettings;

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
    class PAGESETUP : BIFFRecord
    {
        public PageSetting PageSettings { get; set; }

        /// <summary>
        /// BOF OPCODE
        /// </summary>
        protected override ushort OPCODE
        {
            get
            {
                return 0x00A1;
            }
        }

        private ushort GetPageSize(PageSize size)
        {
            switch (size)
            {
                default:
                case PageSize.Default:
                    return 0;
                case PageSize.Letter:
                    return 1;
                case PageSize.LetterSmall:
                    return 2;
                case PageSize.Tabloid:
                    return 3;
                case PageSize.Ledger:
                    return 4;
                case PageSize.Legal:
                    return 5;
                case PageSize.Statement:
                    return 6;
                case PageSize.Executive:
                    return 7;
                case PageSize.A3:
                    return 8;
                case PageSize.A4:
                    return 9;
                case PageSize.A4Small:
                    return 10;
                case PageSize.A5:
                    return 11;
                case PageSize.A6:
                    return 70;
                case PageSize.B4:
                    return 12;
                case PageSize.B5:
                    return 13;
                case PageSize.B6:
                    return 88;
                case PageSize.Folio:
                    return 14;
                case PageSize.Quarto:
                    return 15;
                case PageSize.TenByFourteen:
                    return 16;
                case PageSize.ElevenBySeventeen:
                    return 17;
                case PageSize.Note:
                    return 18;
                case PageSize.Envelope9:
                    return 19;
                case PageSize.Envelope10:
                    return 20;
                case PageSize.Envelope11:
                    return 21;
                case PageSize.Envelope12:
                    return 22;
                case PageSize.Envelope14:
                    return 23;
                case PageSize.Csheet:
                    return 24;
                case PageSize.Dsheet:
                    return 25;
                case PageSize.Esheet:
                    return 26;
                case PageSize.EnvelopeDL:
                    return 27;
                case PageSize.EnvelopeC5:
                    return 28;
                case PageSize.EnvelopeC3:
                    return 29;
                case PageSize.EnvelopeC4:
                    return 30;
                case PageSize.EnvelopeC6:
                    return 31;
                case PageSize.EnvelopeC65:
                    return 32;
                case PageSize.EnvelopeB4:
                    return 33;
                case PageSize.EnvelopeB5:
                    return 34;
                case PageSize.EnvelopeB6:
                    return 35;
                case PageSize.EnvelopeItaly:
                    return 36;
                case PageSize.EnvelopeMonarch:
                    return 37;
                case PageSize.EnvelopePersonal:
                    return 38;
                case PageSize.FanfoldUS:
                    return 39;
                case PageSize.FanfoldStdGerman:
                    return 40;
                case PageSize.FanfoldLegalGerman:
                    return 41;
            }
        }

        /// <summary>
        /// Writes BIFF record to the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public override void Write(EndianStream stream)
        {
            this.WriteHeader(stream, 34);
            stream.WriteUInt16(GetPageSize(PageSettings.PapeSize)); // Page size
            stream.WriteUInt16((UInt16)100); // Scaling factor in percent
            stream.WriteUInt16((UInt16)1); // Start page number
            stream.WriteUInt16((UInt16)0); // Fit worksheet width to this number of pages (0 = use as many as needed)
            stream.WriteUInt16((UInt16)0); // Fit worksheet height to this number of pages (0 = use as many as needed)
            ushort options = 0;
            if (!PageSettings.PrintPagesInColumns)
                options |= 0x0001;
            if (!PageSettings.Landscape)
                options |= 0x0002;
            //  Paper size, scaling factor, paper orientation (portrait/landscape), 
            // print resolution and number of copies are not initialised
            //options |= 0x0004;
            if (!PageSettings.PrintColoured)
                options |= 0x0008;
            if (!PageSettings.PrintQualityDefault)
                options |= 0x0010;
            if (PageSettings.PrintNotes)
                options |= 0x0020;
            // Use default paper orientation (landscape for chart sheets, portrait otherwise)
            //options |= 0x0040;
            // Use start page number above
            //options |= 0x0080;
            // Print notes at end of sheet
            //options |= 0x0200;
            stream.WriteUInt16(options); // Option flags
            stream.WriteUInt16((UInt16)0); // Print resolution in dpi
            stream.WriteUInt16((UInt16)0); // Vertical print resolution in dpi
            stream.WriteDoubleIEEE((double)0.5); // Header margin (IEEE 754 floating-point value, 64-bit double precision)
            stream.WriteDoubleIEEE((double)0.5); // Footer margin (IEEE 754 floating-point value, 64-bit double precision)
            stream.WriteUInt16((UInt16)1); // Number of copies to print
        }

        /// <summary>
        /// Reads BIFF record from the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public override void Read(EndianStream stream)
        {
        }
    }
}
