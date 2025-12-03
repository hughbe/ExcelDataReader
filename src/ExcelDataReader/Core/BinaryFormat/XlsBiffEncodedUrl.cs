using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents BIFF encoded URL.
/// </summary>
internal static class XlsBiffEncodedUrl
{
    public static string Decode(XlsBiffRecord record, int biffVersion, int offset, out int bytesRead, out bool isSelf)
    {
        bytesRead = 0;

        // The cch field gives the length of the supporting document name, which is
        // contained in the rgch field.
        int cch;
        if (biffVersion <= 5)
        {
            cch = record.ReadByte(offset + bytesRead);
            bytesRead += 1;
        }
        else
        {
            cch = record.ReadUInt16(offset + bytesRead);
            bytesRead += 2;
        }

        if (cch == 0)
        {
            isSelf = true;
            return string.Empty;
        }

        var isASCII = biffVersion <= 5;
        if (biffVersion >= 8)
        {
            var flags = record.ReadByte(offset + bytesRead);
            bytesRead += 1;
            isASCII = (flags & 0x01) == 0;
        }

        // Whenever document names are encoded to make BIFF files
        // compatible with file systems other than DOS. Encoded
        // document names are identified by the first character of the
        // rgch field.  The following special characters are recognized:
        //    Name    Value  Meaning
        // ----    -----  -------
        // chEmpty    0   empty sheetname
        // chEncode   1   encoded pathname
        // chSelf     2   self-referential external reference
        var chFirst = record.ReadByte(offset + bytesRead);
        if (chFirst == '\x00')
        {
            // chEmpty is used to store an external reference to the empty
            // sheet, as in the formula =!$A$1.
            bytesRead += 1;
            isSelf = false;
            return string.Empty;
        }
        else if (chFirst == '\x01')
        {
            // chEncode is used when the DOS file name of the supporting document
            // has been translated to a less system-dependent name.
            // The following special characters are recognized in an encoded
            // document name:
            // Name    Value  Related DOS keys
            // ----    -----  ----------------
            // chVolume   1   :
            // chSameVolume   2   none
            // chDownDir  3   .\
            // chUpDir    4   ..\
            // The chVolume key is used to specify a DOS drive letter in a document
            // name.  It is followed by the drive letter. This replaces the
            // DOS-specific ':' character, as in =C:SALES.XLS!Gross.
            // The chSameVolume key is used when the drive letter was omitted, to
            // indicate that the supporting document is on the same DOS drive as
            // the dependent document, as in =SALES.XLS!Gross.
            // The chDownDir key is used to go down a directory level.  It
            // is followed by the subdirectory name.  This replaces the
            // implicit DOS-specific sequence ".\", meaning subdirectory of
            // the current directory.  An example of such an external
            // reference is =AUGUST\SALES.XLS!Gross.
            string encodedUrl;
            if (isASCII)
            {
                encodedUrl = record.ReadAsciiString(offset + bytesRead, cch);
                bytesRead += cch;
            }
            else
            {
                encodedUrl = record.ReadUnicodeStringNoCch(offset + bytesRead, cch, out var bytesRead1);
                bytesRead += bytesRead1;
            }

            StringBuilder urlBuilder = new StringBuilder();
            for (int i = 1; i < encodedUrl.Length; i++)
            {
                char ch = encodedUrl[i];
                switch (ch)
                {
                    case '\x01':
                        {
                            // An MS-DOS drive letter will follow, or “@” and the server name of a UNC path
                            if (i + 1 < encodedUrl.Length)
                            {
                                // Next character is the drive letter.
                                i++;
                                char driveLetter = encodedUrl[i];
                                urlBuilder.Append(driveLetter);
                                urlBuilder.Append(':');
                                urlBuilder.Append('\\');
                            }

                            break;
                        }

                    case '\x02':
                        // Start path name on same drive as own document
                        continue;

                    case '\x03':
                        // End of subdirectory name
                        urlBuilder.Append(".\\");
                        break;

                    case '\x04':
                        // Parent directory
                        urlBuilder.Append("..\\");
                        break;

                    case '\x05':
                        // Unencoded URL. Followed by the length of the URL (1 byte), and the URL itself.
                        if (i + 1 < encodedUrl.Length)
                        {
                            // Read the length byte
                            i++;
                            int urlLength = (byte)encodedUrl[i];

                            // Read the URL
                            for (int j = 0; j < urlLength; j++, i++)
                            {
                                urlBuilder.Append(encodedUrl[i]);
                            }
                        }

                        break;

                    case '\x06':
                        // Start path name in installation directory of Excel
                        continue;

                    case '\x07':
                        // Macro template directory in installation directory of Excel
                        continue;

                    case '\x08':
                        // Sheet in the same workbook (BIFF4W)
                        continue;

                    default:
                        urlBuilder.Append(ch);
                        break;
                }
            }

            isSelf = false;
            return urlBuilder.ToString();
        }
        else if (chFirst == '\x02')
        {
            // chSelf is used to store an external reference where the dependent
            // and supporting documents are the same, for example a worksheet
            // SALES.XLS which contains the formula =SALES.XLS!$A$1.
            // There is no way to store the actual document name in this case,
            // so we use a placeholder.
            isSelf = true;
            return string.Empty;
        }
        else if (chFirst == '\x03')
        {
            // The string length field is decreased by 1, if the EXTERNSHEET stores a
            // reference to one of the own sheets (first character is 03H). Example:
            // The formula =Sheet2!A1 contains a reference to an EXTERNSHEET record with
            // the string “<03H>Sheet2”. The string consists of 7 characters but the
            // string length field contains the value 6.
            string encodedUrl;
            if (isASCII)
            {
                encodedUrl = record.ReadAsciiString(offset + bytesRead, cch);
                bytesRead += cch;
            }
            else
            {
                encodedUrl = record.ReadUnicodeStringNoCch(offset + bytesRead, cch, out var bytesRead1);
                bytesRead += bytesRead1;
            }

            isSelf = true;
            return encodedUrl.Replace('\x03', '|');
        }
        else if (chFirst == '\x04')
        {
            // Reference to the own workbook, sheet is unspecified
            // (nothing will follow)
            isSelf = true;
            return string.Empty;
        }
        else
        {
            // No special encoding.
            //  DDE references are encoded differently.  Only one translation
            // is ever performed on a DDE reference, on the '|' character:
            // Name    Value  Related DOS keys
            // ----    -----  ----------------
            // chDde   3  |
            isSelf = false;
            string encodedUrl;
            if (isASCII)
            {
                encodedUrl = record.ReadAsciiString(offset + bytesRead, cch);
                bytesRead += cch;
            }
            else
            {
                encodedUrl = record.ReadUnicodeStringNoCch(offset + bytesRead, cch, out var bytesRead1);
                bytesRead += bytesRead1;
            }

            return encodedUrl.Replace('\x03', '|');
        }
    }
}
