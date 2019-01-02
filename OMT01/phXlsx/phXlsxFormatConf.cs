using System;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OMT01.phXlsx {
    public class phXlsxFormatConf {
        private static phXlsxFormatConf _instance = null;

        public static phXlsxFormatConf getInstance() {
            if (_instance == null) {
                _instance = new phXlsxFormatConf();
            }

            return _instance;
        }

        Dictionary<string, int> font_map = new Dictionary<string, int>(); 
        Dictionary<string, int> fill_map = new Dictionary<string, int>(); 
        Dictionary<string, uint> numbering_map = new Dictionary<string, uint>(); 
        Dictionary<string, int> border_map = new Dictionary<string, int>(); 
        XmlDocument _doc = null;

        protected phXlsxFormatConf() {
            _doc = new XmlDocument();
            _doc.Load(@"..\..\resources\PhFormatConf.xml");
        }

        public void PushCellFormatsToStylesheet(Stylesheet ss) {
            pushFontsToStylesheet(ss);
            pushFillsToStylesheet(ss);
            pushNumberingsToStylesheet(ss);
            pushBordersToStylesheet(ss);
            foreach (KeyValuePair<string, int> iter in font_map) {
                Console.WriteLine(iter.Key + " -> " + iter.Value);
            }
            foreach (KeyValuePair<string, int> iter in fill_map) {
                Console.WriteLine(iter.Key + " -> " + iter.Value);
            }
            foreach (KeyValuePair<string, uint> iter in numbering_map) {
                Console.WriteLine(iter.Key + " -> " + iter.Value);
            }
            foreach (KeyValuePair<string, int> iter in border_map) {
                Console.WriteLine(iter.Key + " -> " + iter.Value);
            }
        }

        private void pushBordersToStylesheet(Stylesheet ss) {
            var borders = ss.Borders;

            var xn = _doc.SelectSingleNode("stylesheet/borders");
            var nlst = xn.SelectNodes("border");
            Console.WriteLine(nlst.Count);

            foreach (XmlNode f in nlst) {
                var border_id = f.Attributes.GetNamedItem("id").Value;
                var border = new Border();

                var left = f.SelectSingleNode("left");
                var left_style = left.Attributes.GetNamedItem("style").Value;
                var left_color = left.Attributes.GetNamedItem("color").Value;
                var lb = new LeftBorder() { Style = BorderStyleValues.Thin };
                Color lc = new Color() { Rgb = left_color };
                lb.Append(lc);
                border.Append(lb);

                var right = f.SelectSingleNode("right");
                var right_style = right.Attributes.GetNamedItem("style").Value;
                var right_color = right.Attributes.GetNamedItem("color").Value;
                var rb = new RightBorder() { Style = BorderStyleValues.Thin };
                Color rc = new Color() { Rgb = right_color };
                rb.Append(rc);
                border.Append(rb);

                var top = f.SelectSingleNode("top");
                var top_style = top.Attributes.GetNamedItem("style").Value;
                var top_color = top.Attributes.GetNamedItem("color").Value;
                var tb = new TopBorder() { Style = BorderStyleValues.Thin };
                Color tc = new Color() { Rgb = top_color };
                tb.Append(tc);
                border.Append(tb);

                var bottom = f.SelectSingleNode("bottom");
                var bottom_style = bottom.Attributes.GetNamedItem("style").Value;
                var bottom_color = bottom.Attributes.GetNamedItem("color").Value;
                var bb = new TopBorder() { Style = BorderStyleValues.Thin };
                Color bc = new Color() { Rgb = bottom_color };
                bb.Append(bc);
                border.Append(bb);

                borders.Append(border);

                var border_idx = borders.Elements<Border>().Count() - 1;
                border_map.Add(border_id, border_idx);
            }
        }

        private void pushNumberingsToStylesheet(Stylesheet ss) {
            var numberings = ss.NumberingFormats;

            var xn = _doc.SelectSingleNode("stylesheet/numberings");
            var nlst = xn.SelectNodes("numbering");
            Console.WriteLine(nlst.Count);

            foreach(XmlNode f in nlst) {
                var numbering_id = f.Attributes.GetNamedItem("id").Value;
                var numbering_idx = uint.Parse(f.Attributes.GetNamedItem("idx").Value);

                var numbering_code = f.Attributes.GetNamedItem("code").Value;
                Console.WriteLine(numbering_code);

                var nf = new NumberingFormat { NumberFormatId = numbering_idx, FormatCode = numbering_code };
                numberings.Append(nf);

                numbering_map.Add(numbering_id, numbering_idx);
            }
        }

        private void pushFillsToStylesheet(Stylesheet ss) {
            var fills = ss.Fills;

            var xn = _doc.SelectSingleNode("stylesheet/fills");
            var nlst = xn.SelectNodes("fill");
            Console.WriteLine(nlst.Count);

            foreach (XmlNode f in nlst) {
                var fill_id = f.Attributes.GetNamedItem("id").Value;

                var fill_type = f.Attributes.GetNamedItem("type").Value;
                Console.WriteLine(fill_type);
                var fill_color = f.Attributes.GetNamedItem("color").Value;
                Console.WriteLine(fill_color);

                Fill fill = new Fill();

                PatternFill pf = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor fc = new ForegroundColor() { Rgb = fill_color };
                BackgroundColor bc = new BackgroundColor() { Indexed = (UInt32Value)64U };

                pf.Append(fc);
                pf.Append(bc);
                fill.Append(pf);

                fills.Append(fill);
                var fill_idx = fills.Elements<Fill>().Count() - 1;
                fill_map.Add(fill_id, fill_idx);
            }
        }

        private void pushFontsToStylesheet(Stylesheet ss) {
            var fonts = ss.Fonts;

            var xn = _doc.SelectSingleNode("stylesheet/fonts");
            var nlst = xn.SelectNodes("font");
            Console.WriteLine(nlst.Count);

            foreach (XmlNode f in nlst) {
                var font_id = f.Attributes.GetNamedItem("id").Value;

                var font_family = f.Attributes.GetNamedItem("name").Value;
                Console.WriteLine(font_family);
                var font_size = Double.Parse(f.Attributes.GetNamedItem("size").Value);
                Console.WriteLine(font_size);
                var font_color = f.Attributes.GetNamedItem("color").Value;
                Console.WriteLine(font_color);
                var font_bold = Boolean.Parse(f.Attributes.GetNamedItem("bold").Value);
                Console.WriteLine(font_bold);

                Font font = new Font();

                FontSize fontSize = new FontSize() { Val = (Double)font_size };
                Color color = new Color() { Rgb = new HexBinaryValue(font_color) };
                FontName fontName = new FontName() { Val = font_family };
                FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
                FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

                if (font_bold) {
                    Bold bold = new Bold();
                    font.Append(bold);
                }

                font.Append(fontSize);
                font.Append(color);
                font.Append(fontName);
                font.Append(fontFamilyNumbering2);
                font.Append(fontScheme2);

                fonts.Append(font);

                var font_idx = fonts.Elements<Font>().Count() - 1;
                font_map.Add(font_id, font_idx);
            }
        }
    }
}