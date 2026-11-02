using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace PowerPointIntelligence.AddIn
{
    // Original geometric artwork. No external icon packs, files or Office image IDs.
    // Render at the requested pixel size, rather than shrinking a 32 px bitmap to 16 px.
    internal static class RibbonIcons
    {
        private static readonly Dictionary<string, object> Pictures = new Dictionary<string, object>(StringComparer.Ordinal);
        private static readonly Color Ink = Color.FromArgb(111, 129, 153);
        private static readonly Color Blue = Color.FromArgb(37, 99, 235);
        private static readonly Color Teal = Color.FromArgb(13, 148, 136);
        private static readonly Color Amber = Color.FromArgb(217, 119, 6);
        private static readonly Color Pale = Color.FromArgb(232, 240, 255);

        internal static string Key(string id, int size) { return "ppi_" + id + "_" + size; }
        internal static object Load(string key)
        {
            if (key == null || !key.StartsWith("ppi_", StringComparison.Ordinal)) return null;
            int separator = key.LastIndexOf('_'), size;
            if (separator <= 4 || !Int32.TryParse(key.Substring(separator + 1), out size) || (size != 16 && size != 32)) return null;
            string id = key.Substring(4, separator - 4);
            lock (Pictures)
            {
                object picture;
                if (Pictures.TryGetValue(key, out picture)) return picture;
                using (var bitmap = Render(id, size))
                {
                    if (bitmap == null) return null;
                    // AxHost creates the OLE picture's own bitmap handle. Cache the RCW
                    // for the ribbon lifetime; Office may request an image repeatedly.
                    picture = PictureBridge.Convert(bitmap);
                }
                Pictures.Add(key, picture);
                return picture;
            }
        }
        private sealed class PictureBridge : AxHost
        {
            private PictureBridge() : base("") { }
            internal static object Convert(Image image) { return GetIPictureDispFromPicture(image); }
        }
        internal static Bitmap Render(string id, int size)
        {
            if (size != 16 && size != 32) throw new ArgumentOutOfRangeException("size");
            if(String.IsNullOrWhiteSpace(id))return null;
            var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                g.ScaleTransform(size / 32f, size / 32f);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                if (!Draw(g, id)) { bitmap.Dispose(); return null; }
            }
            return bitmap;
        }
        private static void Line(Graphics g, Color color, float x1, float y1, float x2, float y2, float width = 2)
        { using (var p = new Pen(color, width) { StartCap = LineCap.Round, EndCap = LineCap.Round }) g.DrawLine(p, x1, y1, x2, y2); }
        private static void Rect(Graphics g, Color color, float x, float y, float w, float h)
        { using (var b = new SolidBrush(color)) g.FillRectangle(b, x, y, w, h); }
        private static void Outline(Graphics g, Color color, float x, float y, float w, float h)
        { using (var p = new Pen(color, 2)) g.DrawRectangle(p, x, y, w, h); }
        private static void Circle(Graphics g, Color color, float x, float y, float diameter, bool fill)
        { using (var p = new Pen(color, 2)) using (var b = new SolidBrush(color)) { if (fill) g.FillEllipse(b, x, y, diameter, diameter); else g.DrawEllipse(p, x, y, diameter, diameter); } }
        private static void Polygon(Graphics g, Color color, params PointF[] points)
        { using (var b = new SolidBrush(color)) g.FillPolygon(b, points); }
        private static void Letter(Graphics g, string text, float x, float y, float size = 13, Color? color = null)
        {
            using (var font = new Font("Segoe UI", size, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(color ?? Ink))
                g.DrawString(text, font, brush, x, y, StringFormat.GenericTypographic);
        }
        private static void Arrow(Graphics g, Color c, float x1, float y1, float x2, float y2, bool both = false)
        {
            Line(g, c, x1, y1, x2, y2);
            double angle = Math.Atan2(y2 - y1, x2 - x1);
            Action<float, float, double> head = (x, y, a) => {
                Line(g, c, x, y, x - 4 * (float)Math.Cos(a - .65), y - 4 * (float)Math.Sin(a - .65));
                Line(g, c, x, y, x - 4 * (float)Math.Cos(a + .65), y - 4 * (float)Math.Sin(a + .65));
            };
            head(x2, y2, angle); if (both) head(x1, y1, angle + Math.PI);
        }
        private static void Axes(Graphics g) { Line(g, Ink, 4, 4, 4, 28); Line(g, Ink, 4, 28, 29, 28); }
        private static void Chart(Graphics g, string id)
        {
            if (id == "pie" || id == "doughnut")
            {
                using (var b = new SolidBrush(Blue)) g.FillPie(b, 4, 4, 24, 24, -90, 230);
                using (var b = new SolidBrush(Teal)) g.FillPie(b, 4, 4, 24, 24, 142, 126);
                if (id == "doughnut") Circle(g, Color.White, 11, 11, 10, true);
                return;
            }
            Axes(g);
            if (id == "scatter" || id == "bubble")
            {
                Circle(g, Blue, 7, 19, id == "bubble" ? 7 : 4, true);
                Circle(g, Teal, 16, 8, id == "bubble" ? 11 : 4, true);
                Circle(g, Amber, 23, 17, id == "bubble" ? 5 : 4, true); return;
            }
            if (id == "line" || id == "profile" || id == "area" || id == "combo")
            {
                var points = new[] { new PointF(7, 23), new PointF(14, 12), new PointF(21, 17), new PointF(28, 6) };
                if (id == "area") Polygon(g, Color.FromArgb(140, Teal), new PointF(7, 27), points[0], points[1], points[2], points[3], new PointF(28, 27));
                if (id == "combo") { Rect(g, Blue, 8, 17, 4, 10); Rect(g, Blue, 17, 13, 4, 14); Rect(g, Blue, 25, 10, 4, 17); }
                if (id == "profile") points = new[] { new PointF(8, 5), new PointF(23, 12), new PointF(14, 19), new PointF(25, 25) };
                for (int i = 1; i < points.Length; i++) Line(g, id == "combo" ? Amber : Teal, points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y);
                if (id != "area") foreach (var p in points) Circle(g, id == "combo" ? Amber : Teal, p.X - 1.5f, p.Y - 1.5f, 3, true);
                return;
            }
            if (id == "waterfall")
            {
                Rect(g, Blue, 7, 16, 5, 11); Rect(g, Teal, 14, 9, 5, 7); Rect(g, Amber, 21, 9, 5, 4);
                Line(g, Ink, 12, 16, 14, 16, 1); Line(g, Ink, 19, 9, 21, 9, 1); return;
            }
            if (id == "gantt") { Rect(g, Blue, 7, 7, 12, 4); Rect(g, Teal, 12, 15, 16, 4); Rect(g, Amber, 20, 23, 8, 4); return; }
            if (id == "mekko" || id == "mekko-units")
            {
                Rect(g, Blue, 7, 19, 4, 8); Rect(g, Teal, 7, 7, 4, 11); Rect(g, Blue, 13, 14, 7, 13); Rect(g, Teal, 13, 7, 7, 6);
                Rect(g, Blue, 22, 20, 7, 7); Rect(g, Teal, 22, id == "mekko" ? 7 : 12, 7, id == "mekko" ? 12 : 7); return;
            }
            if(id=="column"||id=="bar"){
                for(int i=0;i<3;i++){float a=9+i*4,b=15-i*3;if(id=="bar"){Rect(g,Blue,6,5+i*8,a,3);Rect(g,Teal,6,9+i*8,b,3);}else{Rect(g,Blue,7+i*8,27-a,3,a);Rect(g,Teal,11+i*8,27-b,3,b);}}return;
            }
            bool horizontal = id == "bar" || id == "stacked-bar" || id == "percent-bar";
            bool stacked = id.StartsWith("stacked-", StringComparison.Ordinal) || id.StartsWith("percent-", StringComparison.Ordinal);
            bool percent = id.StartsWith("percent-", StringComparison.Ordinal);
            for (int i = 0; i < 3; i++)
            {
                float total = percent ? 21 : 10 + i * 5;
                if (horizontal)
                {
                    Rect(g, Blue, 6, 5 + i * 8, stacked ? total * .55f : total, 5);
                    if (stacked) Rect(g, Teal, 6 + total * .55f, 5 + i * 8, total * .45f, 5);
                }
                else
                {
                    Rect(g, Blue, 7 + i * 8, 27 - (stacked ? total * .55f : total), 5, stacked ? total * .55f : total);
                    if (stacked) Rect(g, Teal, 7 + i * 8, 27 - total, 5, total * .45f);
                }
            }
            if (percent) Letter(g, "%", 18, -1, 10, Ink);
        }
        private static void Grid(Graphics g, bool header = false)
        {
            Rect(g, Color.White, 4, 5, 24, 22); if (header) Rect(g, Pale, 4, 5, 24, 7);
            Outline(g, Ink, 4, 5, 24, 22);
            Line(g, header ? Blue : Ink, 4, 12, 28, 12, 1.5f); Line(g, Ink, 4, 20, 28, 20, 1.5f);
            Line(g, Ink, 12, 5, 12, 27, 1.5f); Line(g, Ink, 20, 5, 20, 27, 1.5f);
        }
        private static void Document(Graphics g)
        { Rect(g, Color.White, 6, 3, 20, 26); Outline(g, Ink, 6, 3, 20, 26); Line(g, Ink, 10, 9, 22, 9); Line(g, Ink, 10, 14, 19, 14); }
        private static void Slides(Graphics g)
        { Rect(g, Pale, 3, 4, 20, 16); Outline(g, Ink, 3, 4, 20, 16); Rect(g, Color.White, 9, 11, 20, 16); Outline(g, Blue, 9, 11, 20, 16); }
        private static void Folder(Graphics g)
        { Polygon(g, Pale, new PointF(3, 7), new PointF(12, 7), new PointF(15, 11), new PointF(29, 11), new PointF(26, 27), new PointF(3, 27)); Line(g, Ink, 3, 7, 12, 7); Line(g, Ink, 12, 7, 15, 11); Outline(g, Blue, 3, 11, 26, 16); }
        private static void Palette(Graphics g)
        { Circle(g, Pale, 3, 3, 26, true); Circle(g, Ink, 3, 3, 26, false); Circle(g, Blue, 9, 7, 5, true); Circle(g, Teal, 19, 9, 5, true); Circle(g, Amber, 7, 17, 5, true); Circle(g, Color.White, 18, 20, 7, true); }
        private static void Refresh(Graphics g)
        { using (var p = new Pen(Teal, 2.4f)) g.DrawArc(p, 5, 5, 22, 22, 30, 280); Arrow(g, Teal, 26, 4, 26, 12); }
        private static void Magnify(Graphics g)
        { Circle(g, Color.White, 7, 5, 15, true); Circle(g, Blue, 7, 5, 15, false); Line(g, Blue, 20, 19, 28, 27, 3); }
        private static void Chain(Graphics g)
        { using (var p = new Pen(Blue, 3)) { g.DrawEllipse(p, 3, 11, 15, 10); g.DrawEllipse(p, 14, 11, 15, 10); } }
        private static void Align(Graphics g, string id)
        {
            bool vertical = id == "top" || id == "middle" || id == "bottom";
            float anchor = id == "right" || id == "bottom" ? 27 : id == "center" || id == "middle" ? 16 : 5;
            if (vertical)
            {
                Line(g, Amber, 3, anchor, 29, anchor);
                for (int i = 0; i < 2; i++) { float h = i == 0 ? 10 : 17; Rect(g, i == 0 ? Blue : Teal, 8 + i * 11, id == "top" ? anchor : id == "bottom" ? anchor - h : anchor - h / 2, 6, h); }
            }
            else
            {
                Line(g, Amber, anchor, 3, anchor, 29);
                for (int i = 0; i < 2; i++) { float w = i == 0 ? 10 : 17; Rect(g, i == 0 ? Blue : Teal, id == "left" ? anchor : id == "right" ? anchor - w : anchor - w / 2, 8 + i * 11, w, 6); }
            }
        }
        // Geometry stays inside a 2–30 logical-pixel safe area. At 16 px,
        // two logical pixels become one screen pixel; do not rely on tiny captions.
        private static void TextLines(Graphics g,string mode)
        {
            var widths=new[]{22f,14f,19f,11f};
            for(int i=0;i<4;i++){float w=mode=="justify"?22:widths[i];float x=mode=="right"?27-w:mode=="center"?16-w/2:5;Line(g,i==0?Blue:Ink,x,7+i*6,x+w,7+i*6,2.2f);}
        }
        private static void TextFrameIcon(Graphics g,string mode)
        {
            if(mode=="left"||mode=="center"||mode=="right"||mode=="justify"){TextLines(g,mode);return;}
            if(mode=="top"||mode=="middle"||mode=="bottom"){
                Outline(g,Ink,4,3,24,26);float y=mode=="top"?7:mode=="middle"?13:19;
                Line(g,Blue,9,y,23,y,2.3f);Line(g,Blue,9,y+5,20,y+5,2.3f);
                Line(g,Teal,8,mode=="top"?3:mode=="bottom"?29:16,24,mode=="top"?3:mode=="bottom"?29:16);return;
            }
            if(mode=="horizontal"){Letter(g,"A",3,2,23,Blue);Arrow(g,Teal,5,28,27,28);return;}
            if(mode=="downward"){g.TranslateTransform(32,0);g.RotateTransform(90);Letter(g,"A",3,2,23,Blue);Arrow(g,Teal,5,28,27,28);return;}
            g.TranslateTransform(0,32);g.RotateTransform(-90);Letter(g,"A",3,2,23,Blue);Arrow(g,Teal,5,28,27,28);
        }
        private static void Gallery(Graphics g)
        {
            Outline(g,Ink,3,3,26,26);Rect(g,Blue,4,4,24,4);
            for(int y=0;y<2;y++)for(int x=0;x<2;x++){Rect(g,Pale,7+x*11,11+y*9,8,6);Outline(g,x==0&&y==0?Blue:Teal,7+x*11,11+y*9,8,6);}
        }
        private static void Gear(Graphics g,float x,float y)
        {
            Circle(g,Teal,x,y,13,false);Circle(g,Teal,x+4,y+4,5,false);
            for(int i=0;i<8;i++){double angle=i*Math.PI/4;Line(g,Teal,x+6.5f+5*(float)Math.Cos(angle),y+6.5f+5*(float)Math.Sin(angle),x+6.5f+9*(float)Math.Cos(angle),y+6.5f+9*(float)Math.Sin(angle));}
        }
        private static bool WorkflowIcon(Graphics g,string id)
        {
            if(id.StartsWith("text-",StringComparison.Ordinal)){string mode=id.Substring(5);if(!new[]{"left","center","right","justify","top","middle","bottom","horizontal","downward","upward"}.Contains(mode))return false;TextFrameIcon(g,mode);return true;}
            switch(id){
                case "agenda-contents": Outline(g,Ink,3,3,26,26);for(int i=0;i<3;i++){Rect(g,Teal,7,8+i*7,3,3);Line(g,Blue,14,9+i*7,25,9+i*7);}break;
                case "agenda-settings": Outline(g,Ink,3,3,22,26);Line(g,Blue,7,8,19,8);Line(g,Blue,7,14,15,14);Gear(g,15,16);break;
                case "agenda-split": for(int i=0;i<2;i++){Line(g,Blue,4,6+i*6,12,6+i*6);Line(g,Teal,20,20+i*6,28,20+i*6);}Line(g,Amber,3,25,28,3,2.5f);break;
                case "agenda-insert": Outline(g,Ink,3,3,22,26);for(int i=0;i<3;i++){Rect(g,Blue,7,8+i*7,3,3);Line(g,Ink,14,9+i*7,21,9+i*7);}Circle(g,Teal,19,18,12,true);Line(g,Color.White,22,24,28,24);Line(g,Color.White,25,21,25,27);break;
                case "agenda-refresh": Outline(g,Ink,3,3,22,26);for(int i=0;i<3;i++)Line(g,Blue,7,8+i*6,19,8+i*6);using(var pen=new Pen(Teal,2.5f))g.DrawArc(pen,17,17,13,13,25,280);Arrow(g,Teal,29,17,29,22);break;
                case "annotation-refresh": Chart(g,"column");Line(g,Amber,6,10,28,10);Circle(g,Color.White,18,17,13,true);using(var pen=new Pen(Teal,2.3f))g.DrawArc(pen,19,18,11,11,25,280);Arrow(g,Teal,29,18,29,23);break;
                case "percent-area": Axes(g);Polygon(g,Blue,new PointF(7,27),new PointF(7,17),new PointF(17,12),new PointF(28,19),new PointF(28,27));Polygon(g,Teal,new PointF(7,15),new PointF(7,5),new PointF(28,5),new PointF(28,17),new PointF(17,10));break;
                case "axis-settings": Axes(g);Line(g,Blue,4,11,16,11);Line(g,Blue,4,20,12,20);Gear(g,15,13);break;
                case "match-scales": Line(g,Ink,3,3,3,28);Line(g,Ink,29,3,29,28);Rect(g,Blue,6,13,5,14);Rect(g,Teal,21,7,5,20);Line(g,Amber,2,5,30,5);Line(g,Amber,2,28,30,28);Arrow(g,Teal,12,16,20,16,true);break;
                case "series-settings": Rect(g,Blue,4,17,5,11);Rect(g,Blue,13,11,5,17);Line(g,Teal,5,13,15,5);Line(g,Teal,15,5,27,10);Gear(g,15,16);break;
                case "connections": Line(g,Ink,6,16,26,16);Line(g,Ink,16,6,16,26);Circle(g,Blue,11,11,10,true);Circle(g,Teal,2,12,7,true);Circle(g,Teal,23,12,7,true);Circle(g,Teal,12,2,7,true);Circle(g,Teal,12,23,7,true);break;
                case "map-excel": Grid(g,true);using(var pen=new Pen(Teal,3))g.DrawRectangle(pen,11,12,17,15);Arrow(g,Amber,2,29,10,21);break;
                case "relink-selected": Chain(g);Line(g,Color.White,12,10,19,22,4);Arrow(g,Teal,5,5,26,5);Arrow(g,Teal,26,27,5,27);break;
                case "data-manager": Grid(g,true);Circle(g,Color.White,17,16,13,true);Circle(g,Teal,17,16,11,false);Line(g,Teal,26,25,30,30,2.5f);break;
                case "color-fill": Polygon(g,Blue,new PointF(5,12),new PointF(15,3),new PointF(26,14),new PointF(15,25));Line(g,Color.White,7,13,24,13);Polygon(g,Teal,new PointF(27,18),new PointF(23,26),new PointF(27,29),new PointF(31,26));Line(g,Ink,3,29,19,29);break;
                case "color-outline": Outline(g,Blue,4,4,23,23);Outline(g,Color.White,7,7,17,17);Line(g,Teal,18,29,29,18,4);break;
                case "color-text": Letter(g,"A",6,0,27,Ink);Rect(g,Blue,4,27,24,4);break;
                case "quality-review": Outline(g,Ink,3,3,23,26);for(int i=0;i<3;i++){Line(g,Teal,6,8+i*7,8,10+i*7);Line(g,Teal,8,10+i*7,11,6+i*7);Line(g,Ink,15,8+i*7,22,8+i*7);}Circle(g,Color.White,19,19,10,true);Circle(g,Blue,19,19,8,false);Line(g,Blue,26,26,30,30);break;
                case "export-pdf": Document(g);Polygon(g,Blue,new PointF(9,21),new PointF(14,10),new PointF(19,21),new PointF(9,21),new PointF(24,17));Arrow(g,Teal,18,27,30,27);break;
                case "template-palette": Gallery(g);Circle(g,Color.White,15,15,16,true);Circle(g,Blue,18,18,5,true);Circle(g,Teal,25,19,5,true);Circle(g,Amber,20,25,5,true);break;
                case "template-alignment": Outline(g,Ink,3,3,26,26);Line(g,Teal,16,7,16,26);Line(g,Blue,8,10,24,10);Line(g,Blue,11,16,21,16);Line(g,Blue,9,22,23,22);break;
                case "layout-links": Outline(g,Blue,3,4,10,9);Outline(g,Teal,19,19,10,9);Line(g,Ink,13,8,24,8);Line(g,Ink,24,8,24,18);Circle(g,Amber,21,5,6,true);break;
                case "editable-columns": Rect(g,Blue,5,16,6,11);Rect(g,Teal,14,9,6,18);Rect(g,Blue,23,4,6,23);foreach(var point in new[]{new PointF(2,13),new PointF(10,13),new PointF(2,26),new PointF(10,26)}){Rect(g,Color.White,point.X,point.Y,4,4);Outline(g,Ink,point.X,point.Y,4,4);}break;
                default:return false;
            }return true;
        }
        private static bool Draw(Graphics g, string id)
        {
            if(WorkflowIcon(g,id))return true;
            switch (id)
            {
                case "column": case "bar": case "stacked-column": case "stacked-bar": case "percent-column": case "percent-bar":
                case "line": case "profile": case "area": case "pie": case "doughnut": case "scatter": case "bubble": case "combo":
                case "waterfall": case "mekko": case "mekko-units": case "gantt": Chart(g, id); break;
                case "managed-column": Chart(g, "column"); Circle(g, Teal, 21, 1, 10, true); Line(g, Color.White, 23, 6, 25, 8); Line(g, Color.White, 25, 8, 29, 3); break;
                case "charts-gallery":
                    Axes(g);Rect(g,Blue,8,18,5,8);Rect(g,Teal,16,12,5,14);Rect(g,Blue,24,20,5,6);
                    Line(g,Amber,7,13,13,7);Line(g,Amber,13,7,18,9);Circle(g,Pale,22,2,9,true);using(var sector=new SolidBrush(Teal))g.FillPie(sector,22,2,9,9,-90,240);break;
                case "category-charts": Chart(g, "combo"); break;
                case "left": case "center": case "right": case "top": case "middle": case "bottom": Align(g, id); break;
                case "width": case "height": case "size":
                    Outline(g, Blue, 8, 8, 17, 17);
                    if (id != "height") Arrow(g, Teal, 7, 4, 26, 4, true);
                    if (id != "width") Arrow(g, Teal, 4, 7, 4, 26, true); break;
                case "horizontal": case "vertical":
                    for (int i = 0; i < 3; i++) Rect(g, i == 1 ? Teal : Blue, id == "horizontal" ? 4 + i * 10 : 11, id == "vertical" ? 4 + i * 10 : 11, id == "horizontal" ? 4 : 10, id == "vertical" ? 4 : 10);
                    if (id == "horizontal") { Arrow(g, Ink, 4, 25, 28, 25, true); } else Arrow(g, Ink, 25, 4, 25, 28, true); break;
                case "swap": Rect(g, Blue, 3, 12, 9, 9); Rect(g, Teal, 20, 12, 9, 9); Arrow(g, Ink, 7, 7, 25, 7); Arrow(g, Ink, 25, 26, 7, 26); break;
                case "scale-up": case "scale-down": Outline(g, Blue, 7, 7, 18, 18); Line(g, Teal, 11, 16, 21, 16); if (id == "scale-up") Line(g, Teal, 16, 11, 16, 21); break;
                case "rotate": Outline(g, Blue, 10, 10, 12, 12); using (var p = new Pen(Teal, 2)) g.DrawArc(p, 3, 3, 26, 26, 45, 250); Arrow(g, Teal, 21, 3, 28, 9); break;
                case "flip-horizontal": case "flip-vertical":
                    if (id == "flip-vertical") { g.TranslateTransform(32, 0); g.RotateTransform(90); }
                    Polygon(g, Blue, new PointF(4, 8), new PointF(12, 15), new PointF(4, 24)); Polygon(g, Teal, new PointF(28, 8), new PointF(20, 15), new PointF(28, 24)); Line(g, Ink, 16, 4, 16, 28); break;
                case "group": case "ungroup":
                    Outline(g, Blue, 6, 9, 13, 13); Outline(g, Teal, 15, 13, 11, 11);
                    if (id == "group") { Outline(g, Ink, 3, 5, 26, 23); foreach (var x in new[] { 2, 27 }) foreach (var y in new[] { 4, 27 }) Rect(g, Blue, x, y, 3, 3); }
                    else { Arrow(g, Ink, 12, 7, 6, 3); Arrow(g, Ink, 23, 25, 29, 29); } break;
                case "front": case "back":
                    Rect(g, id == "front" ? Pale : Blue, 5, 5, 16, 16); Outline(g, Ink, 5, 5, 16, 16); Rect(g, id == "front" ? Blue : Pale, 12, 12, 16, 16); Outline(g, Ink, 12, 12, 16, 16); break;
                case "duplicate": case "duplicate-slide": Slides(g); if (id == "duplicate") { Line(g, Teal, 17, 19, 23, 19); Line(g, Teal, 20, 16, 20, 22); } break;
                case "auto-layout": case "category-arrange": case "mini-toolbar":
                    for (int row = 0; row < 2; row++) for (int col = 0; col < 3; col++) Rect(g, col == 1 ? Teal : Blue, 4 + col * 9, 7 + row * 11, 6, 7);
                    if (id == "mini-toolbar") Line(g, Ink, 4, 29, 28, 29); break;
                case "pentagon":Polygon(g,Blue,new PointF(3,7),new PointF(22,7),new PointF(29,16),new PointF(22,25),new PointF(3,25));break;
                case "chevron":Polygon(g,Teal,new PointF(3,7),new PointF(20,7),new PointF(29,16),new PointF(20,25),new PointF(3,25),new PointF(12,16));break;
                case "universal-connector":Outline(g,Blue,2,3,8,8);Outline(g,Teal,22,21,8,8);Line(g,Ink,10,7,16,7);Line(g,Ink,16,7,16,25);Arrow(g,Ink,16,25,22,25);break;
                case "element-edit":Outline(g,Ink,3,4,20,24);Line(g,Blue,9,22,27,4,4);Line(g,Teal,8,24,13,22);break;
                case "text": Outline(g,Ink,3,4,26,24);Letter(g,"A",8,4,21,Blue);break;
                case "font": case "category-text": Letter(g, "A", 5, 1, 28, Blue); Line(g, Teal, 5, 29, 27, 29, 3); break;
                case "font-size": Letter(g, "A", 2, 7, 24, Blue); Arrow(g, Teal, 26, 5, 26, 27, true); break;
                case "paste-text": Document(g); Arrow(g, Teal, 2, 22, 15, 22); break;
                case "decimal": Circle(g,Blue,5,8,6,true);Circle(g,Teal,21,8,6,true);Line(g,Teal,25,13,21,19,2.5f);Arrow(g,Ink,5,26,27,26,true);break;
                case "language": Circle(g, Ink, 4, 4, 24, false); using (var p = new Pen(Blue, 1.5f)) g.DrawEllipse(p, 11, 4, 10, 24); Line(g, Teal, 4, 16, 28, 16); break;
                case "rounded": using (var path = new GraphicsPath()) { path.AddArc(4, 7, 8, 8, 180, 90); path.AddArc(20, 7, 8, 8, 270, 90); path.AddArc(20, 17, 8, 8, 0, 90); path.AddArc(4, 17, 8, 8, 90, 90); path.CloseFigure(); using (var b = new SolidBrush(Blue)) g.FillPath(b, path); } break;
                case "process": case "category-elements":
                    for (int i = 0; i < 3; i++) Polygon(g, i == 1 ? Teal : Blue, new PointF(2 + i * 10, 9), new PointF(8 + i * 10, 9), new PointF(12 + i * 10, 16), new PointF(8 + i * 10, 23), new PointF(2 + i * 10, 23), new PointF(6 + i * 10, 16)); break;
                case "checkbox": Outline(g, Blue, 5, 5, 22, 22); Line(g, Teal, 9, 16, 14, 21, 3); Line(g, Teal, 14, 21, 24, 10, 3); break;
                case "harvey": Circle(g, Color.White, 4, 4, 24, true); using (var b = new SolidBrush(Blue)) g.FillPie(b, 4, 4, 24, 24, -90, 180); Circle(g, Ink, 4, 4, 24, false); break;
                case "symbol": Letter(g, "Ω", 3, 1, 28, Teal); break;
                case "table": case "table-data": case "edit-data": case "sheet-style": case "category-data":
                    Grid(g, id != "table"); if (id == "edit-data" || id == "table-data") { Line(g, Amber, 19, 27, 29, 17, 4); } break;
                case "agenda": case "agenda-layout":
                    for (int i = 0; i < 3; i++) { Circle(g, Teal, 4, 6 + i * 8, 4, true); Line(g, Blue, 12, 8 + i * 8, 27, 8 + i * 8); }
                    if (id == "agenda-layout") Outline(g, Ink, 1, 2, 30, 28); break;
                case "image": Rect(g, Color.White, 3, 5, 26, 22); Outline(g, Ink, 3, 5, 26, 22); Circle(g, Amber, 19, 8, 5, true); Polygon(g, Teal, new PointF(5, 25), new PointF(13, 13), new PointF(21, 25)); Polygon(g, Blue, new PointF(16, 25), new PointF(23, 16), new PointF(28, 25)); break;
                case "library-browser": Gallery(g);break;
                case "library": case "category-assets": Folder(g);break;
                case "template": Document(g); Rect(g, Blue, 10, 18, 5, 7); Rect(g, Teal, 18, 18, 5, 7); break;
                case "theme": case "chart-style": case "style-file": Palette(g); if (id == "style-file") { Line(g, Ink, 23, 24, 30, 17, 3); } break;
                case "path": Document(g); Letter(g, "/", 15, 16, 14, Teal); break;
                case "export-slide": case "save-slides": Slides(g); Arrow(g, Teal, 16, 20, 29, 20); if (id == "save-slides") Rect(g, Blue, 13, 14, 6, 5); break;
                case "clean-comments": Outline(g, Ink, 3, 5, 25, 18); Line(g, Ink, 8, 23, 8, 28); Line(g, Ink, 8, 28, 15, 23); Line(g, Amber, 11, 10, 20, 19); Line(g, Amber, 20, 10, 11, 19); break;
                case "legend": for (int i = 0; i < 3; i++) { Rect(g, i == 1 ? Teal : Blue, 4, 6 + i * 8, 5, 5); Line(g, Ink, 13, 8 + i * 8, 27, 8 + i * 8); } break;
                case "labels": Chart(g, "column"); Rect(g, Color.White, 9, 1, 13, 10); Letter(g, "12", 9, 0, 10, Ink); break;
                case "data-table": Chart(g, "column"); Rect(g, Color.White, 2, 20, 28, 11); Outline(g, Ink, 2, 20, 28, 11); Line(g, Ink, 2, 25, 30, 25, 1); Line(g, Ink, 12, 20, 12, 31, 1); Line(g, Ink, 22, 20, 22, 31, 1); break;
                case "value-line": Chart(g, "column"); Line(g, Amber, 2, 13, 30, 13, 2.5f); break;
                case "difference": Outline(g, Blue, 4, 16, 8, 12); Outline(g, Teal, 20, 7, 8, 21); Arrow(g, Amber, 16, 8, 16, 26, true); break;
                case "cagr": case "excel-cagr": Arrow(g, Teal, 4, 26, 28, 5); Letter(g, id == "cagr" ? "%" : "fx", 3, 1, 15, Blue); break;
                case "axes": Axes(g); Arrow(g, Teal, 4, 27, 4, 3); Letter(g, "Y", 11, 2, 10, Blue); Letter(g, "X", 21, 17, 10, Blue); break;
                case "plot-area": Rect(g, Pale, 7, 4, 23, 21); Chart(g, "column"); break;
                case "chart-title": Chart(g, "column"); Rect(g, Color.White, 6, 0, 24, 9); Letter(g, "Title", 7, -1, 9, Ink); break;
                case "scan": Chart(g, "column"); Magnify(g); break;
                case "round-data": Letter(g, ".0", 2, 2, 23, Blue); Arrow(g, Teal, 4, 28, 28, 28); break;
                case "import-json": case "export-json": case "category-automation": Letter(g, "{ }", 2, 1, 23, Blue); if (id != "category-automation") Arrow(g, Teal, id == "import-json" ? 4 : 28, 28, id == "import-json" ? 28 : 4, 28); break;
                case "source-excel": Grid(g,true);Rect(g,Teal,2,11,16,18);Line(g,Color.White,6,15,14,25,2.3f);Line(g,Color.White,14,15,6,25,2.3f);break;
                case "source-csv": Grid(g,true);Rect(g,Color.White,15,16,16,15);Circle(g,Teal,17,18,4,true);Line(g,Teal,20,22,18,26);Circle(g,Teal,25,18,4,true);Line(g,Teal,28,22,26,26);break;
                case "source-json": Document(g);Rect(g,Color.White,8,8,16,18);Letter(g,"{ }",6,8,17,Blue);break;
                case "source-xml": Document(g);Rect(g,Color.White,7,9,18,15);Line(g,Teal,12,12,8,17);Line(g,Teal,8,17,12,22);Line(g,Teal,20,12,24,17);Line(g,Teal,24,17,20,22);break;
                case "link-excel": case "manage-links": Chain(g); if (id == "link-excel") { Rect(g, Color.White, 1, 1, 12, 10); Letter(g, "X", 3, -1, 11, Teal); } break;
                case "refresh-excel": Refresh(g); Letter(g, "X", 10, 9, 13, Teal); break;
                case "slide-workbook": Slides(g); Letter(g, "X", 15, 12, 13, Teal); break;
                case "check-sync": Refresh(g); Line(g, Blue, 10, 16, 14, 20, 3); Line(g, Blue, 14, 20, 23, 10, 3); break;
                case "group-more": for(int row=0;row<3;row++)for(int col=0;col<3;col++)Rect(g,row==1&&col==1?Teal:Blue,5+col*9,5+row*9,5,5);break;
                case "workspace": Outline(g, Ink, 3, 5, 26, 23); Rect(g, Blue, 4, 6, 24, 5); Line(g, Ink, 19, 12, 19, 27); Rect(g, Pale, 21, 14, 6, 11); break;
                case "explore": Magnify(g); Circle(g, Teal, 11, 9, 7, true); break;
                case "compatibility": Outline(g, Blue, 5, 4, 22, 24); Letter(g, "16", 8, 5, 12, Ink); Line(g, Teal, 10, 21, 14, 25); Line(g, Teal, 14, 25, 23, 17); break;
                case "demo-tour": case "category-slides": Slides(g); Polygon(g, Teal, new PointF(15, 14), new PointF(24, 19), new PointF(15, 24)); break;
                case "category-chartoptions": Palette(g); break;
                default: return false;
            }
            return true;
        }
    }
}
