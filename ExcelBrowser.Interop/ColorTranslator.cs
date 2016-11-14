using System.Windows.Media;

namespace ExcelBrowser.Interop {

    public static class ColorTranslator {

        const int RedShift = 0;
        const int GreenShift = 8;
        const int BlueShift = 16;

        /// <summary>
        /// Translates an Ole color value to a System.Media.Color for WPF usage
        /// </summary>
        /// <param name="oleColor">Ole int32 color value</param>
        /// <returns>System.Media.Color color value</returns>
        public static Color FromOle(this int oleColor) {
            return Color.FromRgb(
                (byte)((oleColor >> RedShift) & 0xFF),
                (byte)((oleColor >> GreenShift) & 0xFF),
                (byte)((oleColor >> BlueShift) & 0xFF)
                );
        }

        /// <summary>
        /// Translates the specified System.Media.Color to an Ole color.
        /// </summary>
        /// <param name="wpfColor">System.Media.Color source value</param>
        /// <returns>Ole int32 color value</returns>
        public static int ToOle(Color wpfColor) {
            return wpfColor.R << RedShift 
                | wpfColor.G << GreenShift 
                | wpfColor.B << BlueShift;
        }
    }
}
