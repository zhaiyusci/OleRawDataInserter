namespace OLEPackager.Core
{
    public sealed class OfficeShapeState
    {
        public float Width { get; set; }

        public float Height { get; set; }

        public float Left { get; set; }

        public float Top { get; set; }

        public float Rotation { get; set; }

        public int Visible { get; set; }

        public int HorizontalFlip { get; set; }

        public int VerticalFlip { get; set; }

        public int LockAspectRatio { get; set; }

        public int ZOrderPosition { get; set; }

        public string Name { get; set; }

        public string AlternativeText { get; set; }

        public string Title { get; set; }

        public OfficePictureState Picture { get; set; }

        public OfficeLineState Line { get; set; }

        public OfficeShadowState Shadow { get; set; }
    }

    public sealed class OfficePictureState
    {
        public float Brightness { get; set; }

        public float Contrast { get; set; }

        public float CropBottom { get; set; }

        public float CropLeft { get; set; }

        public float CropRight { get; set; }

        public float CropTop { get; set; }

        public int ColorType { get; set; }

        public int TransparencyColor { get; set; }

        public int TransparentBackground { get; set; }
    }

    public sealed class OfficeLineState
    {
        public int Visible { get; set; }

        public int ForeColorRgb { get; set; }

        public int BackColorRgb { get; set; }

        public float Transparency { get; set; }

        public float Weight { get; set; }

        public int DashStyle { get; set; }

        public int Style { get; set; }

        public int InsetPen { get; set; }

        public int Pattern { get; set; }

        public int BeginArrowheadLength { get; set; }

        public int BeginArrowheadStyle { get; set; }

        public int BeginArrowheadWidth { get; set; }

        public int EndArrowheadLength { get; set; }

        public int EndArrowheadStyle { get; set; }

        public int EndArrowheadWidth { get; set; }
    }

    public sealed class OfficeShadowState
    {
        public int Visible { get; set; }

        public int ForeColorRgb { get; set; }

        public float Transparency { get; set; }

        public int Type { get; set; }

        public int Style { get; set; }

        public float OffsetX { get; set; }

        public float OffsetY { get; set; }

        public float Blur { get; set; }

        public float Size { get; set; }

        public int RotateWithShape { get; set; }

        public int Obscured { get; set; }
    }
}
