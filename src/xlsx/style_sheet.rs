use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};
use crate::{ar_deserable, enum_default};
use super::base::{XlsxResult, ArchiveDeserable};

pub struct StyleSheet {
    sheet: StyleSheetItems,
}

ar_deserable!(StyleSheet, "xl/styles.xml", sheet: StyleSheetItems);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "styleSheet",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
pub struct StyleSheetItems {
    #[yaserde(rename = "numFmts")]
    num_fmts: NumFmts,

    fonts: Fonts,

    fills: Fills,

    borders: Borders,

    #[yaserde(rename = "cellStyleXfs")]
    cell_style_xfs: CellStyleXfs,

    #[yaserde(rename = "cellXfs")]
    cell_xfs: CellXfs,

    #[yaserde(rename = "cellStyles")]
    cell_styles: CellStyles,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct NumFmts {
    #[yaserde(attribute, rename = "count")]
    count: usize,

    #[yaserde(rename = "numFmt")]
    items: Vec<NumFmt>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct NumFmt {
    #[yaserde(attribute, rename = "numFmtId")]
    id: String,

    #[yaserde(attribute, rename = "formatCode")]
    format_code: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct Fonts {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(rename = "font")]
    items: Vec<Font>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "font",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct Font {
    #[yaserde(rename = _)]
    items: Vec<FontProperty>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
enum FontProperty {
    #[yaserde(rename = "b")]
    Bold,

    #[yaserde(rename = "sz")]
    Size { #[yaserde(attribute)] val: String },

    #[yaserde(rename = "name")]
    Name { #[yaserde(attribute)] val: String },

    #[yaserde(rename = "color")]
    Color { #[yaserde(attribute)] theme: String },

    #[yaserde(rename = "family")]
    FontFamily { #[yaserde(attribute)] val: String },

    #[yaserde(rename = "charset")]
    Charset { #[yaserde(attribute)] val: String },

    Unknown,
}

enum_default!(FontProperty::Unknown);

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct Fills {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(rename = "fill")]
    items: Vec<Fill>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "fill",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct Fill {
    #[yaserde(rename = _)]
    items: FillType,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
enum FillType {
    #[yaserde(rename = "patternFill")]
    PatternFill {
        #[yaserde(attribute, rename = "patternType")]
        pattern_type: String,

        #[yaserde(rename = "fgColor")]
        fg_color: Option<Color>,

        #[yaserde(rename = "bgColor")]
        bg_color: Option<Color>,
    },

    #[yaserde(rename = "gradientFill")]
    GradientFill {
        #[yaserde(attribute)]
        degree: Option<String>,

        #[yaserde(attribute, rename = "type")]
        fill_type: Option<String>,

        #[yaserde(attribute)]
        left: Option<String>,

        #[yaserde(attribute)]
        right: Option<String>,

        #[yaserde(attribute)]
        top: Option<String>,

        #[yaserde(attribute)]
        bottom: Option<String>,

        #[yaserde(rename = "stop")]
        stops: Vec<Stop>,
    },

    Unknown,
}

enum_default!(FillType::Unknown);

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct Color {
    #[yaserde(attribute)]
    rgb: Option<String>,

    #[yaserde(attribute)]
    indexed: Option<String>,

    #[yaserde(attribute)]
    theme: Option<String>,

    #[yaserde(attribute)]
    auto: Option<String>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
struct Stop {
    #[yaserde(attribute)]
    position: String,

    color: Color,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct Borders {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(rename = "border")]
    items: Vec<Border>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    rename = "border",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct Border {
    #[yaserde(attribute, rename = "diagonalDown")]
    diagonal_down: Option<String>,

    #[yaserde(attribute, rename = "diagonalUp")]
    diagonal_up: Option<String>,

    left: Option<Side>,
    right: Option<Side>,
    top: Option<Side>,
    bottom: Option<Side>,
    diagonal: Option<Side>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct Side {
    #[yaserde(attribute)]
    style: Option<String>,

    color: Option<Color>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct CellStyleXfs {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(rename = "xf")]
    items: Vec<Xf>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct CellXfs {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(rename = "xf")]
    items: Vec<Xf>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "xf")]
struct Xf {
    #[yaserde(attribute, rename = "numFmtId")]
    numfmt_id: usize,

    #[yaserde(attribute, rename = "fontId")]
    font_id: usize,

    #[yaserde(attribute, rename = "fillId")]
    fill_id: usize,

    #[yaserde(attribute, rename = "borderId")]
    border_id: usize,

    #[yaserde(attribute, rename = "xfId")]
    xf_id: Option<usize>,

    #[yaserde(attribute, rename = "applyAlignment")]
    apply_alignment: Option<String>,

    #[yaserde(attribute, rename = "applyBorder")]
    apply_border: Option<String>,

    #[yaserde(attribute, rename = "applyFill")]
    apply_fill: Option<String>,

    #[yaserde(attribute, rename = "applyNumberFormat")]
    apply_numfmt: Option<String>,

    #[yaserde(attribute, rename = "applyProtection")]
    apply_protection: Option<String>,

    alignment: Option<Alignment>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct Alignment {
    #[yaserde(attribute)]
    horizontal: Option<String>,

    #[yaserde(attribute)]
    vertical: Option<String>,

    #[yaserde(attribute)]
    indent: Option<String>,

    #[yaserde(attribute, rename = "shrinkToFit")]
    shrink_to_fit: Option<String>,

    #[yaserde(attribute, rename = "textRotation")]
    text_rotation: Option<String>,

    #[yaserde(attribute, rename = "wrapText")]
    wrap_text: Option<String>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    rename = "cellStyles",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
)]
struct CellStyles {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(rename = "cellStyle")]
    items: Vec<CellStyle>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "cellStyle")]
struct CellStyle {
    #[yaserde(attribute)]
    name: String,

    #[yaserde(attribute, rename = "xfId")]
    xf_id: usize,

    #[yaserde(attribute, rename = "builtinId")]
    builtin_id: Option<usize>,
}

#[test]
fn test_load_ar() -> XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", StyleSheet::archive_string(&mut ar)?);

    let style_sheet = StyleSheet::from_archive(&mut ar)?;
    println!("{:?}\n", style_sheet.sheet);

    println!("{}\n", style_sheet.to_string()?);

    Ok(())
}