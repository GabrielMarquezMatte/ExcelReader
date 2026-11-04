//! `#[derive(ExcelMapper)]` for excelreader's `ExcelMapper` trait. Do not depend on this crate
//! directly - `excelreader` re-exports the macro as `excelreader::ExcelMapper`
//! (same identifier as the trait; derive macros and traits occupy different namespaces, so
//! there's no collision, same as `serde`/`serde_derive`).

use proc_macro::TokenStream;
use quote::quote;
use syn::spanned::Spanned;
use syn::{
    parse_macro_input, Data, DeriveInput, Field, Fields, GenericArgument, LitStr, PathArguments,
    Type,
};

#[proc_macro_derive(ExcelMapper, attributes(excel))]
pub fn derive_excel_mapper(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    expand(input)
        .unwrap_or_else(|e| e.to_compile_error())
        .into()
}

fn expand(input: DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let struct_name = &input.ident;
    let fields = named_fields(&input)?;

    let parsed = fields
        .iter()
        .map(parse_field)
        .collect::<syn::Result<Vec<_>>>()?;

    let bindings = parsed.iter().map(field_binding).collect::<Vec<_>>();

    let mut declarations = Vec::new();
    let mut pushes = Vec::new();
    let mut columns = Vec::new();
    for (index, field) in parsed.iter().enumerate() {
        let plan = WriteField::new(index, field);
        declarations.push(plan.declarations());
        pushes.push(plan.push());
        columns.push(plan.column());
    }

    let (impl_generics, ty_generics, where_clause) = input.generics.split_for_impl();

    Ok(quote! {
        impl #impl_generics ::excelreader::ExcelMapper for #struct_name #ty_generics #where_clause {
            fn bindings() -> ::std::vec::Vec<::excelreader::__private::ColumnBinding<Self>> {
                ::std::vec![ #(#bindings),* ]
            }
        }

        impl #impl_generics ::excelreader::ExcelWriter for #struct_name #ty_generics #where_clause {
            fn to_columns(
                rows: &[Self],
            ) -> ::core::result::Result<
                ::std::vec::Vec<::excelreader::__private::OwnedColumn>,
                ::excelreader::Error,
            > {
                let n = rows.len();
                #(#declarations)*
                for (row, r) in rows.iter().enumerate() {
                    let _ = row;
                    #(#pushes)*
                }
                ::core::result::Result::Ok(::std::vec![ #(#columns),* ])
            }
        }
    })
}

fn named_fields(
    input: &DeriveInput,
) -> syn::Result<&syn::punctuated::Punctuated<Field, syn::Token![,]>> {
    if let Data::Struct(syn::DataStruct {
        fields: Fields::Named(fields),
        ..
    }) = &input.data
    {
        return Ok(&fields.named);
    }
    Err(syn::Error::new_spanned(
        input,
        "ExcelMapper can only be derived for structs with named fields",
    ))
}

/// Everything both halves of the derive need from one field, resolved once.
///
/// The read half (`field_binding`) and the write half (`WriteField`) used to each re-run
/// `excel_names`/`unwrap_option`/`FieldKind::from_type` over the same field. Parsing once is not
/// just less work: it is what guarantees the two halves cannot disagree about a field's kind or
/// about which of its names is the primary one.
struct ParsedField<'a> {
    ident: &'a syn::Ident,
    /// Primary `#[excel(name = "...")]` first, then any aliases in declared order.
    names: Vec<LitStr>,
    /// The `CellStyle` expression `#[excel(style(...))]` builds, if present.
    style: Option<proc_macro2::TokenStream>,
    kind: FieldKind,
    ty: &'a Type,
    is_option: bool,
}

fn parse_field(field: &Field) -> syn::Result<ParsedField<'_>> {
    let (inner_ty, is_option) = unwrap_option(&field.ty);
    let (names, style) = excel_attributes(field)?;
    Ok(ParsedField {
        ident: field.ident.as_ref().expect("named_fields guarantees Some"),
        names,
        style,
        kind: FieldKind::from_type(inner_ty)?,
        ty: inner_ty,
        is_option,
    })
}

fn field_binding(parsed: &ParsedField<'_>) -> proc_macro2::TokenStream {
    let field_ident = parsed.ident;
    let name = &parsed.names;
    let column_type = parsed.kind.column_type_tokens();
    let value = parsed.kind.value_tokens();
    let assign_value = if parsed.is_option {
        quote! { ::std::option::Option::Some(#value) }
    } else {
        value
    };

    let ty = parsed.ty;
    let fits = match parsed.kind {
        FieldKind::Int => quote! {
            ::core::option::Option::Some(|v| <#ty as ::core::convert::TryFrom<i64>>::try_from(v).is_ok())
        },
        _ => quote! { ::core::option::Option::None },
    };

    quote! {
        ::excelreader::__private::ColumnBinding {
            names: &[#(#name),*],
            column_type: #column_type,
            assign: |r, col, row| r.#field_ident = #assign_value,
            fits: #fits,
        }
    }
}

fn excel_attributes(field: &Field) -> syn::Result<(Vec<LitStr>, Option<proc_macro2::TokenStream>)> {
    let mut name = None;
    let mut aliases = Vec::new();
    let mut style = None;
    for attr in &field.attrs {
        if !attr.path().is_ident("excel") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("name") {
                name = Some(meta.value()?.parse::<LitStr>()?);
                return Ok(());
            }
            if meta.path.is_ident("alias") {
                aliases.push(meta.value()?.parse::<LitStr>()?);
                return Ok(());
            }
            if meta.path.is_ident("style") {
                style = Some(parse_style(&meta)?);
                return Ok(());
            }
            Err(meta.error("unsupported #[excel(...)] key, expected `name`, `alias` or `style`"))
        })?;
    }
    let Some(name) = name else {
        return Err(syn::Error::new(
            field.span(),
            "field is missing #[excel(name = \"...\")]",
        ));
    };
    let mut names = vec![name];
    names.extend(aliases);
    Ok((names, style))
}

/// `style(...)`'s keys and values are Core's `[ExcelStyle]`, checked here against the same limits
/// the native library enforces, so a bad value fails the build instead of the write.
fn parse_style(meta: &syn::meta::ParseNestedMeta<'_>) -> syn::Result<proc_macro2::TokenStream> {
    let mut setters = Vec::new();
    let mut border = None;
    let mut border_color = None;
    meta.parse_nested_meta(|key| {
        let Some(ident) = key.path.get_ident() else {
            return Err(key.error(UNKNOWN_STYLE_KEY));
        };
        match ident.to_string().as_str() {
            flag @ ("bold" | "italic" | "underline" | "strikethrough") => {
                let setter = quote::format_ident!("{flag}");
                setters.push(quote! { .#setter(true) });
            }
            "wrap_text" => setters.push(quote! { .wrap_text(true) }),
            "number_format" => {
                let value: LitStr = key.value()?.parse()?;
                if value.value().encode_utf16().count() > 255 {
                    return Err(syn::Error::new(
                        value.span(),
                        "number_format cannot be longer than 255 characters",
                    ));
                }
                setters.push(quote! { .number_format(#value) });
            }
            "font_name" => {
                let value: LitStr = key.value()?.parse()?;
                if !(1..=31).contains(&value.value().encode_utf16().count()) {
                    return Err(syn::Error::new(
                        value.span(),
                        "font_name must be 1 to 31 characters",
                    ));
                }
                setters.push(quote! { .font_name(#value) });
            }
            "font_size" => {
                let (value, span) = number(&key)?;
                if value != 0.0 && !(1.0..=409.0).contains(&value) {
                    return Err(syn::Error::new(
                        span,
                        "font_size must be between 1 and 409 points",
                    ));
                }
                setters.push(quote! { .font_size(#value) });
            }
            "font_color" => {
                let value = color(&key)?;
                setters.push(quote! { .font_color(#value) });
            }
            "fill_color" => {
                let value = color(&key)?;
                setters.push(quote! { .fill_color(#value) });
            }
            "horizontal_alignment" => {
                let variant = keyword(&key, "horizontal_alignment", H_ALIGN)?;
                setters.push(quote! {
                    .horizontal_alignment(::excelreader::HorizontalAlignment::#variant)
                });
            }
            "vertical_alignment" => {
                let variant = keyword(&key, "vertical_alignment", V_ALIGN)?;
                setters.push(quote! {
                    .vertical_alignment(::excelreader::VerticalAlignment::#variant)
                });
            }
            "indent" => {
                let (value, span) = integer(&key)?;
                if !(0..=15).contains(&value) {
                    return Err(syn::Error::new(span, "indent must be between 0 and 15"));
                }
                setters.push(quote! { .indent(#value) });
            }
            "rotation" => {
                let (value, span) = integer(&key)?;
                if !(-90..=90).contains(&value) && value != 255 {
                    return Err(syn::Error::new(
                        span,
                        "rotation must be between -90 and 90, or 255 for stacked text",
                    ));
                }
                setters.push(quote! { .rotation(#value) });
            }
            "border" => border = Some(keyword(&key, "border", BORDER)?),
            "border_color" => border_color = Some((color(&key)?, key.path.span())),
            _ => return Err(key.error(UNKNOWN_STYLE_KEY)),
        }
        Ok(())
    })?;
    match (border, border_color) {
        (Some(style), color) => {
            let color = color.map_or(-1, |(color, _)| color);
            setters.push(quote! {
                .border(::excelreader::Border::new(::excelreader::BorderStyle::#style, #color))
            });
        }
        (None, Some((_, span))) => {
            return Err(syn::Error::new(span, "border_color needs a `border` style"));
        }
        (None, None) => {}
    }
    Ok(quote! { ::excelreader::CellStyle::new() #(#setters)* })
}

const UNKNOWN_STYLE_KEY: &str = "unsupported style(...) key, expected one of `bold`, `italic`, \
     `underline`, `strikethrough`, `number_format`, `font_name`, `font_size`, `font_color`, \
     `fill_color`, `horizontal_alignment`, `vertical_alignment`, `wrap_text`, `indent`, `rotation`, \
     `border`, `border_color`";

const H_ALIGN: &[(&str, &str)] = &[
    ("general", "General"),
    ("left", "Left"),
    ("center", "Center"),
    ("right", "Right"),
    ("fill", "Fill"),
    ("justify", "Justify"),
    ("center_continuous", "CenterContinuous"),
    ("distributed", "Distributed"),
];

const V_ALIGN: &[(&str, &str)] = &[
    ("bottom", "Bottom"),
    ("top", "Top"),
    ("center", "Center"),
    ("justify", "Justify"),
    ("distributed", "Distributed"),
];

const BORDER: &[(&str, &str)] = &[
    ("none", "None"),
    ("thin", "Thin"),
    ("medium", "Medium"),
    ("dashed", "Dashed"),
    ("dotted", "Dotted"),
    ("thick", "Thick"),
    ("double", "Double"),
    ("hair", "Hair"),
    ("medium_dashed", "MediumDashed"),
    ("dash_dot", "DashDot"),
    ("medium_dash_dot", "MediumDashDot"),
    ("dash_dot_dot", "DashDotDot"),
    ("medium_dash_dot_dot", "MediumDashDotDot"),
    ("slant_dash_dot", "SlantDashDot"),
];

/// The enum variant named by a string value, from `(value, Variant)` pairs.
fn keyword(
    key: &syn::meta::ParseNestedMeta<'_>,
    name: &str,
    choices: &[(&str, &str)],
) -> syn::Result<syn::Ident> {
    let value: LitStr = key.value()?.parse()?;
    let text = value.value();
    let Some((_, variant)) = choices.iter().find(|(choice, _)| *choice == text) else {
        let expected: Vec<String> = choices.iter().map(|(c, _)| format!("\"{c}\"")).collect();
        return Err(syn::Error::new(
            value.span(),
            format!(
                "unknown {name} {text:?}, expected one of {}",
                expected.join(", ")
            ),
        ));
    };
    Ok(syn::Ident::new(variant, value.span()))
}

/// An integer literal, optionally negated.
fn integer(key: &syn::meta::ParseNestedMeta<'_>) -> syn::Result<(i32, proc_macro2::Span)> {
    let input = key.value()?;
    let negative = input.parse::<Option<syn::Token![-]>>()?.is_some();
    let literal: syn::LitInt = input.parse()?;
    let value = literal
        .base10_parse::<i64>()
        .map(|v| if negative { -v } else { v })?;
    let value = i32::try_from(value)
        .map_err(|_| syn::Error::new(literal.span(), "value does not fit an i32"))?;
    Ok((value, literal.span()))
}

/// `-1`, or `0xRRGGBB`.
fn color(key: &syn::meta::ParseNestedMeta<'_>) -> syn::Result<i32> {
    let (value, span) = integer(key)?;
    if value != -1 && !(0..=0xFF_FFFF).contains(&value) {
        return Err(syn::Error::new(
            span,
            "a color must be 0xRRGGBB (0 to 0xFFFFFF), or -1 for unset",
        ));
    }
    Ok(value)
}

/// An integer or float literal, as `f64`.
fn number(key: &syn::meta::ParseNestedMeta<'_>) -> syn::Result<(f64, proc_macro2::Span)> {
    let literal: syn::Lit = key.value()?.parse()?;
    let value = match &literal {
        syn::Lit::Float(value) => value.base10_parse::<f64>()?,
        syn::Lit::Int(value) => value.base10_parse::<f64>()?,
        _ => return Err(syn::Error::new(literal.span(), "expected a number")),
    };
    Ok((value, literal.span()))
}

/// Returns `(T, true)` for a field typed `Option<T>`, `(ty, false)` for anything else.
fn unwrap_option(ty: &Type) -> (&Type, bool) {
    let Type::Path(type_path) = ty else {
        return (ty, false);
    };
    let Some(segment) = type_path.path.segments.last() else {
        return (ty, false);
    };
    if segment.ident != "Option" {
        return (ty, false);
    }
    let PathArguments::AngleBracketed(args) = &segment.arguments else {
        return (ty, false);
    };
    if let Some(GenericArgument::Type(inner)) = args.args.first() {
        return (inner, true);
    }
    (ty, false)
}

#[derive(Clone, Copy)]
enum FieldKind {
    Str,
    /// `i64` itself, and every other integer width. The column is always `XL_T_I64` on the wire; a
    /// narrower field converts through `TryFrom` (see `value_tokens`).
    Int,
    /// `f64` itself, plus `f32`. The column is always `XL_T_F64` on the wire.
    Float,
    Bool,
    Date,
    Time,
    Timestamp,
}

struct WriteField {
    ident: syn::Ident,
    /// The PRIMARY #[excel(name = "...")] only. An alias must never reach a write spec: the ABI
    /// rejects a write spec carrying more than one name.
    name: LitStr,
    style: Option<proc_macro2::TokenStream>,
    kind: FieldKind,
    is_option: bool,
    values: syn::Ident,
    offsets: syn::Ident,
    data: syn::Ident,
    validity: syn::Ident,
}

impl WriteField {
    fn new(index: usize, parsed: &ParsedField<'_>) -> WriteField {
        WriteField {
            ident: parsed.ident.clone(),
            name: parsed
                .names
                .first()
                .expect("excel_names always returns the primary name first")
                .clone(),
            style: parsed.style.clone(),
            kind: parsed.kind,
            is_option: parsed.is_option,
            values: quote::format_ident!("values_{index}"),
            offsets: quote::format_ident!("offsets_{index}"),
            data: quote::format_ident!("data_{index}"),
            validity: quote::format_ident!("validity_{index}"),
        }
    }

    fn declarations(&self) -> proc_macro2::TokenStream {
        let (values, offsets, data, validity) =
            (&self.values, &self.offsets, &self.data, &self.validity);
        let buffers = if matches!(self.kind, FieldKind::Str) {
            quote! {
                let mut #offsets: ::std::vec::Vec<i32> = ::std::vec::Vec::with_capacity(n + 1);
                #offsets.push(0);
                let mut #data: ::std::vec::Vec<u8> = ::std::vec::Vec::new();
            }
        } else {
            let element = self.kind.element_type();
            quote! {
                let mut #values: ::std::vec::Vec<#element> = ::std::vec::Vec::with_capacity(n);
            }
        };
        if !self.is_option {
            return buffers;
        }
        // ponytail: the bitmap is materialized for every Option<_> field, even one where no row
        quote! {
            #buffers
            let mut #validity: ::std::vec::Vec<u8> = ::std::vec![0u8; n.div_ceil(8)];
        }
    }

    fn push(&self) -> proc_macro2::TokenStream {
        let ident = &self.ident;
        let append = self.append_tokens();
        let placeholder = self.placeholder_tokens();
        if !self.is_option {
            return quote! {
                {
                    let value = &r.#ident;
                    #append
                }
            };
        }
        let validity = &self.validity;
        quote! {
            match &r.#ident {
                ::core::option::Option::Some(value) => {
                    ::excelreader::__private::set_valid(&mut #validity, row);
                    #append
                }
                ::core::option::Option::None => { #placeholder }
            }
        }
    }

    /// Appends `value: &Inner` to this field's buffers. The conversions are the exact inverses of
    /// `FieldKind::value_tokens` on the read side - if one gains a type, so must the other.
    fn append_tokens(&self) -> proc_macro2::TokenStream {
        let (values, offsets, data) = (&self.values, &self.offsets, &self.data);
        let field = self.name.value();
        match self.kind {
            FieldKind::Str => quote! {
                ::excelreader::__private::push_str(&mut #offsets, &mut #data, value.as_str())?;
            },
            FieldKind::Int => quote! {
                #values.push(
                    i64::try_from(*value).map_err(|_| ::excelreader::__private::invalid_argument(
                        ::std::format!(
                            "field `{}` holds a value that does not fit an i64 column",
                            #field
                        ),
                    ))?,
                );
            },
            FieldKind::Float => quote! { #values.push(f64::from(*value)); },
            FieldKind::Bool => quote! { #values.push(u8::from(*value)); },
            FieldKind::Date => quote! {
                #values.push(
                    ::core::convert::Into::<::excelreader::Date>::into(*value).days_since_epoch,
                );
            },
            FieldKind::Time => quote! {
                #values.push(
                    ::core::convert::Into::<::excelreader::Time>::into(*value).micros_since_midnight,
                );
            },
            FieldKind::Timestamp => quote! {
                #values.push(
                    ::core::convert::Into::<::excelreader::Timestamp>::into(*value).micros_since_epoch,
                );
            },
        }
    }

    /// A null row still occupies a slot in the values buffer; its cleared validity bit is what
    /// marks it absent. Zero (or the empty string) is the placeholder the writer never reads.
    fn placeholder_tokens(&self) -> proc_macro2::TokenStream {
        let (values, offsets, data) = (&self.values, &self.offsets, &self.data);
        if matches!(self.kind, FieldKind::Str) {
            return quote! { ::excelreader::__private::push_str(&mut #offsets, &mut #data, "")?; };
        }
        quote! { #values.push(::core::default::Default::default()); }
    }

    fn column(&self) -> proc_macro2::TokenStream {
        let (values, offsets, data, validity) =
            (&self.values, &self.offsets, &self.data, &self.validity);
        let name = &self.name;
        let payload = match self.kind {
            FieldKind::Str => quote! {
                ::excelreader::__private::OwnedColumnData::Str { offsets: #offsets, data: #data }
            },
            FieldKind::Int => quote! { ::excelreader::__private::OwnedColumnData::I64(#values) },
            FieldKind::Float => quote! { ::excelreader::__private::OwnedColumnData::F64(#values) },
            FieldKind::Bool => quote! { ::excelreader::__private::OwnedColumnData::Bool(#values) },
            FieldKind::Date => quote! { ::excelreader::__private::OwnedColumnData::Date(#values) },
            FieldKind::Time => quote! { ::excelreader::__private::OwnedColumnData::Time(#values) },
            FieldKind::Timestamp => {
                quote! { ::excelreader::__private::OwnedColumnData::Timestamp(#values) }
            }
        };
        let validity_expr = if self.is_option {
            quote! { ::core::option::Option::Some(#validity) }
        } else {
            quote! { ::core::option::Option::None }
        };
        let style_expr = match &self.style {
            Some(style) => quote! { ::core::option::Option::Some(#style) },
            None => quote! { ::core::option::Option::None },
        };
        quote! {
            ::excelreader::__private::OwnedColumn {
                name: ::core::option::Option::Some(#name),
                data: #payload,
                validity: #validity_expr,
                style: #style_expr,
            }
        }
    }
}

/// Every integer type that maps onto an `XL_T_I64` column. `i64` is included so the conversion in
/// `value_tokens` stays uniform - `TryFrom<i64> for i64` exists via the blanket `From` impl and
/// compiles away to nothing.
const INT_TYPES: &[&str] = &[
    "i8", "i16", "i32", "i64", "isize", "u8", "u16", "u32", "u64", "usize",
];

impl FieldKind {
    fn from_type(ty: &Type) -> syn::Result<FieldKind> {
        let name = match ty {
            Type::Path(type_path) => type_path
                .path
                .segments
                .last()
                .map(|segment| segment.ident.to_string()),
            _ => None,
        };
        match name.as_deref() {
            Some(int) if INT_TYPES.contains(&int) => Ok(FieldKind::Int),
            Some("String") => Ok(FieldKind::Str),
            Some("f32" | "f64") => Ok(FieldKind::Float),
            Some("bool") => Ok(FieldKind::Bool),
            Some("Date" | "NaiveDate") => Ok(FieldKind::Date),
            Some("Time" | "NaiveTime") => Ok(FieldKind::Time),
            Some("Timestamp" | "NaiveDateTime") => Ok(FieldKind::Timestamp),
            _ => Err(syn::Error::new_spanned(
                ty,
                "unsupported field type for #[derive(ExcelMapper)] - supported types are String, \
                 bool, any integer (i8..i64/isize, u8..u64/usize), f32, f64, the temporal newtypes \
                 Date/Time/Timestamp (or chrono's NaiveDate/NaiveTime/NaiveDateTime with the \
                 `chrono` feature), and Option<...> of any of those",
            )),
        }
    }

    fn column_type_tokens(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { ::excelreader::ColumnType::String },
            FieldKind::Int => quote! { ::excelreader::ColumnType::I64 },
            FieldKind::Float => quote! { ::excelreader::ColumnType::F64 },
            FieldKind::Bool => quote! { ::excelreader::ColumnType::Bool },
            FieldKind::Date => quote! { ::excelreader::ColumnType::Date },
            FieldKind::Time => quote! { ::excelreader::ColumnType::Time },
            FieldKind::Timestamp => quote! { ::excelreader::ColumnType::Timestamp },
        }
    }

    fn value_tokens(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { ::excelreader::__private::column_str(col, row).to_string() },
            FieldKind::Int => quote! {
                ::core::convert::TryFrom::try_from(::excelreader::__private::column_i64(col, row))
                    .expect("value does not fit this field's integer type")
            },
            FieldKind::Float => quote! { ::excelreader::__private::column_f64(col, row) as _ },
            FieldKind::Bool => quote! { ::excelreader::__private::column_bool(col, row) },
            FieldKind::Date => {
                quote! { ::core::convert::Into::into(::excelreader::__private::column_date(col, row)) }
            }
            FieldKind::Time => {
                quote! { ::core::convert::Into::into(::excelreader::__private::column_time(col, row)) }
            }
            FieldKind::Timestamp => {
                quote! { ::core::convert::Into::into(::excelreader::__private::column_timestamp(col, row)) }
            }
        }
    }
    /// The Rust element type of this kind's output buffer, at the column type's exact wire width.
    fn element_type(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { u8 },
            FieldKind::Int | FieldKind::Time | FieldKind::Timestamp => quote! { i64 },
            FieldKind::Float => quote! { f64 },
            FieldKind::Bool => quote! { u8 },
            FieldKind::Date => quote! { i32 },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::expand;
    use syn::DeriveInput;

    fn expand_str(src: &str) -> syn::Result<String> {
        let input: DeriveInput = syn::parse_str(src).expect("test input must parse as an item");
        expand(input).map(|tokens| tokens.to_string())
    }

    #[test]
    fn generates_one_binding_per_field_with_inferred_types() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "Nome")]
                nome: String,
                #[excel(name = "Idade")]
                idade: i64,
                #[excel(name = "Peso")]
                peso: f64,
                #[excel(name = "Ativo")]
                ativo: bool,
            }
            "#,
        )
        .expect("expand must succeed");

        assert!(output.contains("impl :: excelreader :: ExcelMapper for Row"));
        assert!(output.contains("ColumnType :: String"));
        assert!(output.contains("ColumnType :: I64"));
        assert!(output.contains("ColumnType :: F64"));
        assert!(output.contains("ColumnType :: Bool"));
        assert!(output.contains("\"Nome\""));
        assert!(output.contains("\"Idade\""));
        assert!(output.contains("\"Peso\""));
        assert!(output.contains("\"Ativo\""));
    }

    #[test]
    fn wraps_option_fields_in_some() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "Ativo")]
                ativo: Option<bool>,
            }
            "#,
        )
        .expect("expand must succeed");

        assert!(output.contains("Option :: Some"));
        assert!(output.contains("ColumnType :: Bool"));
    }

    #[test]
    fn errors_when_name_attribute_is_missing() {
        let err = expand_str(
            r#"
            struct Row {
                nome: String,
            }
            "#,
        )
        .expect_err("must fail: no #[excel(name = ...)]");

        assert!(err.to_string().contains("missing #[excel(name"));
    }

    #[test]
    fn errors_on_unsupported_field_type() {
        let err = expand_str(
            r#"
            struct Row {
                #[excel(name = "Nome")]
                nome: Vec<u8>,
            }
            "#,
        )
        .expect_err("must fail: Vec<u8> is not a supported type");

        assert!(err.to_string().contains("unsupported field type"));
    }

    #[test]
    fn maps_every_integer_width_onto_an_i64_column() {
        for ty in [
            "i8", "i16", "i32", "i64", "isize", "u8", "u16", "u32", "u64", "usize",
        ] {
            let output = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "N")] n: {ty}, }}"#
            ))
            .unwrap_or_else(|e| panic!("{ty} must be supported: {e}"));

            assert!(
                output.contains("ColumnType :: I64"),
                "{ty} must use an I64 column"
            );
            assert!(output.contains("try_from"), "{ty} must convert via TryFrom");
        }
    }

    #[test]
    fn maps_both_float_widths_onto_an_f64_column() {
        for ty in ["f32", "f64"] {
            let output = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "V")] v: {ty}, }}"#
            ))
            .unwrap_or_else(|e| panic!("{ty} must be supported: {e}"));
            assert!(
                output.contains("ColumnType :: F64"),
                "{ty} must use an F64 column"
            );
        }
    }

    #[test]
    fn maps_temporal_types_onto_their_own_column_types() {
        for (ty, expected) in [
            ("Date", "ColumnType :: Date"),
            ("NaiveDate", "ColumnType :: Date"),
            ("Time", "ColumnType :: Time"),
            ("NaiveTime", "ColumnType :: Time"),
            ("Timestamp", "ColumnType :: Timestamp"),
            ("NaiveDateTime", "ColumnType :: Timestamp"),
        ] {
            let output = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "D")] d: {ty}, }}"#
            ))
            .unwrap_or_else(|e| panic!("{ty} must be supported: {e}"));
            assert!(output.contains(expected), "{ty} must use {expected}");
        }
    }

    #[test]
    fn resolves_temporal_types_through_a_qualified_path() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "D")]
                d: chrono::NaiveDate,
            }
            "#,
        )
        .expect("a qualified path must resolve by its last segment");

        assert!(output.contains("ColumnType :: Date"));
    }

    #[test]
    fn errors_on_non_struct_input() {
        let err = expand_str(
            r#"
            enum Row {
                A,
            }
            "#,
        )
        .expect_err("must fail: enum is not a struct");

        assert!(err
            .to_string()
            .contains("can only be derived for structs with named fields"));
    }
    #[test]
    fn collects_aliases_in_declared_order_after_the_primary_name() {
        let output = expand_str(
            r#"
        struct Row {
            #[excel(name = "Nome", alias = "Nom", alias = "Name")]
            nome: String,
        }
        "#,
        )
        .expect("expand must succeed");

        let nome_pos = output.find("\"Nome\"").expect("primary name must appear");
        let nom_pos = output.find("\"Nom\"").expect("first alias must appear");
        let name_pos = output.rfind("\"Name\"").expect("second alias must appear");
        assert!(nome_pos < nom_pos, "primary name must precede its aliases");
        assert!(nom_pos < name_pos, "aliases must stay in declared order");
    }
    #[test]
    fn generates_an_excel_writer_impl_alongside_the_mapper() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "Nome", alias = "Name")]
                nome: String,
                #[excel(name = "Peso")]
                peso: Option<f64>,
            }
            "#,
        )
        .expect("expand must succeed");

        assert!(output.contains("impl :: excelreader :: ExcelWriter for Row"));
        assert!(
            output.contains("push_str"),
            "a String field must go through push_str"
        );
        assert!(
            output.contains("set_valid"),
            "an Option field must set a validity bit"
        );
        let writer_half = output
            .split("ExcelWriter for Row")
            .nth(1)
            .expect("the writer impl must be present");
        assert!(
            !writer_half.contains("\"Name\""),
            "an alias must never reach a write column"
        );
        assert!(
            writer_half.contains("\"Nome\""),
            "the primary name must reach the write column"
        );
    }

    #[test]
    fn a_field_without_style_writes_no_column_style() {
        let output = expand_str(r#"struct Row { #[excel(name = "V")] v: f64, }"#)
            .expect("expand must succeed");
        assert!(output.contains("style : :: core :: option :: Option :: None"));
    }

    #[test]
    fn style_keys_become_cell_style_setters() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "V", style(bold, italic, underline, strikethrough, wrap_text,
                    number_format = "0.00", font_name = "Arial", font_size = 12,
                    font_color = 0xFF0000, fill_color = -1,
                    horizontal_alignment = "center_continuous", vertical_alignment = "distributed", indent = 15, rotation = 255,
                    border = "slant_dash_dot", border_color = 0x0000FF))]
                v: f64,
            }
            "#,
        )
        .expect("expand must succeed");

        for expected in [
            ". bold (true)",
            ". italic (true)",
            ". underline (true)",
            ". strikethrough (true)",
            ". wrap_text (true)",
            ". number_format (\"0.00\")",
            ". font_name (\"Arial\")",
            ". font_size (12f64)",
            ". font_color (16711680i32)",
            ". fill_color (- 1i32)",
            "HorizontalAlignment :: CenterContinuous",
            "VerticalAlignment :: Distributed",
            ". indent (15i32)",
            ". rotation (255i32)",
            "BorderStyle :: SlantDashDot , 255i32",
        ] {
            assert!(
                output.contains(expected),
                "missing `{expected}` in {output}"
            );
        }
    }

    #[test]
    fn rejects_out_of_range_style_values() {
        for (style, message) in [
            ("font_color = 0x1000000", "0xRRGGBB"),
            ("fill_color = -2", "0xRRGGBB"),
            ("border = \"thin\", border_color = 16777216", "0xRRGGBB"),
            ("font_size = 410.0", "font_size"),
            ("font_size = 0.5", "font_size"),
            ("indent = 16", "indent"),
            ("rotation = 91", "rotation"),
            ("rotation = -91", "rotation"),
            ("font_name = \"\"", "font_name"),
            ("horizontal_alignment = \"middle\"", "horizontal_alignment"),
            ("vertical_alignment = \"left\"", "vertical_alignment"),
            ("border = \"thicker\"", "border"),
            ("border_color = 0xFF0000", "border_color"),
            ("colour = 1", "unsupported style"),
        ] {
            let err = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "V", style({style}))] v: f64, }}"#
            ))
            .expect_err(style);
            assert!(
                err.to_string().contains(message),
                "{style}: `{err}` should mention {message}"
            );
        }
    }
}
